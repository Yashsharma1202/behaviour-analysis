"""
finetune_sentiment.py
===============================================================================
Fine-tune FinBERT on what NSE filings were actually FOLLOWED BY, instead of on
US analyst prose.

WHAT THIS MODEL IS — AND IS NOT
-------------------------------
It is NOT a sentiment tagger. It is a FORWARD-RETURN classifier that happens to
read text. Trained on standardised abnormal return, and the measured relation
between filing tone and forward return in this corpus is INVERTED (see
eval_sentiment_car.py: bullish-reading filings are followed by -1.14% over 20
days), this model will learn "text that reads bullish -> lower future return".

So its output must never be rendered as "Sentiment: Positive". The classes are
named UP / FLAT / DOWN after the RETURN, not after the tone.

HONEST EVALUATION
-----------------
Accuracy is a near-useless metric here: the label is a tercile of a quantity
that is mostly noise, so ~40% accuracy can be excellent and ~40% can be
worthless. The metric that decides whether this ships is the one the product
cares about:

    on the HELD-OUT test era, does this model separate forward abnormal
    returns better than the router already in production?

Both are measured on the same rows. If the fine-tune does not beat the router
out of sample, it does not ship, and that is a real result.

LEAKAGE CONTROL
    * split is TEMPORAL, never random — train on the past, test on the future
    * tercile cut-points come from the TRAIN era only, then are applied forward
    * text is known at filing time; the target is strictly forward
    * no symbol appears in a way that leaks (the same symbol spans all eras,
      which is realistic — you would deploy this on stocks you had trained on)

    python finetune_sentiment.py                 # train + evaluate
    python finetune_sentiment.py --epochs 3
    python finetune_sentiment.py --eval-only     # reuse the saved model
===============================================================================
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "processed" / "car_dataset.parquet"
MODEL_DIR = ROOT / "processed" / "finbert_nse"
BASE = "ProsusAI/finbert"

# Named after the RETURN, not the tone. This is the whole point.
CLASSES = ["DOWN", "FLAT", "UP"]

TRAIN_END = 2021          # train  <= 2021   (~44k rows)
VAL_END = 2023            # val    2022-2023 (~12k)
                          # test   2024+     (~17k)


# --------------------------------------------------------------------------- #
def load_split():
    if not DATA.exists():
        raise SystemExit(f"missing {DATA} — run: python build_car_dataset.py")
    df = pd.read_parquet(DATA)
    df["year"] = df["ts"].dt.year
    tr = df[df.year <= TRAIN_END].copy()
    va = df[(df.year > TRAIN_END) & (df.year <= VAL_END)].copy()
    te = df[df.year > VAL_END].copy()

    # Tercile cut-points from TRAIN ONLY — using the full sample would leak the
    # future distribution into the label definition.
    lo, hi = tr["sar"].quantile([1 / 3, 2 / 3])
    for d in (tr, va, te):
        d["y"] = np.where(d["sar"] <= lo, 0, np.where(d["sar"] >= hi, 2, 1))
    return tr, va, te, float(lo), float(hi)


def _load_base(num_labels: int, from_dir: Path | None = None):
    """Load FinBERT, working around both quirks of these old checkpoints."""
    import torch.nn as nn
    from transformers import (AutoTokenizer, BertConfig,
                              BertForSequenceClassification, BertTokenizerFast)

    src = str(from_dir) if from_dir else BASE
    try:
        tok = AutoTokenizer.from_pretrained(src)
    except (ValueError, OSError):
        tok = BertTokenizerFast.from_pretrained(src)

    if from_dir:
        net = BertForSequenceClassification.from_pretrained(src)
        return tok, net

    import json as _json
    from huggingface_hub import hf_hub_download
    raw = _json.load(open(hf_hub_download(BASE, "config.json"), encoding="utf-8"))
    raw.setdefault("model_type", "bert")
    raw["num_labels"] = num_labels
    raw["id2label"] = {str(i): c for i, c in enumerate(CLASSES)}
    raw["label2id"] = {c: i for i, c in enumerate(CLASSES)}
    cfg = BertConfig.from_dict(raw)
    # ignore_mismatched_sizes lets num_labels differ from the checkpoint's 3
    # (e.g. a binary UP/DOWN variant) — the head is re-initialised below anyway.
    net = BertForSequenceClassification.from_pretrained(
        BASE, config=cfg, ignore_mismatched_sizes=True)
    # The pretrained head predicts pos/neg/neutral TONE. We are predicting
    # forward RETURN. Same shape, different meaning — reset it so the old
    # head's priors do not bias the new task.
    net.classifier = nn.Linear(cfg.hidden_size, num_labels)
    return tok, net


def encode(tok, texts, max_len):
    return tok(list(texts), padding="max_length", truncation=True,
               max_length=max_len, return_tensors="pt")


def train(tr, va, epochs, max_len, bs, lr):
    import torch
    from torch.utils.data import DataLoader, TensorDataset

    dev = "cuda" if torch.cuda.is_available() else "cpu"
    tok, net = _load_base(len(CLASSES))
    net.to(dev)

    print(f"  device {dev} | train {len(tr):,} | val {len(va):,} | "
          f"max_len {max_len} | batch {bs} | lr {lr}")

    enc = encode(tok, tr["text"].tolist(), max_len)
    ds = TensorDataset(enc["input_ids"], enc["attention_mask"],
                       torch.tensor(tr["y"].to_numpy(), dtype=torch.long))
    dl = DataLoader(ds, batch_size=bs, shuffle=True, drop_last=True)

    venc = encode(tok, va["text"].tolist(), max_len)
    vds = TensorDataset(venc["input_ids"], venc["attention_mask"],
                        torch.tensor(va["y"].to_numpy(), dtype=torch.long))
    vdl = DataLoader(vds, batch_size=128)

    opt = torch.optim.AdamW(net.parameters(), lr=lr, weight_decay=0.01)
    steps = len(dl) * epochs
    sched = torch.optim.lr_scheduler.OneCycleLR(
        opt, max_lr=lr, total_steps=steps, pct_start=0.1, anneal_strategy="linear")
    scaler = torch.amp.GradScaler(dev) if dev == "cuda" else None

    best, best_state = -1e9, None
    for ep in range(1, epochs + 1):
        net.train()
        t0, tot, n = time.time(), 0.0, 0
        for i, (ids, am, y) in enumerate(dl, 1):
            ids, am, y = ids.to(dev), am.to(dev), y.to(dev)
            opt.zero_grad(set_to_none=True)
            if scaler:
                with torch.autocast("cuda", dtype=torch.bfloat16):
                    out = net(input_ids=ids, attention_mask=am, labels=y)
                out.loss.backward()
            else:
                out = net(input_ids=ids, attention_mask=am, labels=y)
                out.loss.backward()
            torch.nn.utils.clip_grad_norm_(net.parameters(), 1.0)
            opt.step()
            sched.step()
            tot += float(out.loss); n += 1
            if i % 200 == 0:
                print(f"    ep{ep} step {i}/{len(dl)}  loss {tot / n:.4f}  "
                      f"{time.time() - t0:.0f}s")

        # Validation is scored on RANK CORRELATION with sar, not accuracy —
        # a model that orders outcomes correctly is what we want, even if it
        # rarely nails the tercile.
        p = predict(net, tok, vdl, dev)
        score = float(pd.Series(p @ np.array([-1.0, 0.0, 1.0])).corr(
            pd.Series(va["sar"].to_numpy()), method="spearman"))
        print(f"    epoch {ep}: train loss {tot / n:.4f}  "
              f"val rank-corr(sar) {score:+.4f}")
        if score > best:
            best = score
            best_state = {k: v.detach().cpu().clone() for k, v in net.state_dict().items()}

    if best_state:
        net.load_state_dict(best_state)
    MODEL_DIR.mkdir(parents=True, exist_ok=True)
    net.save_pretrained(MODEL_DIR)
    tok.save_pretrained(MODEL_DIR)
    print(f"  best val rank-corr {best:+.4f}  ->  {MODEL_DIR}")
    return net, tok


def predict(net, tok, dl, dev) -> np.ndarray:
    import torch
    net.eval()
    out = []
    with torch.no_grad():
        for batch in dl:
            ids, am = batch[0].to(dev), batch[1].to(dev)
            logits = net(input_ids=ids, attention_mask=am).logits.float()
            out.append(torch.softmax(logits, -1).cpu().numpy())
    return np.concatenate(out)


# --------------------------------------------------------------------------- #
# The metric that decides whether this ships
# --------------------------------------------------------------------------- #
def car_spread(df, col) -> tuple[float, int, int]:
    up = df.loc[df[col] == "UP", "car"]
    dn = df.loc[df[col] == "DOWN", "car"]
    if len(up) < 20 or len(dn) < 20:
        return float("nan"), len(up), len(dn)
    return float(up.mean() - dn.mean()), len(up), len(dn)


def perm_p(df, col, obs, perms, rng) -> float:
    if not np.isfinite(obs):
        return float("nan")
    groups = [(g["car"].to_numpy(), g[col].to_numpy())
              for _s, g in df.groupby("symbol", sort=False)]
    hits = 0
    for _ in range(perms):
        u, d = [], []
        for c, l in groups:
            sh = rng.permutation(l)
            u.append(c[sh == "UP"]); d.append(c[sh == "DOWN"])
        u, d = np.concatenate(u), np.concatenate(d)
        if len(u) < 20 or len(d) < 20:
            continue
        if abs(u.mean() - d.mean()) >= abs(obs):
            hits += 1
    return (hits + 1) / (perms + 1)


def evaluate(te, net, tok, max_len, perms):
    """Head-to-head on the held-out era: fine-tune vs the production router."""
    import torch
    from torch.utils.data import DataLoader, TensorDataset
    import sentiment_finbert as SF

    dev = "cuda" if torch.cuda.is_available() else "cpu"
    net.to(dev)
    enc = encode(tok, te["text"].tolist(), max_len)
    dl = DataLoader(TensorDataset(enc["input_ids"], enc["attention_mask"]),
                    batch_size=128)
    p = predict(net, tok, dl, dev)
    te = te.copy()
    te["ft"] = [CLASSES[i] for i in p.argmax(1)]
    te["ft_score"] = p @ np.array([-1.0, 0.0, 1.0])

    # The router, on the identical rows. Its Positive/Negative are TONE, so map
    # them onto the UP/DOWN axis to compare like with like.
    router = SF.score_many(te["text"].tolist(), te["category"].tolist())
    te["router"] = [{"Positive": "UP", "Negative": "DOWN"}.get(r.label, "FLAT")
                    for r in router]

    rng = np.random.default_rng(42)
    print("\n" + "=" * 90)
    print(f"  HELD-OUT TEST ERA  {te['ts'].min().date()} .. {te['ts'].max().date()}"
          f"   ({len(te):,} filings)")
    print("=" * 90)
    print(f"  {'model':<28}{'UP-DOWN spread':>16}{'perm p':>10}"
          f"{'n UP':>9}{'n DOWN':>9}")
    print("  " + "-" * 86)
    res = {}
    for col, name in (("router", "router (in production)"),
                      ("ft", "fine-tuned on CAR")):
        s, nu, nd = car_spread(te, col)
        if not np.isfinite(s):
            print(f"  {name:<28}{'n/a':>16}   (up={nu}, down={nd})")
            res[col] = None
            continue
        pp = perm_p(te, col, s, perms, rng)
        star = "***" if pp < 0.01 else ("**" if pp < 0.05 else ("*" if pp < 0.10 else ""))
        print(f"  {name:<28}{s:>+15.3f}%{pp:>10.4f} {star:<3}{nu:>7,}{nd:>9,}")
        res[col] = {"spread": s, "p": pp, "n_up": nu, "n_down": nd}

    ic = float(pd.Series(te["ft_score"]).corr(pd.Series(te["sar"]), method="spearman"))
    acc = float((te["ft"] == [CLASSES[i] for i in te["y"]]).mean())
    print(f"\n  fine-tune rank-corr with sar: {ic:+.4f}   tercile accuracy: {acc:.1%} "
          f"(chance 33.3%)")
    print("  Accuracy is not the criterion — the spread row above is.")
    return te, res


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--epochs", type=int, default=2)
    ap.add_argument("--max-len", type=int, default=256)
    ap.add_argument("--batch", type=int, default=32)
    ap.add_argument("--lr", type=float, default=2e-5)
    ap.add_argument("--perms", type=int, default=2000)
    ap.add_argument("--eval-only", action="store_true")
    args = ap.parse_args()

    tr, va, te, lo, hi = load_split()
    print("=" * 90)
    print("  FINE-TUNE FinBERT ON FORWARD ABNORMAL RETURN")
    print("=" * 90)
    print(f"  train  <= {TRAIN_END}        {len(tr):>7,} filings")
    print(f"  val    {TRAIN_END + 1}-{VAL_END}       {len(va):>7,}")
    print(f"  test   {VAL_END + 1}+          {len(te):>7,}   (never seen in training)")
    print(f"  tercile cuts from TRAIN only: sar <= {lo:+.3f} = DOWN, "
          f">= {hi:+.3f} = UP")
    print(f"  class balance (train): "
          f"{np.bincount(tr['y'], minlength=3) / len(tr)}")

    if args.eval_only:
        if not MODEL_DIR.exists():
            raise SystemExit(f"no saved model at {MODEL_DIR}")
        tok, net = _load_base(len(CLASSES), from_dir=MODEL_DIR)
    else:
        net, tok = train(tr, va, args.epochs, args.max_len, args.batch, args.lr)

    te_out, res = evaluate(te, net, tok, args.max_len, args.perms)

    rt = res.get("router") or {}
    ft = res.get("ft") or {}
    print("\n" + "=" * 90)
    if ft and rt and abs(ft["spread"]) > abs(rt["spread"]) and ft["p"] < 0.05:
        print(f"  SHIPS — fine-tune separates outcomes better out of sample "
              f"({ft['spread']:+.3f}% vs {rt['spread']:+.3f}%)")
    else:
        print("  DOES NOT SHIP — the fine-tune does not beat the production router")
        print("  out of sample. Keep the router. This is a real result, not a bug.")
    print("=" * 90)

    (ROOT / "processed" / "finetune_result.json").write_text(
        json.dumps({"router": rt, "finetune": ft,
                    "test_from": str(te_out["ts"].min().date()),
                    "test_to": str(te_out["ts"].max().date()),
                    "n_test": int(len(te_out))}, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()

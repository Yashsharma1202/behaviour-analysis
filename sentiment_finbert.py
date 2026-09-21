"""
sentiment_finbert.py
===============================================================================
FinBERT-backed sentiment for NSE corporate filings — a THREE-STAGE ROUTER, not
a bare model call.

WHY A ROUTER AND NOT JUST FinBERT
---------------------------------
FinBERT is trained on US analyst prose. NSE announcement text is not that: it is
dominated by procedural boilerplate that carries no information about the
business, and every row is wrapped in regulatory preamble. Measured over the
106,925 announcement rows on disk:

    Loss of Share Certificates ....  9,397 rows (8.8%)  identical boilerplate
    Copy of Newspaper Publication .  2,755 rows         identical boilerplate
    Trading Window ................  1,759 rows         identical boilerplate
    ...

Handing those to a transformer is both wasteful and wrong — FinBERT reads "loss"
and returns Negative. sentiment_analyzer.py already encodes the correct answer
for exactly these cases (its 0.0-weight phrase list). So:

    STAGE 1  boilerplate gate   category/text is procedural   -> Neutral, no model
    STAGE 2  FinBERT            everything else               -> model verdict
    STAGE 3  lexicon override   known-certain phrases         -> rules beat model

Stage 3 runs LAST on purpose. A model that says "Positive" about a show-cause
notice is wrong in a way we can prove, so the rule wins.

TEXT CLEANING
-------------
Nearly every filing opens with "Pursuant to Regulation 39(3) of the Securities
and Exchange Board of India (Listing Obligations and Disclosure Requirements)
Regulations, 2015, ..." — 200+ characters of legal furniture ahead of the actual
news. strip_preamble() removes it. Without this the model spends most of its
512-token budget on text every filing shares.

USE
    from sentiment_finbert import score_text, score_many
    r = score_text("Company secured an order worth Rs 500 crore.", "Updates")
    r.label, r.score, r.confidence, r.engine

The model is loaded lazily and only once. If transformers/torch are missing or
the model cannot be fetched, every call transparently falls back to the lexicon
in sentiment_analyzer.py — the dashboard keeps working, it just gets weaker tags.
===============================================================================
"""

from __future__ import annotations

import os
import re
import sys
import threading
from dataclasses import dataclass, asdict

from sentiment_analyzer import LEXICON, analyze_sentiment

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
# ProsusAI/finbert, not yiyanghkust/finbert-tone. Measured on the 200-row
# benchmark (test_sentiment_finbert.py): tone scores 0.321 macro-F1, Prosus
# 0.439. Tone collapses to Neutral on exchange-speak — it was tuned on analyst
# prose with explicit tone words, which NSE filings do not use.
MODEL_NAME = os.environ.get("FINBERT_MODEL", "ProsusAI/finbert")
MAX_TOKENS = 512
BATCH_SIZE = int(os.environ.get("FINBERT_BATCH", "256"))


@dataclass(frozen=True)
class SentimentResult:
    label: str          # Positive | Negative | Neutral
    score: float        # -1.0 .. +1.0
    confidence: float   # 0.0 .. 1.0  (softmax max for finbert; 1.0 for rules)
    engine: str         # rule | finbert | override | lexicon

    def as_cell(self) -> str:
        """The 'Label|score' string the dashboard's sentiment column renders."""
        return f"{self.label}|{self.score}"

    def to_dict(self) -> dict:
        return asdict(self)


NEUTRAL = SentimentResult("Neutral", 0.0, 1.0, "rule")


# --------------------------------------------------------------------------- #
# STAGE 1 — boilerplate gate
# --------------------------------------------------------------------------- #
# Categories (the `desc` column) that are procedural by definition. Kept
# deliberately tight: only filings that are NEVER informative about the business
# belong here. "Investor Presentation" and "Outcome of Board Meeting" do NOT —
# they carry real content and must reach the model.
BOILERPLATE_CATEGORIES = {
    "loss of share certificates",
    "loss of share certificate",
    "loss of certificate",
    "loss of certificates",
    "copy of newspaper publication",
    "newspaper publication",
    "newspaper advertisements",
    "newspaper advertisement",
    "trading window",
    "certificate under sebi (depositories and participants) regulations, 2018",
    "disclosure of voting pattern - clause 35a",
    "book closure",
    "agm/book closure",
    "record date",
    "postal ballot",
    "reg. 39(3) - loss of share certificates",
    "share certificate",
    "duplicate share certificate",
    "duplicate share certificates",
}

# Text-level boilerplate: procedural filings whose category is generic
# ("Updates", "General Updates") but whose body is unmistakably routine.
BOILERPLATE_TEXT = re.compile(
    r"loss of (share )?certificate"
    r"|duplicate share certificate"
    r"|issue of duplicate"
    r"|trading window (closure|shall remain closed|will remain closed)"
    r"|copy of (the )?newspaper (publication|advertisement|clipping)"
    r"|newspaper (publication|advertisement|clipping)s? (is|are|being) (attached|enclosed)"
    r"|submission of (the )?annual (secretarial )?compliance report"
    r"|certificate under regulation 7\(3\)"
    r"|certificate under regulation 40\(9\)"
    r"|reconciliation of share capital audit"
    r"|investor (grievance|complaint)s? (report|redressal)"
    r"|disclosure of voting (pattern|results)"
    r"|scrutinizer'?s report",
    re.I,
)

# Regulatory furniture stripped before the model sees the text. Every one of
# these appears verbatim on thousands of filings, so leaving them in makes all
# filings look alike to the model.
_PREAMBLE = re.compile(
    r"^\s*(pursuant to|in terms of|in compliance with|with reference to|"
    r"further to|in continuation (of|to)|this is (further )?to inform|"
    r"we (hereby )?(wish to )?inform|please find attached|"
    r"we (are )?(hereby )?enclos\w*|disclosure under)\b"
    r"[^.]{0,400}?"
    r"(regulation|regulations|circular|clause|act|sebi|listing obligations)"
    r"[^.]{0,200}?[.,;:\-]\s*",
    re.I,
)
_REG_CITATION = re.compile(
    r"(regulation|reg\.?)\s*\d+[A-Za-z]?(\s*\(\d+\))*\s*"
    r"(of|under)?\s*(the\s*)?(securities and exchange board of india|sebi)?"
    r"[^.]{0,160}?regulations?,?\s*\d{4}",
    re.I,
)
_SEBI_CIRCULAR = re.compile(r"sebi\s*/?\s*circular[^.]{0,160}", re.I)
_WS = re.compile(r"\s+")


def strip_preamble(text: str) -> str:
    """Drop the regulatory furniture so the model reads the actual news."""
    if not text:
        return ""
    t = _PREAMBLE.sub("", text, count=1)
    t = _REG_CITATION.sub(" ", t)
    t = _SEBI_CIRCULAR.sub(" ", t)
    t = _WS.sub(" ", t).strip(" ,.;:-")
    # If stripping ate everything meaningful, fall back to the original.
    return t if len(t) >= 15 else _WS.sub(" ", text).strip()


def is_boilerplate(category: str, text: str) -> bool:
    """STAGE 1 — True when the filing is procedural and carries no signal."""
    if (category or "").strip().lower() in BOILERPLATE_CATEGORIES:
        return True
    blob = f"{category} {text}"
    if BOILERPLATE_TEXT.search(blob):
        return True
    return False


# --------------------------------------------------------------------------- #
# STAGE 2a — the numeric results parser
# --------------------------------------------------------------------------- #
# NSE "Results Update" filings are machine-generated from a fixed template:
#
#   "... Net Profit / (Loss) of Rs. 98957 lacs for the year ending on
#    31-MAR-2011 against Rs. 108259 lacs for the year ending on 31-MAR-2010."
#
# The direction of that filing is ARITHMETIC, not tone. No 3-class sentiment
# model can compare two integers, and the benchmark proves it: FinBERT scores
# these at chance. 38 of the 200 benchmark rows are this shape. So parse them.
#
# Parentheses mean a loss: "Rs. (91724) lacs" is -91,724.
_PAT_BLOCK = re.compile(
    r"net\s*profit\s*/?\s*\(?\s*loss\s*\)?\s*(?:of)?\s*(?:of)?\s*"
    r"(?P<body>.{0,400})", re.I | re.S,
)
_MONEY = re.compile(r"rs\.?\s*(?P<neg>\()?\s*(?P<n>[\d,]+(?:\.\d+)?)\s*\)?", re.I)

# Below this, a move is noise rather than news (rounding, tiny base effects).
RESULT_MOVE_CUT = 0.02


def _money(m: re.Match) -> float:
    v = float(m.group("n").replace(",", ""))
    return -v if m.group("neg") else v


def parse_result_move(text: str):
    """(score, label) for a templated results filing, or None if not one.

    Returns Neutral when the filing quotes no prior period (nothing to compare)
    or when it quotes several that DISAGREE — a quarter that beat last year but
    missed last quarter is genuinely mixed, and saying so is more honest than
    picking whichever comparison the template happened to print first.
    """
    if not text:
        return None
    blk = _PAT_BLOCK.search(text)
    if not blk:
        return None
    body = blk.group("body")
    # "against" separates the current figure from the prior ones.
    head, sep, tail = body.partition("against")
    cur_m = _MONEY.search(head)
    if not cur_m:
        return None
    if not sep:
        return 0.0, "Neutral"                     # filed, but nothing to compare
    cur = _money(cur_m)
    priors = [_money(m) for m in _MONEY.finditer(tail)][:3]
    priors = [p for p in priors if p != 0]
    if not priors:
        return 0.0, "Neutral"

    dirs = []
    for p in priors:
        chg = (cur - p) / abs(p)
        dirs.append(0 if abs(chg) < RESULT_MOVE_CUT else (1 if chg > 0 else -1))
    if all(d > 0 for d in dirs):
        worst = min((cur - p) / abs(p) for p in priors)
        return round(min(1.0, max(0.15, worst)), 3), "Positive"
    if all(d < 0 for d in dirs):
        best = max((cur - p) / abs(p) for p in priors)
        return round(max(-1.0, min(-0.15, best)), 3), "Negative"
    return 0.0, "Neutral"                         # mixed or below the noise cut


# --------------------------------------------------------------------------- #
# STAGE 3 — lexicon overrides (rules that beat the model)
# --------------------------------------------------------------------------- #
# Phrases the existing lexicon pins to exactly 0.0 — hard-won knowledge about
# Indian filing language that no US-trained model has.
FORCE_NEUTRAL = tuple(sorted(
    (p for p, w in LEXICON.items() if w == 0.0 and " " in p),
    key=len, reverse=True,
))

# Filings whose direction is not a matter of opinion. If the model disagrees
# with one of these, the model is wrong.
FORCE_NEGATIVE = re.compile(
    r"\b(insolvency|bankruptcy|liquidation|winding[- ]up petition"
    r"|show[- ]cause notice|defaulted?|wilful defaulter"
    r"|fraud|misappropriation|embezzl\w+"
    r"|search (and seizure )?operation|raid(ed|s)? by"
    r"|resignation of (the )?(statutory )?auditor|auditor has resigned"
    r"|qualified opinion|adverse opinion|disclaimer of opinion"
    r"|suspension of trading|delisting notice"
    r"|forensic audit|class action)\b",
    re.I,
)
FORCE_POSITIVE = re.compile(
    r"\b(bonus issue|stock split"
    r"|(secured|received|bagged|awarded|won)\s+(an?\s+)?"
    r"(new\s+)?(order|contract|mandate|tender|loa|letter of award)"
    r"|order (win|inflow)s?"
    r"|commercial production commenced|commenced commercial production)\b",
    re.I,
)


def _apply_overrides(text: str, base: SentimentResult) -> SentimentResult:
    """STAGE 3 — deterministic rules take precedence over the model."""
    low = text.lower()
    for phrase in FORCE_NEUTRAL:
        if phrase in low:
            return SentimentResult("Neutral", 0.0, 1.0, "override")
    if FORCE_NEGATIVE.search(text):
        if base.label != "Negative":
            return SentimentResult("Negative", min(base.score, -0.6), 1.0, "override")
        return base
    if FORCE_POSITIVE.search(text) and base.label == "Negative":
        # The model mistook an order win for bad news (it reads "loss"/"claim"
        # elsewhere in the sentence). Rules win.
        return SentimentResult("Positive", 0.5, 1.0, "override")
    return base


# --------------------------------------------------------------------------- #
# STAGE 2 — FinBERT
# --------------------------------------------------------------------------- #
_MODEL = {"tok": None, "net": None, "dev": None, "order": None, "failed": False}
_LOAD_LOCK = threading.Lock()


def _load_model():
    """Lazily load FinBERT once. Returns False if unavailable (caller falls back)."""
    if _MODEL["failed"]:
        return False
    if _MODEL["net"] is not None:
        return True
    with _LOAD_LOCK:
        if _MODEL["net"] is not None:
            return True
        if _MODEL["failed"]:
            return False
        try:
            import torch
            from transformers import AutoModelForSequenceClassification, AutoTokenizer

            dev = "cuda" if torch.cuda.is_available() else "cpu"
            # Both FinBERT checkpoints predate tokenizer.json and ship only a
            # WordPiece vocab.txt. transformers 5.x dropped slow tokenizers, and
            # AutoTokenizer's conversion path demands sentencepiece/tiktoken that
            # a WordPiece model does not need. BertTokenizerFast reads vocab.txt
            # directly, so try Auto first and fall back to it.
            try:
                tok = AutoTokenizer.from_pretrained(MODEL_NAME)
            except (ValueError, OSError):
                from transformers import BertTokenizerFast
                tok = BertTokenizerFast.from_pretrained(MODEL_NAME)

            # yiyanghkust/finbert-tone was exported before `model_type` became
            # mandatory, so AutoModel cannot infer the architecture. Both FinBERT
            # checkpoints are BertForSequenceClassification — rebuild the config
            # with the missing key and name the class directly when Auto fails.
            try:
                net = AutoModelForSequenceClassification.from_pretrained(MODEL_NAME)
            except (ValueError, KeyError):
                import json as _json

                from huggingface_hub import hf_hub_download
                from transformers import BertConfig, BertForSequenceClassification

                raw = _json.load(open(hf_hub_download(MODEL_NAME, "config.json"),
                                      encoding="utf-8"))
                raw.setdefault("model_type", "bert")
                cfg = BertConfig.from_dict(raw)
                net = BertForSequenceClassification.from_pretrained(MODEL_NAME, config=cfg)
            net.eval().to(dev)
            if dev == "cuda":
                net.half()                       # fp16 — Blackwell handles this natively

            # Map the model's own id2label to our three slots. finbert-tone uses
            # Neutral/Positive/Negative; ProsusAI/finbert uses positive/negative/
            # neutral. Never assume an index order — read it off the config.
            id2label = {int(k): str(v).lower() for k, v in net.config.id2label.items()}
            order = {}
            for idx, name in id2label.items():
                if name.startswith("pos"):
                    order["pos"] = idx
                elif name.startswith("neg"):
                    order["neg"] = idx
                else:
                    order["neu"] = idx
            if len(order) != 3:
                raise ValueError(f"unexpected label map {id2label}")

            _MODEL.update(tok=tok, net=net, dev=dev, order=order)
            return True
        except Exception as e:                   # noqa: BLE001
            print(f"  ! FinBERT unavailable ({type(e).__name__}: {e}) "
                  f"-> falling back to the lexicon", file=sys.stderr)
            _MODEL["failed"] = True
            return False


def model_info() -> dict:
    """What actually got loaded — for the benchmark banner and the UI footer."""
    if not _load_model():
        return {"available": False, "model": MODEL_NAME, "device": None}
    return {"available": True, "model": MODEL_NAME, "device": _MODEL["dev"],
            "labels": _MODEL["order"]}


def _finbert_batch(texts: list[str]) -> list[tuple[float, float, str]]:
    """(score, confidence, label) per text.

    label is the model's ARGMAX class — its own decision rule, with no free
    parameter. An earlier version thresholded P(pos)-P(neg) at +/-0.05 instead;
    on the benchmark that scored 0.592 macro-F1 against argmax's 0.685, and
    false-flagged 53 procedural filings against argmax's 23. A hand-picked cut
    was both worse and one more knob to overfit.

    score stays continuous (P(pos) - P(neg)) because Phase 5 wants a signed
    intensity as an ML feature, not just a three-way label.
    """
    import torch

    tok, net, dev, order = _MODEL["tok"], _MODEL["net"], _MODEL["dev"], _MODEL["order"]
    idx2name = {order["pos"]: "Positive", order["neg"]: "Negative",
                order["neu"]: "Neutral"}
    out: list[tuple[float, float, str]] = []
    with torch.no_grad():
        for i in range(0, len(texts), BATCH_SIZE):
            chunk = texts[i:i + BATCH_SIZE]
            enc = tok(chunk, padding=True, truncation=True,
                      max_length=MAX_TOKENS, return_tensors="pt").to(dev)
            logits = net(**enc).logits.float()
            probs = torch.softmax(logits, dim=-1)
            score = (probs[:, order["pos"]] - probs[:, order["neg"]]).cpu().tolist()
            top = probs.max(dim=-1)
            conf = top.values.cpu().tolist()
            labs = [idx2name[int(j)] for j in top.indices.cpu().tolist()]
            out.extend(zip(score, conf, labs))
    return out


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #
def score_many(texts, categories=None) -> list[SentimentResult]:
    """Route a list of filings through all three stages. Batched — this is the
    entry point score_feeds.py uses, and the only one that is fast at scale."""
    texts = ["" if t is None else str(t) for t in texts]
    n = len(texts)
    if categories is None:
        categories = [""] * n
    categories = ["" if c is None else str(c) for c in categories]

    results: list[SentimentResult | None] = [None] * n

    # STAGE 1 — gate the procedural rows out before any model work happens.
    # STAGE 2a — templated results filings are decided by arithmetic, not tone.
    pending_idx, pending_txt = [], []
    for i, (t, c) in enumerate(zip(texts, categories)):
        if is_boilerplate(c, t):
            results[i] = NEUTRAL
            continue
        num = parse_result_move(t)
        if num is not None:
            sc, lab = num
            results[i] = SentimentResult(lab, sc, 1.0, "numeric")
            continue
        pending_idx.append(i)
        pending_txt.append(strip_preamble(f"{c}. {t}" if c else t))

    # STAGE 2 — model (or lexicon fallback if the model could not be loaded).
    if pending_txt:
        if _load_model():
            for i, (sc, cf, lab) in zip(pending_idx, _finbert_batch(pending_txt)):
                results[i] = SentimentResult(lab, round(float(sc), 3),
                                             round(float(cf), 3), "finbert")
        else:
            for i, txt in zip(pending_idx, pending_txt):
                lab, sc = analyze_sentiment(txt)
                results[i] = SentimentResult(lab, sc, 0.5, "lexicon")

    # STAGE 3 — deterministic overrides, applied to the model's verdict only.
    for i in pending_idx:
        results[i] = _apply_overrides(texts[i], results[i])

    return results                                # type: ignore[return-value]


def score_text(text: str, category: str = "") -> SentimentResult:
    """Single-row convenience wrapper. Prefer score_many() in bulk."""
    return score_many([text], [category])[0]


# --------------------------------------------------------------------------- #
# Conviction — how hard is the engine actually leaning?
# --------------------------------------------------------------------------- #
# Argmax gives a label even when the model is a whisker from a coin flip. Across
# the corpus the MEDIAN FinBERT |score| is 0.020 — P(pos)-P(neg) of two percent —
# and ~75% of directional tags sit below 0.057. Displaying those identically to a
# 0.9 lean overstates what the model said, and trading them is what made the
# backtest a coin flip (51.1% win rate, t=+1.34).
#
# Thresholds are per-engine because the scores are not the same quantity:
#   finbert   |P(pos) - P(neg)|, a probability gap
#   numeric   relative change in net profit vs the prior period it quotes
#   override  deterministic rule — always full conviction
CONVICTION_CUTS = {
    "finbert": (0.15, 0.50),      # (moderate above, strong above)
    "numeric": (0.10, 0.25),      # 10% and 25% profit moves
}


def conviction(res: "SentimentResult") -> str:
    """'strong' | 'moderate' | 'weak' | 'none'. Never invent conviction for a
    Neutral tag — there is no direction to be confident about."""
    if res.label == "Neutral":
        return "none"
    if res.engine == "override":
        return "strong"
    if res.engine == "rule":
        return "none"
    mod, strong = CONVICTION_CUTS.get(res.engine, (0.15, 0.50))
    mag = abs(res.score)
    if mag >= strong:
        return "strong"
    if mag >= mod:
        return "moderate"
    return "weak"


def conviction_of(label: str, score: float, engine: str) -> str:
    """Same rule, from the raw cached fields — so stock_server does not have to
    rebuild a SentimentResult just to ask."""
    return conviction(SentimentResult(label, float(score), 1.0, engine))


def build_text(row) -> str:
    """The canonical text for one announcement row — used by score_feeds.py and
    stock_server.py so the cache key and the live path never disagree."""
    return f"{row.get('desc', '')} {row.get('attchmntText', '')}".strip()


if __name__ == "__main__":                        # quick smoke test
    info = model_info()
    print(f"model={info['model']}  device={info.get('device')}  "
          f"available={info['available']}")
    samples = [
        ("Loss of Share Certificates",
         "Pursuant to Regulation 39(3) of the SEBI (Listing Obligations and "
         "Disclosure Requirements) Regulations, 2015, we enclose the information "
         "regarding loss of share certificate(s) received from the shareholder."),
        ("Acquisition", "Acquisition of 100% equity stake in Southern Health Foods."),
        ("Updates", "The Company has received a show-cause notice from the tax "
                    "department demanding Rs 250 crore."),
        ("Trading Window", "Trading Window closure pursuant to SEBI (PIT) Regulations, 2015"),
        ("Press Release", "Revenue grew 18% YoY and EBITDA margin expanded 220 bps."),
    ]
    for cat, txt in samples:
        r = score_text(txt, cat)
        print(f"  [{r.engine:8}] {r.label:8} {r.score:+.3f} conf={r.confidence:.2f}  "
              f"{cat} :: {txt[:60]}...")

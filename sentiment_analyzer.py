"""
sentiment_analyzer.py
===============================================================================
Lexicon-based financial sentiment analyzer for corporate filings/announcements.
Tailored for NSE corporate updates with basic negation handling and scoring.
No external dependencies (runs purely on standard library).
===============================================================================
"""

from __future__ import annotations
import re

# Financial/Corporate sentiment terms with specific weights
LEXICON = {
    # Strong positive / growth
    "growth": 1.0,
    "increase": 1.0,
    "increased": 1.0,
    "profit": 1.2,
    "revenue": 1.0,
    "dividend": 1.5,
    "dividends": 1.5,
    "bonus": 1.8,
    "acquisition": 1.2,
    "acquisitions": 1.2,
    "acquire": 1.2,
    "acquired": 1.2,
    "partnership": 1.0,
    "partnerships": 1.0,
    "collaboration": 1.0,
    "collaborations": 1.0,
    "expansion": 1.2,
    "expansions": 1.2,
    "secures": 1.5,
    "secured": 1.5,
    "won": 1.5,
    "awarded": 1.5,
    "successful": 1.2,
    "successfully": 1.2,
    "success": 1.0,
    "commissioned": 1.2,
    "commissioning": 1.2,
    "commercial production": 1.5,
    "gain": 1.0,
    "gains": 1.0,
    "exceeded": 1.0,
    "exceeds": 1.0,
    "surpasses": 1.2,
    "surpassed": 1.2,
    "soars": 1.5,
    "soared": 1.5,
    "highest": 1.2,
    "record": 1.2,
    "upgrade": 1.2,
    "upgraded": 1.2,
    "approval": 1.0,
    "approved": 1.2,
    "approve": 1.2,
    "incorporation": 1.0,
    "strategic": 1.0,
    "launch": 1.2,
    "launches": 1.2,
    "launched": 1.2,
    "invest": 1.0,
    "investment": 1.0,
    "investments": 1.0,
    
    # Strong negative / risks
    "loss": -1.5,
    "losses": -1.5,
    "decrease": -1.0,
    "decreased": -1.0,
    "negative": -1.0,
    "strike": -1.5,
    "strikes": -1.5,
    "dispute": -1.2,
    "disputes": -1.2,
    "penalty": -1.5,
    "penalties": -1.5,
    "penalized": -1.5,
    "fine": -1.0,
    "fined": -1.2,
    "fines": -1.0,
    "resigned": -1.0,
    "resignation": -1.0,
    "resignations": -1.0,
    "termination": -1.5,
    "decline": -1.0,
    "declined": -1.0,
    "fall": -1.0,
    "fell": -1.0,
    "drop": -1.0,
    "dropped": -1.0,
    "default": -2.0,
    "defaults": -2.0,
    "defaulted": -2.0,
    "fraud": -2.0,
    "investigation": -1.2,
    "investigations": -1.2,
    "disruption": -1.2,
    "disruptions": -1.2,
    "closure": -1.2,
    "closures": -1.2,
    "shutdown": -1.5,
    "court order": -1.2,
    "dismissed": -1.2,
    "dismissal": -1.2,
    "delay": -1.0,
    "delayed": -1.0,
    "delays": -1.0,
    "tax demand": -1.5,
    "cancellation": -1.5,
    "cancelled": -1.5,
    "breach": -1.8,
    "breached": -1.8,
    "litigation": -1.2,
    "show cause": -1.5,
    "demand": -0.8,
    "demands": -0.8,
    "rejected": -1.5,
    "reject": -1.2,
    "failure": -1.5,
    "failed": -1.5,
    "complaints": -1.0,
    "complaint": -1.0,
    "unauthorized": -1.5,
    "suspension": -1.5,
    "suspended": -1.5,
    "seizure": -1.5,
    "seized": -1.5,
    "raid": -1.8,
    "raids": -1.8,
    "defaulter": -2.0,
    "insolvency": -2.0,
    "bankruptcy": -2.0,
    "proceedings": -0.5,

    # Specific phrases overridden to be Neutral (0.0 weight)
    "loss of share certificate": 0.0,
    "loss of share certificates": 0.0,
    "loss of certificate": 0.0,
    "loss of certificates": 0.0,
    "loss of share": 0.0,
    "loss of shares": 0.0,
    "lost share certificate": 0.0,
    "lost share certificates": 0.0,
    "lost certificates": 0.0,
    "issue of duplicate": 0.0,
    "issue of duplicate share": 0.0,
    "duplicate share certificate": 0.0,
    "duplicate share certificates": 0.0,
    "loss of documents": 0.0,
}

# Negation words that reverse the sentiment polarity
NEGATIONS = {
    "no", "not", "never", "failed to", "without", "incorrect", "non", 
    "neither", "nor", "unable", "lack", "lacked", "lacks", "cannot", "cant",
    "didnt", "doesnt", "wasnt", "werent", "havent", "hasnt", "hadnt"
}

def analyze_sentiment(text: str) -> tuple[str, float]:
    """
    Analyze the sentiment of a given text string.
    Returns a tuple of (label, score) where:
      - label is one of 'Positive', 'Negative', 'Neutral'
      - score is a float compound score between -1.0 and +1.0
    """
    if not text or not isinstance(text, str):
        return "Neutral", 0.0
    
    # Preprocess text: lowercase
    text_lower = text.lower()
    
    # Simple tokenization: alphanumeric plus %
    words = re.findall(r'[a-z0-9%]+', text_lower)
    
    score = 0.0
    sentiment_words_count = 0
    
    i = 0
    while i < len(words):
        # Try matching phrases of length 4 down to 2 first
        matched_phrase = False
        for length in (4, 3, 2):
            if i + length <= len(words):
                phrase = " ".join(words[i:i+length])
                if phrase in LEXICON:
                    weight = LEXICON[phrase]
                    # Check for negation in the 4 words preceding the phrase start
                    negated = False
                    for j in range(max(0, i-4), i):
                        if words[j] in NEGATIONS:
                            negated = True
                            break
                    if negated:
                        weight = -weight * 0.5  # flip and scale down
                    score += weight
                    sentiment_words_count += 1
                    i += length
                    matched_phrase = True
                    break
        
        if matched_phrase:
            continue
            
        word = words[i]
        if word in LEXICON:
            weight = LEXICON[word]
            # Check for negation in the 4 words preceding
            negated = False
            for j in range(max(0, i-4), i):
                if words[j] in NEGATIONS:
                    negated = True
                    break
            if negated:
                weight = -weight * 0.5  # flip and scale down
            score += weight
            sentiment_words_count += 1
        
        i += 1

    # Normalize score
    if sentiment_words_count > 0:
        # Simple scaling capped between -1.0 and 1.0
        normalized_score = score / sentiment_words_count
        normalized_score = max(-1.0, min(1.0, normalized_score))
    else:
        normalized_score = 0.0

    # Determine label based on thresholds
    if normalized_score >= 0.05:
        label = "Positive"
    elif normalized_score <= -0.05:
        label = "Negative"
    else:
        label = "Neutral"
        
    return label, round(normalized_score, 2)

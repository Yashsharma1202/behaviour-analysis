from sentiment_analyzer import analyze_sentiment

test_cases = [
    # Positive cases
    ("Record profit growth and dividend of Rs 10 approved by the board.", "Positive"),
    ("Secured a major new contract worth Rs 500 crores from the government.", "Positive"),
    ("Successfully launched the new product range with double digit increase in sales.", "Positive"),
    ("Incorporation of a wholly owned subsidiary to expand EV passenger vehicles business.", "Positive"),

    # Negative cases
    ("Nashik plant tool-down strike declared by workmen.", "Negative"),
    ("Company penalized with a fine of Rs 5 lakhs for compliance delay.", "Negative"),
    ("The tax department issued a tax demand notice and show cause notice.", "Negative"),
    ("Defaulter list inclusion and insolvency proceedings initiated.", "Negative"),

    # Neutral cases
    ("Newspaper publication regarding loss of share certificate.", "Neutral"),
    ("Board meeting scheduled to consider financial results.", "Neutral"),
    ("Submission of investor presentation for the analyst call.", "Neutral"),

    # Negation cases (should flip or reduce sentiment)
    ("The board did not approve the dividend payment.", "Negative"), # "not approved" -> negative
    ("No growth was reported in the current quarter.", "Negative"), # "no growth" -> negative
]

all_pass = True
for text, expected in test_cases:
    label, score = analyze_sentiment(text)
    print(f"Text: {text[:60]}...")
    print(f"  Got: {label} ({score}) | Expected: {expected}")
    if label != expected:
        print("  FAIL!")
        all_pass = False
    else:
        print("  PASS")

if all_pass:
    print("\nALL TEST CASES PASSED SUCCESSFULLY!")
else:
    print("\nSOME TEST CASES FAILED.")
    exit(1)

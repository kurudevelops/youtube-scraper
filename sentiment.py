"""
Add sentiment (positive / neutral / negative) to the cleaned comments.

Setup:
    pip install transformers torch pandas openpyxl
    (first run downloads the model, roughly 1 GB, so it needs internet)

Run:
    python sentiment.py --in data_clean.csv --out data_sentiment.xlsx
    python sentiment.py --in data_clean.csv --sample 300 --min-confidence 0.7

Output: data_sentiment.xlsx (sheets: comments_sentiment, hand_label) + a CSV.
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

# Multilingual model trained on tweets in many languages, including Tagalog.
# English-only sentiment models do poorly on Taglish.
MODEL = "cardiffnlp/twitter-xlm-roberta-base-sentiment"
LABEL_FALLBACK = {"label_0": "negative", "label_1": "neutral", "label_2": "positive"}


def norm_label(label):
    """Model may return 'Negative' or 'LABEL_0' depending on its config."""
    l = label.lower()
    return LABEL_FALLBACK.get(l, l)


def load(path):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name="comments_clean")


def score_texts(texts, batch_size):
    import torch
    from transformers import pipeline

    clf = pipeline(
        "text-classification", model=MODEL, top_k=None,
        truncation=True, max_length=256,
        device=0 if torch.cuda.is_available() else -1,
    )
    scores = []
    for i in range(0, len(texts), batch_size):
        batch = texts[i:i + batch_size]
        for result in clf(batch, batch_size=batch_size):
            scores.append({norm_label(r["label"]): r["score"] for r in result})
        print(f"  {min(i + batch_size, len(texts))}/{len(texts)}")
    return scores


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", default="data_clean.csv")
    ap.add_argument("--out", default="data_sentiment.xlsx")
    ap.add_argument("--batch-size", type=int, default=32)
    ap.add_argument("--min-confidence", type=float, default=0.6,
                    help="below this, the row is flagged low_confidence")
    ap.add_argument("--sample", type=int, default=200,
                    help="rows to export for hand-labelling (0 to skip)")
    args = ap.parse_args()

    df = load(args.inp)
    df["text_clean"] = df["text_clean"].fillna("").astype(str)
    print(f"Scoring {len(df)} comments (first run downloads the model)...")

    scores = score_texts(df["text_clean"].tolist(), args.batch_size)
    probs = pd.DataFrame(scores).reindex(
        columns=["negative", "neutral", "positive"]).add_prefix("p_")
    df = pd.concat([df.reset_index(drop=True), probs], axis=1)

    df["sentiment"] = probs.idxmax(axis=1).str.replace("p_", "", regex=False)
    df["confidence"] = probs.max(axis=1)
    df["low_confidence"] = df["confidence"] < args.min_confidence

    out = Path(args.out)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="comments_sentiment", index=False)
        ws = xw.sheets["comments_sentiment"]
        ws.freeze_panes = "A2"
        for i, col in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = (
                70 if col in ("text", "text_clean") else 14)

        if args.sample:
            # Label these by hand to measure how accurate the model is.
            sample = df.sample(min(args.sample, len(df)), random_state=42)[
                ["comment_id", "text_clean", "sentiment"]].rename(
                columns={"sentiment": "model_label"})
            sample["my_label"] = ""
            sample.to_excel(xw, sheet_name="hand_label", index=False)
            ws2 = xw.sheets["hand_label"]
            ws2.column_dimensions["A"].width = 28
            ws2.column_dimensions["B"].width = 90
            ws2.column_dimensions["C"].width = 14
            ws2.column_dimensions["D"].width = 14
    df.to_csv(out.with_suffix(".csv"), index=False, encoding="utf-8-sig")

    print("\nSentiment:\n" + df["sentiment"].value_counts().to_string())
    print(f"\nLow-confidence rows: {int(df['low_confidence'].sum())}")
    if "topic_tag" in df.columns:
        print("\nBy topic:\n" + pd.crosstab(df["topic_tag"], df["sentiment"]).to_string())
    print(f"\nSaved -> {out} and {out.with_suffix('.csv')}")


if __name__ == "__main__":
    main()
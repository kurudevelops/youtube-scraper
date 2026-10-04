"""
Clean the YouTube comments collected by main.py / youtube_collector.py.

Setup:
    pip install pandas openpyxl

Run:
    python clean_data.py --in data.xlsx --out data_clean.xlsx
    python clean_data.py --in data.xlsx --min-words 6 --only-tagged

Reads the `comments` sheet; writes data_clean.xlsx (sheets: comments_clean,
report) and a UTF-8 CSV of the same cleaned rows next to it.
The original `text` column is kept; the cleaned version is `text_clean`.
"""
import argparse
import html
import re
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

TAG_RE = re.compile(r"<[^>]+>")
URL_RE = re.compile(r"https?://\S+|www\.\S+", re.I)
MENTION_RE = re.compile(r"@[\w.\-]+")
TIMESTAMP_RE = re.compile(r"\b\d{1,2}:\d{2}(?::\d{2})?\b")  # "1:23" video jumps
REPEAT_RE = re.compile(r"(.)\1{2,}")  # "hirap naaaaa" -> "hirap naa"
WORD_RE = re.compile(r"[^\W\d_]+(?:'[^\W\d_]+)?", re.UNICODE)  # letters only

SPAM_RE = re.compile(
    r"subscribe|sub\s?back|check (out )?my (channel|video|page)|giveaway|"
    r"whatsapp|telegram|dm me|click (the )?link|free gift|follow me|visit my|"
    r"promo\s?code|earn \$|crypto|forex", re.I)

MATATAG_RE = re.compile(r"matatag|kto10|k\s?to\s?10", re.I)
K12_RE = re.compile(r"k[\s-]?12|k\s?to\s?12|senior\s?high|\bshs\b", re.I)

# Tiny word lists for a rough language guess. Taglish defeats most language
# detectors, so we count common function words instead.
TL_WORDS = {
    "ang", "ng", "mga", "sa", "na", "ay", "po", "ako", "ka", "mo", "ko",
    "naman", "lang", "din", "rin", "kasi", "pero", "hindi", "wala", "yung",
    "dapat", "sana", "nga", "talaga", "mas", "para", "kung", "bakit", "gusto",
    "ayaw", "hirap", "guro", "bata", "estudyante", "paaralan", "ito", "niya",
    "natin", "namin", "kayo", "siya", "may", "ba", "pa", "daw", "raw", "ni",
}
EN_WORDS = {
    "the", "is", "are", "and", "to", "of", "in", "that", "it", "for", "this",
    "with", "not", "they", "be", "have", "should", "because", "but", "was",
    "will", "from", "just", "more", "students", "teachers", "teacher", "so",
    "we", "you", "their", "has", "can", "if", "what", "when", "all", "our",
}


def clean_text(t):
    """Normalise one comment. Emojis and punctuation are kept on purpose
    (they carry sentiment); only noise is removed."""
    t = TAG_RE.sub(" ", str(t))
    t = html.unescape(t)
    t = unicodedata.normalize("NFKC", t)  # fancy unicode letters -> plain
    t = URL_RE.sub(" ", t)
    t = MENTION_RE.sub(" ", t)
    t = TIMESTAMP_RE.sub(" ", t)
    t = REPEAT_RE.sub(r"\1\1", t)
    return re.sub(r"\s+", " ", t).strip()


def tag_topic(text):
    m, k = bool(MATATAG_RE.search(text)), bool(K12_RE.search(text))
    return "both" if m and k else "matatag" if m else "k12" if k else "none"


def guess_lang(text):
    words = [w.lower() for w in WORD_RE.findall(text)]
    tl = sum(w in TL_WORDS for w in words)
    en = sum(w in EN_WORDS for w in words)
    if tl >= 2 and en >= 2:
        return "taglish"
    if tl > en:
        return "tl"
    if en > tl:
        return "en"
    return "unknown"


def dedupe_key(text):
    """Letters/digits only, lowercase: catches copy-pasted spam that differs
    only in spacing, punctuation, or emoji."""
    return re.sub(r"[\W_]+", "", text.lower())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", default="data.xlsx")
    ap.add_argument("--out", default="data_clean.xlsx")
    ap.add_argument("--min-words", type=int, default=4,
                    help="drop comments with fewer letter-words than this")
    ap.add_argument("--only-tagged", action="store_true",
                    help="keep only comments that name MATATAG / K-12")
    args = ap.parse_args()

    df = pd.read_excel(args.inp, sheet_name="comments")
    report = [("loaded", len(df))]

    df = df.drop_duplicates(subset="comment_id")
    report.append(("after dropping duplicate comment_id", len(df)))

    df["text"] = df["text"].fillna("").astype(str)
    df["text_clean"] = df["text"].map(clean_text)

    df = df[~df["text_clean"].map(lambda t: bool(SPAM_RE.search(t)))]
    report.append(("after removing spam/promo", len(df)))

    df["word_count"] = df["text_clean"].map(lambda t: len(WORD_RE.findall(t)))
    df = df[df["word_count"] >= args.min_words]
    report.append((f"after dropping < {args.min_words} words", len(df)))

    df = df[~df["text_clean"].map(dedupe_key).duplicated()]
    report.append(("after dropping copy-pasted duplicates", len(df)))

    # Re-tag on the cleaned text (the collector tagged the raw text).
    df["topic_tag"] = df["text_clean"].map(tag_topic)
    if args.only_tagged:
        df = df[df["topic_tag"] != "none"]
        report.append(("after keeping only tagged comments", len(df)))

    df["lang_guess"] = df["text_clean"].map(guess_lang)
    df["is_reply"] = df["parent_id"].notna()
    when = pd.to_datetime(df["published_at"], utc=True, errors="coerce")
    df["published_at"] = when.dt.tz_localize(None)
    df["year_month"] = when.dt.strftime("%Y-%m")

    cols = ["comment_id", "video_id", "parent_id", "is_reply", "published_at",
            "year_month", "like_count", "topic_tag", "lang_guess",
            "word_count", "text", "text_clean"]
    df = df[cols].reset_index(drop=True)

    out = Path(args.out)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="comments_clean", index=False)
        pd.DataFrame(report, columns=["step", "rows"]).to_excel(
            xw, sheet_name="report", index=False)
        ws = xw.sheets["comments_clean"]
        ws.freeze_panes = "A2"
        for i, w in enumerate([28, 14, 28, 9, 20, 11, 10, 10, 11, 11, 70, 70], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        xw.sheets["report"].column_dimensions["A"].width = 42
    df.to_csv(out.with_suffix(".csv"), index=False, encoding="utf-8-sig")

    print("\nCleaning report")
    for step, n in report:
        print(f"  {n:>7}  {step}")
    print("\nTopic tags:\n" + df["topic_tag"].value_counts().to_string())
    print("\nLanguage guess:\n" + df["lang_guess"].value_counts().to_string())
    print(f"\nSaved -> {out} and {out.with_suffix('.csv')}")


if __name__ == "__main__":
    main()
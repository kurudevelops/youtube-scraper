"""
YouTube collector for MATATAG / K-12 opinion scraping -> Excel output.

Setup:
    pip install google-api-python-client openpyxl python-dotenv
    .env:
        YOUTUBE_API_KEY=your_key

Run:
    python youtube_collector.py
    python youtube_collector.py --max-videos 20 --max-comments 500 --out data.xlsx

Rerunning appends only new videos/comments (deduped by ID). Close the
Excel file before running, or the save will fail.
"""
import argparse
import os
import re

from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

load_dotenv()

SEARCH_QUERIES = [
    "MATATAG curriculum",
    "MATATAG kurikulum DepEd",
    "MATATAG curriculum teachers",
    "K-12 curriculum Philippines",
    "K to 12 Philippines opinion",
    "senior high school Philippines removal",
    "DepEd curriculum review",
    "MATATAG reaksyon",
]

MATATAG_RE = re.compile(r"matatag|kto10|k\s?to\s?10", re.I)
K12_RE = re.compile(r"k[\s-]?12|k\s?to\s?12|senior\s?high|\bshs\b", re.I)
ILLEGAL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")  # openpyxl rejects these

VIDEO_COLS = ["video_id", "title", "channel_title", "published_at", "query"]
COMMENT_COLS = ["comment_id", "video_id", "parent_id", "text",
                "like_count", "published_at", "topic_tag"]
# Author names are deliberately not stored (privacy, if results get published).


def clean(v):
    """Excel cells: strip illegal control chars, cap at the 32,767-char limit."""
    return ILLEGAL_RE.sub("", v)[:32000] if isinstance(v, str) else v


def tag_topic(text):
    m, k = bool(MATATAG_RE.search(text)), bool(K12_RE.search(text))
    return "both" if m and k else "matatag" if m else "k12" if k else "none"


def open_workbook(path):
    """Load existing file (and the IDs already in it) or start a new one."""
    if os.path.exists(path):
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.remove(wb.active)
    sheets = {}
    for name, cols in (("videos", VIDEO_COLS), ("comments", COMMENT_COLS)):
        if name in wb.sheetnames:
            ws = wb[name]
        else:
            ws = wb.create_sheet(name)
            ws.append(cols)
            for c in ws[1]:
                c.font = Font(name="Arial", bold=True)
            ws.freeze_panes = "A2"
        sheets[name] = ws
    seen_videos = {r[0] for r in sheets["videos"].iter_rows(min_row=2, values_only=True)}
    seen_comments = {r[0] for r in sheets["comments"].iter_rows(min_row=2, values_only=True)}
    return wb, sheets, seen_videos, seen_comments


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def search_videos(yt, query, max_videos):
    """search.list costs 100 quota units per call; keep queries few."""
    found, token = [], None
    while len(found) < max_videos:
        resp = yt.search().list(
            q=query, part="snippet", type="video",
            maxResults=min(50, max_videos - len(found)),
            regionCode="PH", relevanceLanguage="en", pageToken=token,
        ).execute()
        for item in resp["items"]:
            s = item["snippet"]
            found.append((item["id"]["videoId"], s["title"],
                          s["channelTitle"], s["publishedAt"], query))
        token = resp.get("nextPageToken")
        if not token:
            break
    return found


def parse_comment(c, video_id, parent_id):
    s = c["snippet"]
    text = s["textDisplay"]
    return (c["id"], video_id, parent_id, text, s.get("likeCount", 0),
            s["publishedAt"], tag_topic(text))


def fetch_comments(yt, video_id, max_comments):
    """commentThreads.list costs 1 unit per call (100 threads max)."""
    rows, token = [], None
    while len(rows) < max_comments:
        try:
            resp = yt.commentThreads().list(
                part="snippet,replies", videoId=video_id, maxResults=100,
                order="relevance", textFormat="plainText", pageToken=token,
            ).execute()
        except HttpError as e:
            if e.resp.status in (403, 404):  # comments disabled / video gone
                print(f"  skip {video_id}: {e.resp.status}")
                return rows
            raise
        for thread in resp["items"]:
            top = thread["snippet"]["topLevelComment"]
            rows.append(parse_comment(top, video_id, None))
            for reply in thread.get("replies", {}).get("comments", []):
                rows.append(parse_comment(reply, video_id, top["id"]))
        token = resp.get("nextPageToken")
        if not token:
            break
    return rows[:max_comments]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--max-videos", type=int, default=15, help="per query")
    ap.add_argument("--max-comments", type=int, default=300, help="per video")
    ap.add_argument("--out", default="matatag_youtube.xlsx")
    args = ap.parse_args()

    yt = build("youtube", "v3", developerKey=os.environ["YOUTUBE_API_KEY"])
    wb, sheets, seen_videos, seen_comments = open_workbook(args.out)

    for q in SEARCH_QUERIES:
        print(f"[search] {q}")
        for vid, title, ch, pub, query in search_videos(yt, q, args.max_videos):
            if vid in seen_videos:
                continue
            seen_videos.add(vid)
            sheets["videos"].append([clean(x) for x in (vid, title, ch, pub, query)])
            new = 0
            for row in fetch_comments(yt, vid, args.max_comments):
                if row[0] in seen_comments:
                    continue
                seen_comments.add(row[0])
                sheets["comments"].append([clean(x) for x in row])
                new += 1
            wb.save(args.out)  # save per video so a crash/quota stop loses nothing
            print(f"  {vid}: {new} new comments | {title[:60]}")

    set_widths(sheets["videos"], [14, 60, 28, 22, 30])
    set_widths(sheets["comments"], [28, 14, 28, 90, 10, 22, 12])
    wb.save(args.out)
    print(f"Saved -> {args.out}")


if __name__ == "__main__":
    main()
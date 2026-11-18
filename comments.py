"""Шаг 2: комментарии под роликами-выбросами конкурентов.
Квота: commentThreads.list = 1 единица за 100 комментариев.
"""
import re
import sys

import pandas as pd

from yt_research import call

PER_GROUP = int(sys.argv[1]) if len(sys.argv) > 1 else 15
import glob
SRC = sorted(glob.glob('output/yt_research_*.xlsx'))[-1]  # последний прогон
STAMP = SRC.split('_')[-1][:-5]

# комментарии-запросы: просят продолжение или новую тему
REQ = re.compile(r"(part ?2|part two|next|please (make|do|tell)|can you (make|do)|do one|story about|"
                 r"parte ?2|segunda parte|haz (una|un|otro)|cuenta(n)? (la|una)|por favor|"
                 r"teil ?2|bitte (mach|erzähl)|fortsetzung|nächste|mehr davon|"
                 r"what about|how about|should do|would love)", re.I)

d = pd.read_excel(SRC, sheet_name='Все_ролики')
d = d[(d.outlier_x >= 3) & (d.age_days >= 7) & (d.age_days <= 400) & (d.comments > 0)]
pick = (d.sort_values('outlier_x', ascending=False).groupby(['niche', 'lang', 'channel']).head(2)
        .groupby(['niche', 'lang']).head(PER_GROUP))

rows = []
for _, v in pick.iterrows():
    try:
        data = call("commentThreads", part="snippet", videoId=v.video_id, order="relevance",
                    maxResults=100, textFormat="plainText")
    except RuntimeError as e:  # комментарии отключены и т.п.
        print("skip", v.video_id, str(e)[:80])
        continue
    for it in data.get("items", []):
        c = it["snippet"]["topLevelComment"]["snippet"]
        txt = c["textDisplay"].replace("\n", " ")
        rows.append({"niche": v.niche, "lang": v.lang, "video_title": v.title, "video_url": v.url,
                     "outlier_x": v.outlier_x, "likes": c.get("likeCount", 0),
                     "replies": it["snippet"].get("totalReplyCount", 0),
                     "is_request": bool(REQ.search(txt)), "comment": txt})
c = pd.DataFrame(rows)
c.to_csv(f'output/comments_{STAMP}.csv', index=False)
print(len(pick), "videos,", len(c), "comments,", c.is_request.sum(), "requests")

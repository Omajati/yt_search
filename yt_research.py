"""Анализ ниш и банк тем по YouTube Data API v3.

Ключ: подставляется прокси окружения (Network secret на www.googleapis.com,
заголовок X-Goog-Api-Key) либо берётся из переменной YOUTUBE_API_KEY.

Запуск:
    python3 yt_research.py            # полный прогон (кэш сохраняется в cache/)
    python3 yt_research.py --check    # проверить, что ключ работает (1 единица)
"""
import datetime as dt
import hashlib
import json
import os
import re
import statistics
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
import requests

import config as C

API = "https://www.googleapis.com/youtube/v3/"
HERE = Path(__file__).parent
CACHE = HERE / "cache"
OUT = HERE / "output"
CACHE.mkdir(exist_ok=True)
OUT.mkdir(exist_ok=True)

session = requests.Session()
if os.environ.get("YOUTUBE_API_KEY"):
    session.headers["X-Goog-Api-Key"] = os.environ["YOUTUBE_API_KEY"]

units_used = 0
COST = {"search": 100}


def call(endpoint, **params):
    """GET с кэшем на диске: повторный прогон не тратит квоту."""
    global units_used
    key = hashlib.sha1(json.dumps([endpoint, params], sort_keys=True).encode()).hexdigest()
    f = CACHE / f"{endpoint}_{key}.json"
    if f.exists():
        return json.loads(f.read_text())
    if endpoint == "search" and "--cached-search-only" in sys.argv:
        print("  (пропущен: нет в кэше)")
        return {"items": []}
    for attempt in range(6):
        r = session.get(API + endpoint, params=params, timeout=30)
        if r.status_code in (403, 429) and "rateLimitExceeded" in r.text:
            wait = 65 * (attempt + 1)
            print(f"  поминутный лимит, жду {wait} c")
            time.sleep(wait)
            continue
        break
    if r.status_code != 200:
        raise RuntimeError(f"{endpoint} {r.status_code}: {r.text[:500]}")
    units_used += COST.get(endpoint, 1)
    data = r.json()
    f.write_text(json.dumps(data))
    return data


def chunks(xs, n=50):
    for i in range(0, len(xs), n):
        yield xs[i:i + n]


def iso_dur(s):
    m = re.match(r"P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", s or "")
    if not m:
        return 0
    d, h, mi, se = (int(x or 0) for x in m.groups())
    return d * 86400 + h * 3600 + mi * 60 + se


def get_videos(ids):
    out = {}
    for ch in chunks(list(ids)):
        data = call("videos", part="snippet,statistics,contentDetails", id=",".join(ch), maxResults=50)
        for it in data.get("items", []):
            out[it["id"]] = it
    return out


def get_channels(ids):
    out = {}
    for ch in chunks(list(ids)):
        data = call("channels", part="snippet,statistics,contentDetails", id=",".join(ch), maxResults=50)
        for it in data.get("items", []):
            out[it["id"]] = it
    return out


def get_uploads(playlist_id, limit):
    ids, token = [], None
    while len(ids) < limit:
        p = dict(part="contentDetails", playlistId=playlist_id, maxResults=50)
        if token:
            p["pageToken"] = token
        try:
            data = call("playlistItems", **p)
        except RuntimeError as e:
            print("  ! uploads:", e)
            break
        ids += [it["contentDetails"]["videoId"] for it in data.get("items", [])]
        token = data.get("nextPageToken")
        if not token:
            break
    return ids[:limit]


def video_row(v, now):
    sn, st = v["snippet"], v.get("statistics", {})
    pub = dt.datetime.fromisoformat(sn["publishedAt"].replace("Z", "+00:00"))
    age = max((now - pub).total_seconds() / 86400, 1)
    views = int(st.get("viewCount", 0))
    secs = iso_dur(v["contentDetails"].get("duration"))
    thumbs = sn.get("thumbnails", {})
    thumb = (thumbs.get("maxres") or thumbs.get("high") or thumbs.get("medium") or {}).get("url", "")
    return {
        "video_id": v["id"],
        "title": sn["title"],
        "channel_id": sn["channelId"],
        "channel": sn["channelTitle"],
        "published": pub.date().isoformat(),
        "age_days": round(age),
        "views": views,
        "views_per_day": round(views / age, 1),
        "likes": int(st.get("likeCount", 0)),
        "comments": int(st.get("commentCount", 0)),
        "duration_s": secs,
        "is_short": secs <= C.SHORTS_MAX_SECONDS,
        "lang_audio": sn.get("defaultAudioLanguage", ""),
        "url": f"https://youtu.be/{v['id']}",
        "thumbnail": thumb,
    }


TITLE_FEATURES = {
    "has_number": r"\d",
    "question": r"\?",
    "exclaim": r"!",
    "caps_word": r"\b[A-ZÄÖÜÁÉÍÓÚÑ]{3,}\b",
    "first_person": r"\b(I|my|me|yo|mi|me|ich|mein|meine|mich|mir)\b",
    "year": r"\b20\d\d\b",
    "brackets": r"[\(\[]",
    "emoji": r"[\U0001F300-\U0001FAFF☀-➿]",
}


def title_feats(t):
    return {k: bool(re.search(p, t)) for k, p in TITLE_FEATURES.items()} | {"title_len": len(t)}


STOP = set("""the a an of to and in on for is it my me i you your with at by from this that be was are her his he she
they them we our how what why who when after before about into out up just not no but or so as
el la los las de del y en un una que por con para su sus mi me se lo le es al más pero como
der die das und in den von zu mit ist ein eine sich nicht auf für im dem des ich mein meine er sie es""".split())


def ngrams(titles, n):
    c = Counter()
    for t in titles:
        w = [x for x in re.findall(r"[\wäöüßáéíóúñ']+", t.lower()) if x not in STOP and not x.isdigit()]
        c.update(" ".join(w[i:i + n]) for i in range(len(w) - n + 1))
    return c


def main():
    if "--lite" in sys.argv:  # экономный режим при остатке квоты
        C.TOP_CHANNELS_PER_GROUP, C.UPLOADS_PER_CHANNEL = 8, 150
    if "--check" in sys.argv:
        r = session.get(API + "videos", params={"part": "id", "id": "dQw4w9WgXcQ"}, timeout=30)
        print(r.status_code, r.text[:300])
        return

    now = dt.datetime.now(dt.timezone.utc)
    after = (now - dt.timedelta(days=30 * C.SEARCH_MONTHS_BACK)).strftime("%Y-%m-%dT00:00:00Z")

    # 1. Поиск по запросам -> ролики выдачи
    search_rows = []
    for niche, langs in C.NICHES.items():
        for lang, queries in langs.items():
            region = C.LANGS[lang]["region"]
            for q in queries:
              for dur in C.SEARCH_DURATIONS:
                print(f"search [{niche}/{lang}/{dur}] {q}")
                data = call("search", part="snippet", q=q, type="video", order="viewCount",
                            publishedAfter=after, relevanceLanguage=lang, regionCode=region,
                            videoDuration=dur, maxResults=C.SEARCH_RESULTS)
                for rank, it in enumerate(data.get("items", []), 1):
                    search_rows.append({"niche": niche, "lang": lang, "query": q, "duration_filter": dur,
                                        "rank": rank, "video_id": it["id"]["videoId"]})

    vids = get_videos({r["video_id"] for r in search_rows})
    chans = get_channels({v["snippet"]["channelId"] for v in vids.values()})

    def subs(cid):
        st = chans.get(cid, {}).get("statistics", {})
        return None if st.get("hiddenSubscriberCount") else int(st.get("subscriberCount", 0))

    serp = []
    for r in search_rows:
        v = vids.get(r["video_id"])
        if not v:
            continue
        row = r | video_row(v, now)
        if row["is_short"]:
            continue  # только длинные ролики
        row["channel_subs"] = subs(row["channel_id"])
        row["views_per_sub"] = round(row["views"] / row["channel_subs"], 2) if row["channel_subs"] else None
        serp.append(row)
    serp = pd.DataFrame(serp)

    # 2. Глубокий разбор топ-каналов по каждой (нише, языку)
    deep_rows, chan_rows = [], []
    done = set()
    for (niche, lang), g in serp.groupby(["niche", "lang"]):
        # профильные каналы: язык совпадает, появляются по нескольким запросам ниши
        a = g["lang_audio"].fillna("").str[:2]
        g = g[(a == lang) | (a == "")]
        agg = g.groupby("channel_id").agg(q=("query", "nunique"), n=("video_id", "nunique"),
                                          vps=("views_per_sub", "median"))
        agg = agg[(agg.q >= 2) | (agg.n >= 3)]
        top = agg.sort_values(["q", "n", "vps"], ascending=False).head(C.TOP_CHANNELS_PER_GROUP).index
        for cid in top:
            ch = chans.get(cid)
            if not ch:
                continue
            print(f"channel [{niche}/{lang}] {ch['snippet']['title']}")
            if cid not in done:
                up = ch["contentDetails"]["relatedPlaylists"]["uploads"]
                ids = get_uploads(up, C.UPLOADS_PER_CHANNEL)
                cv = get_videos(ids)
                rows = [video_row(v, now) for v in cv.values()]
                # медиана отдельно для длинных и Shorts — у них разные масштабы
                for is_short in (True, False):
                    part = [x for x in rows if x["is_short"] == is_short and x["age_days"] >= 7]
                    med = statistics.median([x["views"] for x in part]) if part else None
                    for x in rows:
                        if x["is_short"] == is_short:
                            x["channel_median"] = med
                            x["outlier_x"] = round(x["views"] / med, 2) if med else None
                for x in rows:
                    x.update(title_feats(x["title"]))
                deep_rows += [{"niche": niche, "lang": lang} | x for x in rows if not x["is_short"]]
                done.add(cid)
                long_med = next((x["channel_median"] for x in rows if not x["is_short"]), None)
                chan_rows.append({
                    "niche": niche, "lang": lang, "channel": ch["snippet"]["title"],
                    "url": f"https://youtube.com/channel/{cid}",
                    "subs": subs(cid), "total_views": int(ch["statistics"].get("viewCount", 0)),
                    "videos": int(ch["statistics"].get("videoCount", 0)),
                    "created": ch["snippet"]["publishedAt"][:10],
                    "median_views_long": long_med,
                    "uploads_last_30d": sum(1 for x in rows if x["age_days"] <= 30),
                    "shorts_share": round(sum(x["is_short"] for x in rows) / max(len(rows), 1), 2),
                    "outliers": sum(1 for x in rows if (x["outlier_x"] or 0) >= C.OUTLIER_THRESHOLD),
                })
    deep = pd.DataFrame(deep_rows)
    channels = pd.DataFrame(chan_rows)

    # 3. Сводка по нишам
    summary = []
    for (niche, lang), g in serp.groupby(["niche", "lang"]):
        g = g.drop_duplicates("video_id")
        hits = g[g["views"] >= 100_000]
        small = hits[hits["channel_subs"].fillna(1e12) < C.SMALL_CHANNEL_SUBS]
        summary.append({
            "niche": niche, "lang": lang, "videos_in_serp": len(g),
            "median_views": int(g["views"].median()),
            "median_views_per_day": round(g["views_per_day"].median(), 1),
            "videos_100k+": len(hits),
            "small_channel_hits": len(small),
            "small_channel_hit_share": round(len(small) / max(len(hits), 1), 2),
            "fresh_90d_share": round((g["age_days"] <= 90).mean(), 2),
            "shorts_share": round(g["is_short"].mean(), 2),
            "unique_channels": g["channel_id"].nunique(),
            "median_views_per_sub": g["views_per_sub"].median(),
        })
    summary = pd.DataFrame(summary)

    # 4. Банк тем: ролики-выбросы
    outl = deep[(deep["outlier_x"] >= C.OUTLIER_THRESHOLD) & (deep["age_days"] >= 7)]
    outl = outl.sort_values(["niche", "lang", "outlier_x"], ascending=[True, True, False])

    # 5. Паттерны названий: выбросы vs остальные
    pat = []
    for (niche, lang), g in deep.groupby(["niche", "lang"]):
        o = g[g["outlier_x"] >= C.OUTLIER_THRESHOLD]
        rest = g[g["outlier_x"] < 1]
        for f in list(TITLE_FEATURES) + ["title_len"]:
            pat.append({"niche": niche, "lang": lang, "feature": f,
                        "outliers": round(o[f].mean(), 2) if len(o) else None,
                        "below_median": round(rest[f].mean(), 2) if len(rest) else None})
    patterns = pd.DataFrame(pat)

    phrases = []
    for (niche, lang), g in deep.groupby(["niche", "lang"]):
        o = g[g["outlier_x"] >= C.OUTLIER_THRESHOLD]["title"]
        for n in (1, 2, 3):
            for ph, cnt in ngrams(o, n).most_common(40):
                if cnt >= 2:
                    phrases.append({"niche": niche, "lang": lang, "n": n, "phrase": ph, "count_in_outliers": cnt})
    phrases = pd.DataFrame(phrases)

    stamp = now.strftime("%Y-%m-%d")
    path = OUT / f"yt_research_{stamp}.xlsx"
    with pd.ExcelWriter(path) as xw:
        summary.to_excel(xw, sheet_name="Ниши_сводка", index=False)
        outl.to_excel(xw, sheet_name="Банк_тем_выбросы", index=False)
        channels.to_excel(xw, sheet_name="Каналы", index=False)
        patterns.to_excel(xw, sheet_name="Названия_признаки", index=False)
        phrases.to_excel(xw, sheet_name="Названия_фразы", index=False)
        serp.to_excel(xw, sheet_name="Выдача", index=False)
        deep.to_excel(xw, sheet_name="Все_ролики", index=False)
    deep.to_csv(OUT / f"all_videos_{stamp}.csv", index=False)
    outl.to_csv(OUT / f"outliers_{stamp}.csv", index=False)
    print(f"\nГотово: {path}\nКвоты потрачено в этом прогоне: ~{units_used} единиц")


if __name__ == "__main__":
    main()

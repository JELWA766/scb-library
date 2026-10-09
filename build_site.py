#!/usr/bin/env python3
"""
Build the public fan website from a library workbook.

  python build_site.py --workbook public_library.xlsx --out site

Produces:
  site/index.html                      single self-contained page (data embedded)
  site/SCB_Library_Fan_Edition.xlsx    the Excel tracker, with NO personal data

Personal data (Mark, Watched On, Notes, Watched Through) is stripped; descriptions are kept in a
cleaned, shortened form (URLs and repeated sponsor text removed) so they can be searched. Videos
that are no longer on YouTube (or not yet public) are left out.
"""
import argparse, collections, json, re, shutil, sys
from datetime import datetime
from pathlib import Path

import scb_library as S
import events_data as EV

HERE = Path(__file__).resolve().parent
XLSX_NAME = "SCB_Library_Fan_Edition.xlsx"


DESC_MAX = 1200          # characters kept per description
BOILERPLATE_MIN = 25     # a line repeated in at least this many descriptions is treated as boilerplate


def _lines(text):
    return [re.sub(r"\s+", " ", re.sub(r"https?://\S+", "", ln)).strip() for ln in (text or "").splitlines()]


def clean_descriptions(vids):
    """Strip URLs and sponsor/social boilerplate; keep a searchable summary of each description."""
    count = collections.Counter(ln for v in vids for ln in set(_lines(v.get("description"))) if len(ln) > 3)
    common = {ln for ln, c in count.items() if c >= BOILERPLATE_MIN}
    for v in vids:
        keep = [ln for ln in _lines(v.get("description")) if ln and ln not in common]
        v["description"] = " ".join(keep)[:DESC_MAX]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=str(HERE / "SCB_Library.xlsx"))
    ap.add_argument("--out", default=str(HERE / "site"))
    ap.add_argument("--template", default=str(HERE / "site_template.html"))
    a = ap.parse_args()

    src = Path(a.workbook)
    if not src.exists():
        sys.exit(f"Workbook not found: {src}")
    d = S.read_workbook(src)
    tzname = d["settings"].get("tz") or S.DEFAULT_TZ
    tz = S.get_tz(tzname)
    now = datetime.now().replace(microsecond=0)

    channels = [dict(c, through=None, refreshed=now) for c in d["channels"] if c.get("id")]
    vids = []
    for v in d["videos"]:
        if v.get("availability") in ("Not in last refresh", "Upcoming/Live"):
            continue
        v = dict(v)
        v.update(mark=None, watched_on=None, notes=None, first_seen=None)
        vids.append(v)
    clean_descriptions(vids)
    vids = S.finalize(vids, channels, tz)
    # drop channels with no videos
    used = {v["channel_id"] for v in vids}
    channels = [c for c in channels if c["id"] in used]
    cidx = {c["id"]: i for i, c in enumerate(channels)}

    rows = [[v.get("video_id") or "", v["title"], cidx[v["channel_id"]], v["local"].strftime("%Y-%m-%d"),
             v["local"].strftime("%H:%M"), v["seconds"], v.get("views") or 0, v.get("description") or ""] for v in vids if v["channel_id"] in cidx]
    payload = dict(
        generated=now.strftime("%B %-d, %Y") if sys.platform != "win32" else now.strftime("%B %#d, %Y"),
        tz=tzname,
        xlsx=XLSX_NAME,
        channels=[dict(n=c["name"], id=c["id"], c="#" + S.PALETTE[i % len(S.PALETTE)]) for i, c in enumerate(channels)],
        videos=rows)

    # ---- calendar events: built-in + auto milestones + the owner's optional my_events.csv
    y0 = min(v["local"].year for v in vids)
    y1 = max(max(v["local"].year for v in vids), now.year) + 1
    extra = EV.load_csv(HERE / "my_events.csv")
    built = EV.all_events(y0, y1, vids, channels, extra=extra)
    ev_out = []
    for e in built:
        cat = EV.CAT_NAMES.index(e["cat"]) if e["cat"] in EV.CAT_NAMES else len(EV.CAT_NAMES) - 1
        for st, en, n in EV.occurrences(e, y1, y0 - 1):
            title = e["title"].replace("{n}", str(n if n is not None else 0))
            ev_out.append([st.isoformat(), en.isoformat() if en != st else "", title, cat, e.get("notes") or "", e.get("kw") or 0])
    ev_out.sort(key=lambda r: r[0])
    payload["events"] = ev_out
    payload["cats"] = [[c[0], c[1], "#" + c[2]] for c in EV.CATEGORIES]

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    xlsx = out / XLSX_NAME
    S.build(xlsx, vids, channels, tzname, {}, now, fan=True, events=built)
    size_mb = xlsx.stat().st_size / 1e6
    payload["xlsx_mb"] = round(size_mb, 1)

    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = Path(a.template).read_text(encoding="utf-8").replace("__DATA__", data)
    (out / "index.html").write_text(html, encoding="utf-8")
    print(f"Site built in {out}: {len(rows)} videos, {len(channels)} channels, "
          f"index.html {len(html) / 1e3:.0f} KB, {XLSX_NAME} {size_mb:.1f} MB")
    ids = sum(1 for r in rows if r[0])
    if ids < len(rows):
        print(f"Note: {len(rows) - ids} videos have no YouTube ID yet (links fall back to a YouTube search). "
              "Run the updater once to fill them in.")


if __name__ == "__main__":
    main()

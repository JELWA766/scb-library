#!/usr/bin/env python3
"""
Build the public fan website from a library workbook.

  python build_site.py --workbook public_library.xlsx --out site

Produces:
  site/index.html                      single self-contained page (data embedded)
  site/SCB_Library_Fan_Edition.xlsx    the Excel tracker, with NO personal data

Personal data (Mark, Watched On, Notes, Watched Through, descriptions) is stripped, and videos
that are no longer on YouTube (or not yet public) are left out.
"""
import argparse, json, shutil, sys
from datetime import datetime
from pathlib import Path

import scb_library as S

HERE = Path(__file__).resolve().parent
XLSX_NAME = "SCB_Library_Fan_Edition.xlsx"


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
        v.update(mark=None, watched_on=None, notes=None, description="", first_seen=None)
        vids.append(v)
    vids = S.finalize(vids, channels, tz)
    # drop channels with no videos
    used = {v["channel_id"] for v in vids}
    channels = [c for c in channels if c["id"] in used]
    cidx = {c["id"]: i for i, c in enumerate(channels)}

    rows = [[v.get("video_id") or "", v["title"], cidx[v["channel_id"]], v["local"].strftime("%Y-%m-%d"),
             v["local"].strftime("%H:%M"), v["seconds"], v.get("views") or 0] for v in vids if v["channel_id"] in cidx]
    payload = dict(
        generated=now.strftime("%B %-d, %Y") if sys.platform != "win32" else now.strftime("%B %#d, %Y"),
        tz=tzname,
        xlsx=XLSX_NAME,
        channels=[dict(n=c["name"], id=c["id"], c="#" + S.PALETTE[i % len(S.PALETTE)]) for i, c in enumerate(channels)],
        videos=rows)

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    xlsx = out / XLSX_NAME
    S.build(xlsx, vids, channels, tzname, {}, now, fan=True)
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

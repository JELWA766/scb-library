#!/usr/bin/env python3
"""
SCB Video Library - updater / workbook builder.

  python scb_library.py                      update from YouTube (normal use)
  python scb_library.py --offline            rebuild workbook only (no API calls)
  python scb_library.py --import-legacy scb_timeline.xlsx --offline
                                             one-time import of the old export

The workbook (SCB_Library.xlsx) is the database. Your personal data (Mark, Watched On,
Notes, Watched Through dates, settings, channel list) is read back from it, merged with fresh
YouTube data keyed on Video ID, and the whole workbook is regenerated. A timestamped backup is
taken first.
"""
import argparse, os, re, shutil, sys, time
from datetime import datetime, date, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = Path(__file__).resolve().parent
WORKBOOK = HERE / "SCB_Library.xlsx"
BACKUPS = HERE / "backups"
API = "https://www.googleapis.com/youtube/v3"
DEFAULT_TZ = "America/New_York"
MAX_SLOTS = 5          # videos shown per calendar day
VERSION = "1.0"

DEFAULT_CHANNELS = [
    dict(name="Super Carlin Brothers", handle="https://www.youtube.com/@SuperCarlinBrothers", id="UCKZo4N0lVPccBkSiuyVh4yg", yt="SuperCarlinBrothers"),
    dict(name="Super Carlin Gaming", handle="https://www.youtube.com/SuperCarlinGaming", id="UCor_9jRnNNZVzTuQ92PKsaQ", yt="Super Carlin Gaming"),
    dict(name="Jonathan Carlin", handle="https://www.youtube.com/@JonathanCarlin", id="UCYPKFe86e6nm4TVUYKLsFUQ", yt="Jonathan Carlin"),
    dict(name="Through The Griffin Door", handle="https://www.youtube.com/@ThroughTheGriffinDoor", id="UCEscP6ETYXWSSvsNrDHekMQ", yt="Through The Griffin Door"),
    dict(name="Popcorn Culture", handle="https://www.youtube.com/@APOPcast", id="UCHfIbq9thHPC8yrKjAdJgDA", yt="Popcorn Culture"),
    dict(name="With the Carlins", handle="https://www.youtube.com/@withthecarlins9970", id="UCVBTjN9WJU5Ac8yU7b2K-5g", yt="With the Carlins"),
]

# ----------------------------------------------------------------------------- styling
FONT = "Arial"
NAVY, SLATE, TEAL, GOLD = "1F2A44", "3B4A63", "0F7B8A", "B7791F"
PALETTE = ["3B6FB6", "2E9E6B", "E08A1E", "8E5BB5", "D1495B", "1F9EA8",
           "C9A227", "D96CA3", "8C6A4F", "5C6B7A", "7DA83A", "4A4AC8"]


def tint(h, f=0.80):
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "%02X%02X%02X" % tuple(int(c + (255 - c) * f) for c in (r, g, b))


def fnt(size=10, bold=False, color="000000", italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(h):
    return PatternFill("solid", start_color=h, end_color=h)


THIN = Side(style="thin", color="D5D9E2")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def put(ws, ref, value, font=None, fl=None, al=None, fmt=None, border=None, unlock=False):
    c = ws[ref]
    if isinstance(value, str):
        value = ILLEGAL.sub("", value)[:32000]
        c.value = value
        if value.startswith("=") and not getattr(put, "formula", False):
            pass
    else:
        c.value = value
    if font: c.font = font
    if fl: c.fill = fill(fl)
    if al: c.alignment = al
    if fmt: c.number_format = fmt
    if border: c.border = border
    if unlock: c.protection = Protection(locked=False)
    return c


def text(ws, ref, value, **kw):
    """Write literal text (even if it starts with '=')."""
    c = put(ws, ref, value, **kw)
    if isinstance(value, str) and value.startswith("="):
        c.data_type = "s"
    return c


def f(ws, ref, formula, **kw):
    return put(ws, ref, formula, **kw)


def title_block(ws, title, subtitle=None, width_to="M"):
    ws.sheet_view.showGridLines = False
    put(ws, "B2", title, font=fnt(20, True, NAVY))
    if subtitle:
        put(ws, "B3", subtitle, font=fnt(9, False, "6B7385", italic=True))
    ws.row_dimensions[2].height = 30


def nav(ws, ref, sheet, label, anchor="A1"):
    f(ws, ref, f'=HYPERLINK("#{sheet}!{anchor}","{label}")', font=fnt(9, True, TEAL, underline="single"))


def section(ws, row, label, c1="B", c2="M"):
    for ci in range(ws[c1 + "1"].column, ws[c2 + "1"].column + 1):
        ws.cell(row=row, column=ci).fill = fill(NAVY)
    put(ws, f"{c1}{row}", label, font=fnt(11, True, "FFFFFF"), fl=NAVY, al=LEFT)
    ws.row_dimensions[row].height = 20


def head(ws, row, col0, labels, fl=SLATE, height=32):
    for i, lab in enumerate(labels):
        put(ws, f"{CL(col0 + i)}{row}", lab, font=fnt(9, True, "FFFFFF"), fl=fl, al=CENTER, border=BOX)
    ws.row_dimensions[row].height = height


def protect(ws):
    ws.protection.sheet = True


# ----------------------------------------------------------------------------- library schema
LIB = ["Open", "Status", "Mark", "Watched On", "Published", "Channel", "Title", "Length", "Views",
       "Likes", "Comments", "Notes", "Video ID", "Channel ID", "Published UTC", "Day", "Seq",
       "Availability", "First Seen", "Last Seen", "Description", "Cal Key", "Queue Key"]
LC = {n: CL(i + 1) for i, n in enumerate(LIB)}
HDR = 5
MARKS = ["Watched", "Unwatched", "Watch Later", "Skipped"]
CH_HDR = 5   # Channels header row; channel i is on row CH_HDR + i


HM_FMT = '[h]" h "mm" m"'
HM_Z = '[h]" h "mm" m";;"·"'


def hm(expr):
    """Excel text expression: a duration (in days) shown as '2 h 15 m'."""
    r = f"ROUND(({expr})*1440,0)"
    return f'IF({r}<60,{r}&" m",IF(MOD({r},60)=0,INT({r}/60)&" h",INT({r}/60)&" h "&MOD({r},60)&" m"))'


def tidy(ch, legend=True, left=0.07, width=0.90):
    """Give a chart explicit room for title, axis labels and legend so they never overlap."""
    from openpyxl.chart.layout import Layout, ManualLayout
    if ch.title is not None:
        ch.title.overlay = False
    ch.plot_area.layout = Layout(manualLayout=ManualLayout(
        layoutTarget="inner", xMode="edge", yMode="edge", x=left, y=0.17, w=width, h=0.48 if legend else 0.63))
    if legend and ch.legend is not None:
        ch.legend.position = "b"; ch.legend.overlay = False
        ch.legend.layout = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.03, y=0.87, w=0.94, h=0.11))


def V(col):
    return f"tblVideos[{col}]"


# ----------------------------------------------------------------------------- helpers: time
def get_tz(name):
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, Exception):
        print(f"!! Unknown time zone '{name}' - using {DEFAULT_TZ}. (On Windows: pip install tzdata)")
        return ZoneInfo(DEFAULT_TZ)


def to_local(utc_naive, tz):
    return utc_naive.replace(tzinfo=timezone.utc).astimezone(tz).replace(tzinfo=None)


def parse_iso_duration(s):
    m = re.fullmatch(r"P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?", s or "")
    if not m:
        return 0
    d, h, mi, se = (int(x or 0) for x in m.groups())
    return d * 86400 + h * 3600 + mi * 60 + se


# ----------------------------------------------------------------------------- YouTube API
class ApiError(RuntimeError):
    pass


class YT:
    def __init__(self, key):
        import requests
        self.s, self.key, self.calls = requests.Session(), key, 0

    def get(self, endpoint, **params):
        params["key"] = self.key
        for attempt in range(5):
            self.calls += 1
            try:
                r = self.s.get(f"{API}/{endpoint}", params=params, timeout=30)
            except Exception as e:           # network blip
                if attempt == 4:
                    raise ApiError(f"Network error: {e}")
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 200:
                return r.json()
            if r.status_code in (500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            try:
                err = r.json()["error"]
                reason = err["errors"][0].get("reason", "")
                msg = err.get("message", "")
            except Exception:
                reason, msg = "", r.text[:200]
            if reason in ("quotaExceeded", "dailyLimitExceeded", "rateLimitExceeded"):
                raise ApiError("YouTube API quota exhausted for today. Try again tomorrow.")
            if reason in ("keyInvalid", "forbidden", "accessNotConfigured"):
                raise ApiError(f"API key problem ({reason}): {msg}")
            raise ApiError(f"{endpoint} failed ({r.status_code} {reason}): {msg}")
        raise ApiError(f"{endpoint}: YouTube kept returning server errors.")


def load_api_key():
    k = os.environ.get("YT_API_KEY", "").strip()
    p = HERE / "api_key.txt"
    if not k and p.exists():
        k = p.read_text().strip()
    if not k:
        sys.exit("No API key found. Put your YouTube Data API key in api_key.txt next to this script "
                 "(or set the YT_API_KEY environment variable).")
    return k


def resolve_channel(yt, text_):
    t = (text_ or "").strip()
    m = re.search(r"(UC[0-9A-Za-z_-]{22})", t)
    if m:
        return m.group(1)
    m = re.search(r"@([\w.\-]+)", t)
    handle = m.group(1) if m else None
    if handle:
        d = yt.get("channels", part="id", forHandle="@" + handle)
        if d.get("items"):
            return d["items"][0]["id"]
    m = re.search(r"/user/([\w\-]+)", t)
    if m:
        d = yt.get("channels", part="id", forUsername=m.group(1))
        if d.get("items"):
            return d["items"][0]["id"]
    bare = t.rstrip("/").split("/")[-1].lstrip("@")
    if bare:
        d = yt.get("channels", part="id", forHandle="@" + bare)
        if d.get("items"):
            return d["items"][0]["id"]
        d = yt.get("search", part="snippet", type="channel", q=bare, maxResults=1)   # 100 quota units
        if d.get("items"):
            return d["items"][0]["snippet"]["channelId"]
    return None


def fetch_channel_info(yt, ids):
    info = {}
    for i in range(0, len(ids), 50):
        d = yt.get("channels", part="snippet,contentDetails", id=",".join(ids[i:i + 50]))
        for it in d.get("items", []):
            info[it["id"]] = dict(title=it["snippet"]["title"],
                                  uploads=it["contentDetails"]["relatedPlaylists"]["uploads"])
    return info


def fetch_uploads(yt, playlist):
    ids, token = [], None
    while True:
        p = dict(part="contentDetails", playlistId=playlist, maxResults=50)
        if token:
            p["pageToken"] = token
        d = yt.get("playlistItems", **p)
        ids += [it["contentDetails"]["videoId"] for it in d.get("items", [])]
        token = d.get("nextPageToken")
        if not token:
            return ids


def fetch_details(yt, ids):
    out = {}
    for i in range(0, len(ids), 50):
        d = yt.get("videos", part="snippet,contentDetails,statistics", id=",".join(ids[i:i + 50]))
        for it in d.get("items", []):
            sn, st = it["snippet"], it.get("statistics", {})
            pub = datetime.strptime(sn["publishedAt"][:19], "%Y-%m-%dT%H:%M:%S")
            out[it["id"]] = dict(
                video_id=it["id"], channel_id=sn["channelId"], channel_title=sn["channelTitle"],
                title=sn["title"], description=sn.get("description", ""), published_utc=pub,
                seconds=parse_iso_duration(it["contentDetails"].get("duration")),
                views=int(st["viewCount"]) if "viewCount" in st else None,
                likes=int(st["likeCount"]) if "likeCount" in st else None,
                comments=int(st["commentCount"]) if "commentCount" in st else None,
                live=sn.get("liveBroadcastContent", "none"))
    return out


# ----------------------------------------------------------------------------- read workbook
def find_header(ws, must, rows=15):
    for r in range(1, rows + 1):
        vals = {c.value: c.column for c in ws[r] if isinstance(c.value, str)}
        if all(m in vals for m in must):
            return r, vals
    return None, None


SELECTORS = {"Calendar": ["C4", "E4", "G4"], "Queue": ["C4", "D4", "E4"], "History": ["C4", "G4", "K4"]}


def length_to_seconds(x):
    """Length cells come back from openpyxl as timedelta, time, datetime or a day-fraction float."""
    if x is None or x == "":
        return 0
    if isinstance(x, timedelta):
        return int(round(x.total_seconds()))
    if isinstance(x, datetime):
        return int(round((x - datetime(1899, 12, 30)).total_seconds()))
    if hasattr(x, "hour"):
        return x.hour * 3600 + x.minute * 60 + x.second
    return int(round(float(x) * 86400))


def read_workbook(path):
    wb = load_workbook(path)
    out = dict(settings={}, channels=[], videos=[], selectors={})
    if "Settings" in wb.sheetnames:
        out["settings"]["tz"] = wb["Settings"]["C4"].value
    for sh, cells in SELECTORS.items():
        if sh in wb.sheetnames:
            for c in cells:
                out["selectors"][(sh, c)] = wb[sh][c].value
    if "Channels" in wb.sheetnames:
        ws = wb["Channels"]
        hr, cols = find_header(ws, ["Display Name", "Channel ID"])
        if hr:
            for r in range(hr + 1, ws.max_row + 1):
                g = lambda n: ws.cell(row=r, column=cols[n]).value if n in cols else None
                if not (g("Display Name") or g("Channel ID") or g("Handle or URL")):
                    continue
                out["channels"].append(dict(
                    name=g("Display Name"), handle=g("Handle or URL"), id=g("Channel ID"),
                    enabled=str(g("Enabled") or "Yes").strip().lower() not in ("no", "n", "false", "0"),
                    through=g("Watched Through"), yt=g("YouTube Title"), refreshed=g("Last Refresh")))
    if "Library" in wb.sheetnames:
        ws = wb["Library"]
        hr, cols = find_header(ws, ["Title", "Video ID"])
        if hr:
            for r in range(hr + 1, ws.max_row + 1):
                g = lambda n: ws.cell(row=r, column=cols[n]).value if n in cols else None
                if g("Title") is None and g("Video ID") is None:
                    continue
                out["videos"].append(dict(
                    video_id=g("Video ID") or None, channel_id=g("Channel ID"), channel_title=None,
                    title=g("Title") or "", description=g("Description") or "",
                    published_utc=g("Published UTC"),
                    seconds=length_to_seconds(g("Length")),
                    views=g("Views"), likes=g("Likes"), comments=g("Comments"),
                    mark=g("Mark"), watched_on=g("Watched On"), notes=g("Notes"),
                    first_seen=g("First Seen"), last_seen=g("Last Seen"), live="none"))
    return out


def read_legacy(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    rows = wb.active.iter_rows(values_only=True)
    hdr = list(next(rows))
    ix = {h: i for i, h in enumerate(hdr)}
    vids = []
    for row in rows:
        if row[ix["Title"]] is None:
            continue
        m, s = str(row[ix["Duration"]]).split(":")
        vids.append(dict(
            video_id=None, channel_id=None, channel_title=row[ix["Channel"]], title=row[ix["Title"]],
            description=row[ix["Description"]] or "", published_utc=row[ix["Published"]],
            seconds=int(m) * 60 + int(s), views=row[ix["Views"]], likes=row[ix["Likes"]],
            comments=row[ix["Comments"]], mark=None, watched_on=None, notes=None,
            first_seen=None, last_seen=None, live="none"))
    return vids


# ----------------------------------------------------------------------------- merge
def minute(dt):
    return dt.replace(second=0, microsecond=0) if dt else None


def merge(existing, fetched, fetched_channels, now):
    """existing: list of records (some without video_id). fetched: {vid: rec}."""
    by_id = {v["video_id"]: v for v in existing if v.get("video_id")}
    legacy = [v for v in existing if not v.get("video_id")]
    stats = dict(new=0, updated=0, adopted=0, missing=0)

    # Fallback matching for rows that have no Video ID yet (imported history). Three passes, each
    # only accepting a match that is unambiguous on both sides.
    adopt, used = {}, set()
    keyfns = [lambda r: (r.get("channel_id"), minute(r["published_utc"]), r["title"]),
              lambda r: (r.get("channel_id"), minute(r["published_utc"])),
              lambda r: (minute(r["published_utc"]), r["title"])]
    for kf in keyfns:
        old_idx, new_idx = {}, {}
        for v in legacy:
            if id(v) not in used:
                old_idx.setdefault(kf(v), []).append(v)
        for vid, n in fetched.items():
            if vid not in by_id and vid not in adopt:
                new_idx.setdefault(kf(n), []).append(vid)
        for key, vids in new_idx.items():
            olds = old_idx.get(key, [])
            if len(vids) == len(olds) and (len(olds) == 1 or kf is keyfns[0]):
                for vid, o in zip(vids, olds):
                    adopt[vid] = o; used.add(id(o))

    consumed = set(id(x) for x in adopt.values())
    result = []
    for vid, n in fetched.items():
        old = by_id.get(vid) or adopt.get(vid)
        if old is not None:
            consumed.add(id(old))
        rec = dict(old) if old else dict(mark=None, watched_on=None, notes=None, first_seen=now)
        if vid in by_id:
            stats["updated"] += 1
        elif old:
            stats["adopted"] += 1
        else:
            stats["new"] += 1
        rec.update(n)
        rec["last_seen"] = now
        rec["availability"] = "Upcoming/Live" if n["live"] in ("upcoming", "live") else "OK"
        result.append(rec)
    for v in existing:
        if id(v) in consumed:
            continue
        keep = dict(v)
        if v.get("channel_id") in fetched_channels:
            keep["availability"] = "Not in last refresh"
            stats["missing"] += 1
        else:
            keep["availability"] = v.get("availability") or "OK"
        result.append(keep)
    return result, stats


# ----------------------------------------------------------------------------- finalize
def finalize(videos, channels, tz):
    # 0:00 entries are live-stream / premiere placeholders that never became real videos
    videos[:] = [v for v in videos if (v.get("seconds") or 0) > 0]
    name_by_id = {c["id"]: c["name"] for c in channels}
    yt_to_id = {c.get("yt"): c["id"] for c in channels if c.get("yt")}
    for v in videos:
        if not v.get("channel_id") and v.get("channel_title") in yt_to_id:
            v["channel_id"] = yt_to_id[v["channel_title"]]
        v["channel"] = name_by_id.get(v.get("channel_id"), v.get("channel_title") or "Unknown")
        v["local"] = to_local(v["published_utc"], tz)
        v["day"] = v["local"].replace(hour=0, minute=0, second=0, microsecond=0)
    videos.sort(key=lambda v: (v["published_utc"], v["title"], v.get("video_id") or ""))
    for i, v in enumerate(videos, 1):
        v["seq"] = i
        v.setdefault("availability", "OK")
        if not v.get("video_id"):
            v["availability"] = v.get("availability") if v.get("availability") != "OK" else "No Video ID yet"
    return videos


# ----------------------------------------------------------------------------- BUILD
def build(path, videos, channels, tzname, selectors, now, fan=False):
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    n = len(channels)
    N = len(videos)
    colors = [PALETTE[i % len(PALETTE)] for i in range(n)]
    cidx = {c["id"]: i for i, c in enumerate(channels)}

    names = ["Dashboard", "Library", "Queue", "Calendar", "History", "Stats", "Channels"] + \
            (["Carry Over"] if fan else []) + ["Settings", "Lists"]
    ws_d = wb.active; ws_d.title = "Dashboard"
    W = {nm: (ws_d if nm == "Dashboard" else wb.create_sheet(nm)) for nm in names}
    tabs = dict(Dashboard=NAVY, Library=TEAL, Queue=TEAL, Calendar="2E9E6B", History="2E9E6B", Stats="2E9E6B",
                Channels=GOLD, Settings=GOLD)
    for k, c in tabs.items():
        W[k].sheet_properties.tabColor = c
    W["Lists"].sheet_state = "hidden"

    years = list(range(videos[0]["local"].year, videos[-1]["local"].year + 1)) if videos else [now.year]
    MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
              "October", "November", "December"]

    # ---------------- Lists
    L = W["Lists"]
    for i, y in enumerate(years, 1):
        L.cell(row=i, column=1, value=y)
    for i, m in enumerate(MONTHS, 1):
        L.cell(row=i, column=3, value=m)
    L["E1"] = "All channels"
    for i in range(n):
        L.cell(row=i + 2, column=5, value=f'=IF(Channels!$C${CH_HDR + 1 + i}="","",Channels!$C${CH_HDR + 1 + i})')
    wb.defined_names["YearList"] = DefinedName("YearList", attr_text=f"Lists!$A$1:$A${len(years)}")
    wb.defined_names["MonthList"] = DefinedName("MonthList", attr_text="Lists!$C$1:$C$12")
    wb.defined_names["ChannelList"] = DefinedName(
        "ChannelList", attr_text='OFFSET(Lists!$E$1,0,0,COUNTIF(Lists!$E$1:$E$40,"?*"),1)')

    def dv_list(ws, ref, formula):
        dv = DataValidation(type="list", formula1=formula, allow_blank=False, showErrorMessage=True)
        ws.add_data_validation(dv)
        dv.add(ref)

    def selector(ws, ref, value, label_ref=None, label=None, merge_to=None):
        if label_ref:
            put(ws, label_ref, label, font=fnt(9, True, "6B7385"))
        if merge_to:
            ws.merge_cells(f"{ref}:{merge_to}")
        put(ws, ref, value, font=fnt(11, True, NAVY), fl="FFF8E1", al=CENTER, border=BOX, unlock=True)

    def sel(sheet, cell, default):
        v = selectors.get((sheet, cell))
        return default if v in (None, "") else v

    chan_names = ["All channels"] + [c["name"] for c in channels]
    latest = videos[-1]["local"] if videos else now

    def crit_formula(cell_sel):
        return (f'=IF({cell_sel}="All channels","*",IFERROR(INDEX(tblChannels[Channel ID],'
                f'MATCH({cell_sel},tblChannels[Display Name],0)),"none"))')

    def chan_mask(crit):
        return f'((tblVideos[Channel ID]={crit})+({crit}="*")>0)'

    # =============================================================== CHANNELS
    ws = W["Channels"]
    ws.sheet_view.showGridLines = False
    title_block(ws, "Channels", "Add a channel: type its @handle or URL in a new row below the table, then run the updater. "
                                 "Set 'Watched Through' to treat everything up to that date as already watched.")
    nav(ws, "B4", "Dashboard", "◄ Dashboard")
    cols = ["Color", "Enabled", "Display Name", "Handle or URL", "Channel ID", "Watched Through", "YouTube Title",
            "Videos", "Watched", "Backlog", "Latest Upload", "Last Refresh"]
    head(ws, CH_HDR, 1, cols)
    for i, c in enumerate(channels):
        r = CH_HDR + 1 + i
        put(ws, f"A{r}", "", fl=colors[i], border=BOX)
        put(ws, f"B{r}", "Yes" if c["enabled"] else "No", al=CENTER, border=BOX, unlock=True)
        put(ws, f"C{r}", c["name"], font=fnt(10, True), border=BOX)
        put(ws, f"D{r}", c.get("handle"), border=BOX)
        put(ws, f"E{r}", c["id"], font=fnt(9, color="6B7385"), border=BOX)
        put(ws, f"F{r}", c.get("through"), fl="FFF8E1", fmt="yyyy-mm-dd", al=CENTER, border=BOX)
        put(ws, f"G{r}", c.get("yt"), font=fnt(9, color="6B7385"), border=BOX)
        f(ws, f"H{r}", f"=COUNTIFS({V('Channel ID')},E{r})", fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"I{r}", f'=COUNTIFS({V("Channel ID")},E{r},{V("Status")},"Watched")', fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"J{r}", f'=COUNTIFS({V("Channel ID")},E{r},{V("Status")},"Unwatched")+COUNTIFS({V("Channel ID")},E{r},{V("Status")},"Watch Later")',
          fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"K{r}", f'=IF(H{r}=0,"",_xlfn.MAXIFS({V("Day")},{V("Channel ID")},E{r}))', fmt="yyyy-mm-dd", al=CENTER, border=BOX)
        put(ws, f"L{r}", c.get("refreshed"), fmt="yyyy-mm-dd hh:mm", font=fnt(9, color="6B7385"), al=CENTER, border=BOX)
    last_ch = CH_HDR + max(n, 1)
    t = Table(displayName="tblChannels", ref=f"A{CH_HDR}:L{last_ch}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(t)
    for col, w in zip("ABCDEFGHIJKL", [7, 9, 28, 44, 28, 17, 26, 10, 10, 10, 15, 18]):
        ws.column_dimensions[col].width = w
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"B{CH_HDR + 1}:B{CH_HDR + 40}")
    dv2 = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                         errorTitle="Date needed", error="Enter a date such as 2019-12-31.")
    ws.add_data_validation(dv2); dv2.add(f"F{CH_HDR + 1}:F{CH_HDR + 40}")
    ws.freeze_panes = f"A{CH_HDR + 1}"
    r0 = last_ch + 3
    for i, line in enumerate([
        "How these columns work",
        "• Enabled - 'No' stops the updater fetching that channel (its existing videos stay in the library).",
        "• Display Name - what you see everywhere in the workbook. Rename freely; run the updater to refresh labels in the Library.",
        "• Handle or URL / Channel ID - give either. If Channel ID is blank, the updater looks it up for you.",
        "• Watched Through - handy for catching up: every video on or before this date counts as Watched (unless you override it in Library > Mark).",
        "• YouTube Title - the channel's real name on YouTube, filled by the updater so you can confirm the right channel was matched.",
        "• Colors follow row order. Please don't sort this table by hand."]):
        put(ws, f"A{r0 + i}", line, font=fnt(10, i == 0, NAVY if i == 0 else "444444"))

    # =============================================================== LIBRARY
    ws = W["Library"]
    ws.sheet_view.showGridLines = False
    put(ws, "A1", "Video Library", font=fnt(18, True, NAVY)); ws.row_dimensions[1].height = 28
    put(ws, "A2", "Slate columns come from YouTube and are refreshed by the updater.  Teal columns are calculated.  "
                  "Gold columns (Mark, Watched On, Notes) are YOURS and are never overwritten.", font=fnt(9, False, "6B7385", italic=True))
    nav(ws, "A3", "Dashboard", "◄ Dashboard")
    kind = {}
    for nm in LIB:
        kind[nm] = "user" if nm in ("Mark", "Watched On", "Notes") else ("calc" if nm in ("Open", "Status", "Cal Key", "Queue Key") else "raw")
    for i, nm in enumerate(LIB, 1):
        put(ws, f"{CL(i)}{HDR}", nm, font=fnt(9, True, "FFFFFF"),
            fl={"user": GOLD, "calc": TEAL, "raw": SLATE}[kind[nm]], al=CENTER, border=BOX)
    ws.row_dimensions[HDR].height = 24
    tmp_cut = "IFERROR(INDEX(tblChannels[Watched Through],MATCH({N}{r},tblChannels[Channel ID],0)),0)"
    order = sorted(videos, key=lambda v: -v["seq"])
    last = HDR + max(N, 1)
    ab = lambda c: f"${LC[c]}${HDR + 1}:${LC[c]}${last}"
    tint_by = {c["id"]: tint(colors[i]) for i, c in enumerate(channels)}
    for k, v in enumerate(order):
        r = HDR + 1 + k
        f(ws, f"A{r}", f'=IF({LC["Video ID"]}{r}="","",HYPERLINK("https://www.youtube.com/watch?v="&{LC["Video ID"]}{r},"▶"))',
          font=fnt(10, True, "1155CC"), al=CENTER)
        f(ws, f"B{r}", f'=IF(C{r}<>"",C{r},IF({LC["Day"]}{r}<=' + tmp_cut.format(N=LC["Channel ID"], r=r) + ',"Watched","Unwatched"))',
          al=CENTER)
        put(ws, f"C{r}", v.get("mark") or None, fl="FFF8E1", al=CENTER, unlock=True)
        put(ws, f"D{r}", v.get("watched_on"), fl="FFF8E1", fmt="yyyy-mm-dd", al=CENTER, unlock=True)
        put(ws, f"E{r}", v["local"], fmt="yyyy-mm-dd  h:mm AM/PM", al=CENTER)
        put(ws, f"F{r}", v["channel"], fl=tint_by.get(v.get("channel_id"), "EEEEEE"))
        text(ws, f"G{r}", v["title"])
        put(ws, f"H{r}", v["seconds"] / 86400, fmt="[h]:mm:ss", al=CENTER)
        put(ws, f"I{r}", v.get("views"), fmt="#,##0")
        put(ws, f"J{r}", v.get("likes"), fmt="#,##0")
        put(ws, f"K{r}", v.get("comments"), fmt="#,##0")
        text(ws, f"L{r}", v.get("notes") or None, fl="FFF8E1", unlock=True)
        put(ws, f"M{r}", v.get("video_id"), font=fnt(9, color="6B7385"))
        put(ws, f"N{r}", v.get("channel_id"), font=fnt(9, color="6B7385"))
        put(ws, f"O{r}", v["published_utc"], fmt="yyyy-mm-dd hh:mm:ss", font=fnt(9, color="6B7385"))
        put(ws, f"P{r}", v["day"], fmt="yyyy-mm-dd", font=fnt(9, color="6B7385"))
        put(ws, f"Q{r}", v["seq"], font=fnt(9, color="6B7385"))
        put(ws, f"R{r}", v.get("availability", "OK"), font=fnt(9, color="6B7385"))
        put(ws, f"S{r}", v.get("first_seen"), fmt="yyyy-mm-dd", font=fnt(9, color="6B7385"))
        put(ws, f"T{r}", v.get("last_seen"), fmt="yyyy-mm-dd", font=fnt(9, color="6B7385"))
        text(ws, f"U{r}", (v.get("description") or "") or None, font=fnt(9, color="6B7385"))
        f(ws, f"V{r}", f'=IF(AND(P{r}>=Calendar!$K$2,P{r}<EDATE(Calendar!$K$2,1),OR(Calendar!$K$4="*",N{r}=Calendar!$K$4)),'
                       f'P{r}*100+COUNTIFS({ab("Day")},P{r},{ab("Seq")},"<="&Q{r},{ab("Channel ID")},Calendar!$K$4),"")', font=fnt(9, color="6B7385"))
        f(ws, f"W{r}", f'=IF(AND(B{r}=Queue!$C$4,OR(Queue!$K$4="*",N{r}=Queue!$K$4)),Q{r},"")', font=fnt(9, color="6B7385"))
    t = Table(displayName="tblVideos", ref=f"A{HDR}:{CL(len(LIB))}{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    for nm, w in zip(LIB, [7, 12, 13, 13, 21, 24, 70, 10, 11, 9, 10, 32, 13, 14, 17, 12, 7, 17, 12, 12, 60, 10, 10]):
        ws.column_dimensions[LC[nm]].width = w
    ws.column_dimensions.group("M", "T", hidden=True)
    ws.column_dimensions.group("V", "W", hidden=True)
    ws.freeze_panes = f"A{HDR + 1}"
    rng = lambda c: f"{LC[c]}{HDR + 1}:{LC[c]}{last}"
    dv = DataValidation(type="list", formula1='"' + ",".join(MARKS) + '"', allow_blank=True, showErrorMessage=True,
                        errorTitle="Pick from the list", error="Use: " + ", ".join(MARKS) + " (or clear the cell).")
    ws.add_data_validation(dv); dv.add(rng("Mark"))
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                        errorTitle="Date needed", error="Enter a date, e.g. 2026-10-02 (Ctrl+; inserts today).")
    ws.add_data_validation(dv); dv.add(rng("Watched On"))
    s = rng("Status")
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Watched"'], fill=fill("DDF2E2"), font=Font(color="1E6B34")))
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Watch Later"'], fill=fill("FFEFC2"), font=Font(color="8A5A00")))
    ws.conditional_formatting.add(s, CellIsRule(operator="equal", formula=['"Skipped"'], fill=fill("E6E8EC"), font=Font(color="7A8294")))
    ws.conditional_formatting.add(f"G{HDR + 1}:G{last}", FormulaRule(formula=[f'$B{HDR + 1}="Watched"'], font=Font(color="7A8294")))
    ws.conditional_formatting.add(rng("Availability"), CellIsRule(operator="equal", formula=['"Not in last refresh"'], font=Font(color="C0392B", bold=True)))

    # =============================================================== SETTINGS
    ws = W["Settings"]
    title_block(ws, "Settings & Guide")
    nav(ws, "B3", "Dashboard", "◄ Dashboard")
    put(ws, "B4", "Time zone (IANA name)", font=fnt(10, True))
    put(ws, "C4", tzname, fl="FFF8E1", border=BOX, unlock=True)
    put(ws, "D4", "Used to decide which calendar day a video belongs to. Changing it takes effect at the next update.", font=fnt(9, color="6B7385", italic=True))
    put(ws, "B5", "Last updated", font=fnt(10, True))
    put(ws, "C5", now, fmt="yyyy-mm-dd h:mm AM/PM", border=BOX, al=LEFT)
    put(ws, "B6", "Videos in library", font=fnt(10, True))
    f(ws, "C6", f"=COUNTA({V('Seq')})", fmt="#,##0", border=BOX, al=LEFT)
    put(ws, "B7", "Workbook version", font=fnt(10, True)); put(ws, "C7", VERSION, border=BOX, al=LEFT)
    guide = [
        ("Using the workbook", True),
        ("Dashboard - the home page: progress, per-channel breakdown, up-next list and newest uploads.", False),
        ("Library - every video. Filter with the header arrows. Use Mark (dropdown) to set Watched / Watch Later / Skipped, and Watched On for the date (Ctrl+; = today).", False),
        ("   Marking a lot at once: filter, type the mark in the first visible cell, copy it, and paste over the rest of the visible cells.", False),
        ("   Catching up on a back catalogue: set 'Watched Through' for a channel on the Channels sheet instead of marking thousands of rows.", False),
        ("   'Skipped' videos (trailers, clips you'll never watch) are left out of the percentage calculations.", False),
        ("Queue - pick a channel / list / order to get the next 30 videos to watch, each with a link.", False),
        ("Calendar - pick a year, month and channel to see that month's uploads on a Monday-first calendar. Click a title to open it.", False),
        ("History - upload heatmap (year x month), year view, uploads per year and a month-by-month timeline with charts.", False),
        ("Stats - progress per channel, upload profile, watching-vs-uploading pace and records.", False),
        ("Sheets other than Library / Channels / Settings are protected (no password) so formulas aren't overwritten by accident: Review > Unprotect Sheet.", False),
        ("", False),
        *([("About this copy", True),
           ("This is a snapshot taken from the fan website. It does not update itself - download a fresh copy from the site for newer videos (your marks would not carry over).", False),
           ("Your marks live in the gold columns of the Library sheet and in 'Watched Through' on the Channels sheet. Keep a backup copy of the file.", False),
           ("Got a newer download? Use the 'Carry Over' sheet in the NEW file to bring your marks across from the old one (steps are on that sheet).", False)] if fan else
          [("Updating from YouTube", True),
           ("1. Close this workbook in Excel.   2. Double-click 'Update Library.bat' (or run: python scb_library.py).   3. Re-open the workbook.", False),
           ("Every update backs up the previous workbook to the 'backups' folder, then re-pulls all channels, adds new videos, refreshes views/likes, and keeps your marks.", False),
           ("Please don't rearrange or rename columns/sheets by hand - the workbook is regenerated on each update. Your data in the gold columns, Channels and Settings is carried over.", False)]),
        ("", False),
        ("Statistics caveats", True),
        ("• Watch-activity charts only count videos that have a 'Watched On' date. Videos counted as watched via 'Watched Through' or without a date are not in them.", False),
        ("• View counts grow with age, so old videos vs new ones isn't a like-for-like comparison.", False),
    ]
    for i, (line, b) in enumerate(guide):
        put(ws, f"B{10 + i}", line, font=fnt(10, b, NAVY if b else "333333"))
    ws.column_dimensions["A"].width = 3; ws.column_dimensions["B"].width = 34; ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 90

    # =============================================================== STATS (built before History refs; needs History rows)
    # --- History layout first (row numbers needed by Stats)
    H = W["History"]
    ws = H
    title_block(ws, "History")
    put(ws, "B5", "How the channels uploaded over the years - and how much of it you've watched.", font=fnt(9, False, "6B7385", italic=True))
    nav(ws, "B4", "Dashboard", "◄ Dashboard")
    H_ROWS = {}
    selector(ws, "C4", sel("History", "C4", "All channels"), "C3", "Channel", merge_to="E4")
    selector(ws, "G4", sel("History", "G4", "Videos"), "G3", "Heatmap shows", merge_to="H4")
    selector(ws, "K4", sel("History", "K4", years[-1]), "K3", "Year view", merge_to="L4")
    ws["B4"].alignment = LEFT
    dv_list(ws, "C4", "=ChannelList"); dv_list(ws, "G4", '"Videos,Hours"'); dv_list(ws, "K4", "=YearList")
    f(ws, "T4", crit_formula("$C$4"))     # helper: channel criterion
    ws.column_dimensions.group("T", "U", hidden=True)
    ws.column_dimensions["A"].width = 2; ws.column_dimensions["B"].width = 16
    for ci in range(3, 20):
        ws.column_dimensions[CL(ci)].width = 10.5

    row = 6
    section(ws, row, "Upload heatmap  -  darker = more (hours are rounded)", "B", "O")
    hh = row + 1
    head(ws, hh, 2, ["Year"] + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"] + ["Total"], height=20)
    hm0 = hh + 1
    for yi, y in enumerate(years):
        r = hm0 + yi
        put(ws, f"B{r}", y, font=fnt(10, True), al=CENTER, border=BOX)
        for m in range(1, 13):
            c = CL(2 + m)
            lo = f'">="&DATE($B{r},{m},1)'; hi = f'"<"&DATE($B{r},{m + 1},1)'
            cond = f'{V("Day")},{lo},{V("Day")},{hi},{V("Channel ID")},$T$4'
            f(ws, f"{c}{r}", f'=IF($G$4="Hours",ROUND(SUMIFS({V("Length")},{cond})*24,0),COUNTIFS({cond}))',
              fmt='#,##0;-#,##0;"·"', al=CENTER, border=BOX)
        f(ws, f"O{r}", f"=SUM(C{r}:N{r})", fmt="#,##0", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    hm1 = hm0 + len(years) - 1
    r = hm1 + 1
    put(ws, f"B{r}", "All years", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    for ci in range(3, 16):
        c = CL(ci)
        f(ws, f"{c}{r}", f"=SUM({c}{hm0}:{c}{hm1})", fmt="#,##0", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    ws.conditional_formatting.add(f"C{hm0}:N{hm1}", ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                                                  end_type="max", end_color="0F7B8A"))
    row = r + 3

    # ---- Year view
    section(ws, row, "Year view", "B", "O")
    f(ws, f"B{row}", '="Year view  -  "&$K$4', font=fnt(11, True, "FFFFFF"), fl=NAVY, al=LEFT)
    yh = row + 1
    head(ws, yh, 2, ["Month"] + [None] * n + ["Total", "Time"], height=46)
    for i in range(n):
        f(ws, f"{CL(3 + i)}{yh}", f"=Channels!$C${CH_HDR + 1 + i}", font=fnt(9, True, "FFFFFF"), fl=colors[i], al=CENTER, border=BOX)
    for m in range(1, 13):
        r = yh + m
        put(ws, f"B{r}", MONTHS[m - 1], font=fnt(10, True), border=BOX)
        for i in range(n):
            f(ws, f"{CL(3 + i)}{r}",
              f'=COUNTIFS({V("Day")},">="&DATE($K$4,{m},1),{V("Day")},"<"&DATE($K$4,{m + 1},1),{V("Channel ID")},Channels!$E${CH_HDR + 1 + i})',
              fmt='#,##0;-#,##0;"·"', al=CENTER, border=BOX)
        f(ws, f"{CL(3 + n)}{r}", f"=SUM(C{r}:{CL(2 + n)}{r})", fmt='#,##0;-#,##0;"·"', font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
        f(ws, f"{CL(4 + n)}{r}",
          f'=SUMIFS({V("Length")},{V("Day")},">="&DATE($K$4,{m},1),{V("Day")},"<"&DATE($K$4,{m + 1},1))',
          fmt=HM_Z, al=CENTER, border=BOX)
    r = yh + 13
    put(ws, f"B{r}", "Year total", font=fnt(10, True), border=BOX, fl="EEF0F5")
    for ci in range(3, 5 + n):
        c = CL(ci)
        f(ws, f"{c}{r}", f"=SUM({c}{yh + 1}:{c}{yh + 12})", fmt=HM_FMT if ci == 4 + n else "#,##0", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")

    def stacked(ws, hdr_row, first_row, last_row, col1, col2, cat_col, anchor, title, w=26, h=9.5, cols=None, ytitle=None):
        ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100; ch.gapWidth = 40
        ch.title = title; ch.width, ch.height = w, h
        ch.add_data(Reference(ws, min_col=col1, max_col=col2, min_row=hdr_row, max_row=last_row), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=cat_col, min_row=first_row, max_row=last_row))
        for s_, c in zip(ch.series, cols or colors):
            s_.graphicalProperties.solidFill = c; s_.graphicalProperties.line.solidFill = c
        ch.x_axis.delete = False; ch.y_axis.delete = False
        if ytitle: ch.y_axis.title = ytitle
        tidy(ch)
        ws.add_chart(ch, anchor)

    stacked(ws, yh, yh + 1, yh + 12, 3, 2 + n, 2, f"B{yh + 15}", "Uploads by month (selected year)")
    row = yh + 15 + 20

    # ---- Per year
    section(ws, row, "Uploads per year", "B", "O")
    ph = row + 1
    head(ws, ph, 2, ["Year"] + [None] * n + ["Total", "Time uploaded", "Watched (dated)", "Time watched (dated)"], height=46)
    for i in range(n):
        f(ws, f"{CL(3 + i)}{ph}", f"=Channels!$C${CH_HDR + 1 + i}", font=fnt(9, True, "FFFFFF"), fl=colors[i], al=CENTER, border=BOX)
    py0 = ph + 1
    for yi, y in enumerate(years):
        r = py0 + yi
        put(ws, f"B{r}", y, font=fnt(10, True), al=CENTER, border=BOX)
        lo = f'">="&DATE($B{r},1,1)'; hi = f'"<"&DATE($B{r}+1,1,1)'
        for i in range(n):
            f(ws, f"{CL(3 + i)}{r}", f'=COUNTIFS({V("Day")},{lo},{V("Day")},{hi},{V("Channel ID")},Channels!$E${CH_HDR + 1 + i})',
              fmt='#,##0;-#,##0;"·"', al=CENTER, border=BOX)
        tc = CL(3 + n)
        f(ws, f"{tc}{r}", f"=SUM(C{r}:{CL(2 + n)}{r})", fmt="#,##0", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
        f(ws, f"{CL(4 + n)}{r}", f'=SUMIFS({V("Length")},{V("Day")},{lo},{V("Day")},{hi})', fmt=HM_Z, al=CENTER, border=BOX)
        f(ws, f"{CL(5 + n)}{r}", f'=COUNTIFS({V("Watched On")},{lo},{V("Watched On")},{hi},{V("Status")},"Watched")', fmt='#,##0;-#,##0;"·"', al=CENTER, border=BOX)
        f(ws, f"{CL(6 + n)}{r}", f'=SUMIFS({V("Length")},{V("Watched On")},{lo},{V("Watched On")},{hi},{V("Status")},"Watched")', fmt=HM_Z, al=CENTER, border=BOX)
    py1 = py0 + len(years) - 1
    H_ROWS.update(py0=py0, py1=py1, total_col=CL(3 + n), hours_col=CL(4 + n))
    stacked(ws, ph, py0, py1, 3, 2 + n, 2, f"B{py1 + 2}", "Videos uploaded per year")
    row = py1 + 2 + 20

    # ---- Monthly timeline
    section(ws, row, "Month-by-month timeline", "B", "O")
    months = []
    if videos:
        y, m = videos[0]["local"].year, videos[0]["local"].month
        while (y, m) <= (latest.year, latest.month):
            months.append(date(y, m, 1)); m += 1
            if m == 13: y, m = y + 1, 1
    mh = row + 1 + 42       # leave space for charts above the table
    stacked_row_anchor = row + 1
    head(ws, mh, 2, ["Month"] + [None] * n + ["Total", "Time uploaded", "Watched (dated)", "Time watched (dated)", "Activity"], height=46)
    for i in range(n):
        f(ws, f"{CL(3 + i)}{mh}", f"=Channels!$C${CH_HDR + 1 + i}", font=fnt(9, True, "FFFFFF"), fl=colors[i], al=CENTER, border=BOX)
    m0 = mh + 1
    tc = CL(3 + n)
    for k, mo in enumerate(months):
        r = m0 + k
        put(ws, f"B{r}", mo, fmt="mmm yyyy", font=fnt(10, True), al=CENTER, border=BOX)
        lo = f'">="&$B{r}'; hi = f'"<"&EDATE($B{r},1)'
        for i in range(n):
            f(ws, f"{CL(3 + i)}{r}", f'=COUNTIFS({V("Day")},{lo},{V("Day")},{hi},{V("Channel ID")},Channels!$E${CH_HDR + 1 + i})',
              fmt='#,##0;-#,##0;"·"', al=CENTER, border=BOX)
        f(ws, f"{tc}{r}", f"=SUM(C{r}:{CL(2 + n)}{r})", fmt='#,##0;-#,##0;"·"', font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
        f(ws, f"{CL(4 + n)}{r}", f'=SUMIFS({V("Length")},{V("Day")},{lo},{V("Day")},{hi})', fmt=HM_Z, al=CENTER, border=BOX)
        f(ws, f"{CL(5 + n)}{r}", f'=COUNTIFS({V("Watched On")},{lo},{V("Watched On")},{hi},{V("Status")},"Watched")', fmt='#,##0;-#,##0;"·"', al=CENTER, border=BOX)
        f(ws, f"{CL(6 + n)}{r}", f'=SUMIFS({V("Length")},{V("Watched On")},{lo},{V("Watched On")},{hi},{V("Status")},"Watched")', fmt=HM_Z, al=CENTER, border=BOX)
        f(ws, f"{CL(7 + n)}{r}", f'=IF({tc}{r}=0,"Gap",IF({tc}{r}>AVERAGE(${tc}${m0}:${tc}${m0 + len(months) - 1})+2*STDEV(${tc}${m0}:${tc}${m0 + len(months) - 1}),"High",""))',
          al=CENTER, border=BOX, font=fnt(9, True, "6B7385"))
    m1 = m0 + max(len(months), 1) - 1
    H_ROWS.update(m0=m0, m1=m1, mtotal=tc)
    ws.conditional_formatting.add(f"{CL(7 + n)}{m0}:{CL(7 + n)}{m1}", CellIsRule(operator="equal", formula=['"High"'], fill=fill("FFD9C2")))
    ws.conditional_formatting.add(f"{CL(7 + n)}{m0}:{CL(7 + n)}{m1}", CellIsRule(operator="equal", formula=['"Gap"'], fill=fill("E6E8EC")))
    # month charts (above table)
    chm = BarChart(); chm.type = "col"; chm.grouping = "stacked"; chm.overlap = 100; chm.gapWidth = 10
    chm.title = "Videos uploaded per month, by channel"; chm.width, chm.height = 33, 9.5
    chm.add_data(Reference(ws, min_col=3, max_col=2 + n, min_row=mh, max_row=m1), titles_from_data=True)
    chm.set_categories(Reference(ws, min_col=2, min_row=m0, max_row=m1))
    for s_, c in zip(chm.series, colors):
        s_.graphicalProperties.solidFill = c; s_.graphicalProperties.line.solidFill = c
    chm.x_axis.delete = False; chm.y_axis.delete = False; chm.x_axis.number_format = "mmm yyyy"
    tidy(chm, left=0.045, width=0.94)
    ws.add_chart(chm, f"B{stacked_row_anchor}")
    chw = BarChart(); chw.type = "col"; chw.gapWidth = 10
    chw.title = "Videos watched per month (only those with a Watched On date)"; chw.width, chw.height = 33, 9
    chw.add_data(Reference(ws, min_col=5 + n, min_row=mh, max_row=m1), titles_from_data=True)
    chw.set_categories(Reference(ws, min_col=2, min_row=m0, max_row=m1))
    chw.series[0].graphicalProperties.solidFill = "2E9E6B"; chw.legend = None
    chw.x_axis.delete = False; chw.y_axis.delete = False; chw.x_axis.number_format = "mmm yyyy"
    tidy(chw, legend=False, left=0.045, width=0.94)
    ws.add_chart(chw, f"B{stacked_row_anchor + 21}")
    ws.freeze_panes = "A5"
    protect(ws)

    # =============================================================== STATS
    ws = W["Stats"]
    title_block(ws, "Stats")
    nav(ws, "B4", "Dashboard", "◄ Dashboard")
    ws.column_dimensions["A"].width = 2; ws.column_dimensions["B"].width = 30
    for ci in range(3, 18):
        ws.column_dimensions[CL(ci)].width = 13
    idr = lambda i: f"Channels!$E${CH_HDR + i}"
    nmr = lambda i: f"Channels!$C${CH_HDR + i}"
    ST = V("Status"); LN = V("Length"); CI = V("Channel ID")

    section(ws, 6, "Library progress", "B", "N")
    head(ws, 7, 2, ["Channel", "Videos", "Watched", "Watch Later", "Unwatched", "Skipped", "% Watched", "Runtime",
                    "Time watched", "Time remaining", "% Runtime Watched", "Average length", "Median length"], height=36)
    A0 = 8
    for i in range(1, n + 1):
        r = A0 + i - 1
        c = idr(i)
        f(ws, f"B{r}", f"={nmr(i)}", font=fnt(10, True), fl=tint(colors[i - 1]), border=BOX)
        f(ws, f"C{r}", f"=COUNTIFS({CI},{c})", fmt="#,##0", al=CENTER, border=BOX)
        for col, st in zip("DEFG", ["Watched", "Watch Later", "Unwatched", "Skipped"]):
            f(ws, f"{col}{r}", f'=COUNTIFS({CI},{c},{ST},"{st}")', fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"H{r}", f"=IFERROR(D{r}/(C{r}-G{r}),0)", fmt="0.0%", al=CENTER, border=BOX)
        f(ws, f"I{r}", f"=SUMIFS({LN},{CI},{c})", fmt=HM_FMT, al=CENTER, border=BOX)
        f(ws, f"J{r}", f'=SUMIFS({LN},{CI},{c},{ST},"Watched")', fmt=HM_FMT, al=CENTER, border=BOX)
        f(ws, f"K{r}", f'=(SUMIFS({LN},{CI},{c},{ST},"Unwatched")+SUMIFS({LN},{CI},{c},{ST},"Watch Later"))', fmt=HM_FMT, al=CENTER, border=BOX)
        f(ws, f"L{r}", f'=IFERROR(J{r}/(I{r}-SUMIFS({LN},{CI},{c},{ST},"Skipped")),0)', fmt="0.0%", al=CENTER, border=BOX)
        f(ws, f"M{r}", f"=IFERROR(AVERAGEIFS({LN},{CI},{c}),0)", fmt="[h]:mm:ss", al=CENTER, border=BOX)
        ws[f"N{r}"] = ArrayFormula(f"N{r}", f"=IFERROR(MEDIAN(IF({CI}={c},{LN})),0)")
        ws[f"N{r}"].number_format = "[h]:mm:ss"; ws[f"N{r}"].alignment = CENTER; ws[f"N{r}"].border = BOX; ws[f"N{r}"].font = fnt(10)
    AT = A0 + n
    put(ws, f"B{AT}", "All channels", font=fnt(10, True), fl="EEF0F5", border=BOX)
    for col in "CDEFGIJK":
        f(ws, f"{col}{AT}", f"=SUM({col}{A0}:{col}{AT - 1})", fmt=HM_FMT if col in "IJK" else "#,##0", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"H{AT}", f"=IFERROR(D{AT}/(C{AT}-G{AT}),0)", fmt="0.0%", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"L{AT}", f'=IFERROR(J{AT}/(I{AT}-SUMIFS({LN},{ST},"Skipped")),0)', fmt="0.0%", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"M{AT}", f"=IFERROR(AVERAGE({LN}),0)", fmt="[h]:mm:ss", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"N{AT}", f"=IFERROR(MEDIAN({LN}),0)", fmt="[h]:mm:ss", font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    ws.conditional_formatting.add(f"H{A0}:H{AT}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="2E9E6B"))
    put(ws, f"B{AT + 1}", "% Watched = Watched / (Videos - Skipped).  Remaining = Unwatched + Watch Later.", font=fnt(8, color="6B7385", italic=True))

    r = AT + 4
    section(ws, r, "Upload profile", "B", "Q")
    head(ws, r + 1, 2, ["Channel", "First upload", "Latest upload", "Active span (yrs)", "Videos / month", "Avg days between",
                        "Busiest year", "Videos that year", "Longest", "Shortest (>0)", "Total views", "Avg views / video", "Most viewed video"],
         height=36)
    ws.merge_cells(start_row=r + 1, start_column=14, end_row=r + 1, end_column=17)
    B0 = r + 2
    yr_rng = f"History!$B${H_ROWS['py0']}:$B${H_ROWS['py1']}"
    DY = V("Day")
    for i in range(1, n + 1):
        rr = B0 + i - 1; c = idr(i); ar = A0 + i - 1
        hc = CL(2 + i)
        hist = f"History!${hc}${H_ROWS['py0']}:${hc}${H_ROWS['py1']}"
        f(ws, f"B{rr}", f"={nmr(i)}", font=fnt(10, True), fl=tint(colors[i - 1]), border=BOX)
        f(ws, f"C{rr}", f'=IF(C{ar}=0,"",_xlfn.MINIFS({DY},{CI},{c}))', fmt="yyyy-mm-dd", al=CENTER, border=BOX)
        f(ws, f"D{rr}", f'=IF(C{ar}=0,"",_xlfn.MAXIFS({DY},{CI},{c}))', fmt="yyyy-mm-dd", al=CENTER, border=BOX)
        f(ws, f"E{rr}", f'=IF(C{ar}=0,"",(D{rr}-C{rr})/365.25)', fmt="0.0", al=CENTER, border=BOX)
        f(ws, f"F{rr}", f'=IF(C{ar}=0,"",C{ar}/MAX(1,(D{rr}-C{rr})/30.4375))', fmt="0.0", al=CENTER, border=BOX)
        f(ws, f"G{rr}", f'=IF(C{ar}<2,"",(D{rr}-C{rr})/(C{ar}-1))', fmt="0.0", al=CENTER, border=BOX)
        f(ws, f"H{rr}", f'=IF(C{ar}=0,"",INDEX({yr_rng},MATCH(MAX({hist}),{hist},0)))', fmt="0", al=CENTER, border=BOX)
        f(ws, f"I{rr}", f"=MAX({hist})", fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"J{rr}", f'=IF(C{ar}=0,"",_xlfn.MAXIFS({LN},{CI},{c}))', fmt="[h]:mm:ss", al=CENTER, border=BOX)
        f(ws, f"K{rr}", f'=IF(C{ar}=0,"",_xlfn.MINIFS({LN},{CI},{c},{LN},">0"))', fmt="[h]:mm:ss", al=CENTER, border=BOX)
        f(ws, f"L{rr}", f"=SUMIFS({V('Views')},{CI},{c})", fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"M{rr}", f'=IFERROR(L{rr}/C{ar},0)', fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"N{rr}", f'=IF(C{ar}=0,"",IFERROR(INDEX({V("Title")},MATCH(_xlfn.MAXIFS({V("Seq")},{CI},{c},{V("Views")},_xlfn.MAXIFS({V("Views")},{CI},{c})),{V("Seq")},0)),""))',
          border=BOX, font=fnt(9))
        ws.merge_cells(start_row=rr, start_column=14, end_row=rr, end_column=17)
    BT = B0 + n
    put(ws, f"B{BT}", "All channels", font=fnt(10, True), fl="EEF0F5", border=BOX)
    f(ws, f"C{BT}", f"=MIN({DY})", fmt="yyyy-mm-dd", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"D{BT}", f"=MAX({DY})", fmt="yyyy-mm-dd", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"E{BT}", f"=(D{BT}-C{BT})/365.25", fmt="0.0", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"F{BT}", f"=C{AT}/MAX(1,(D{BT}-C{BT})/30.4375)", fmt="0.0", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"G{BT}", f"=(D{BT}-C{BT})/MAX(1,C{AT}-1)", fmt="0.0", al=CENTER, border=BOX, fl="EEF0F5")
    tcol = H_ROWS["total_col"]; thist = f"History!${tcol}${H_ROWS['py0']}:${tcol}${H_ROWS['py1']}"
    f(ws, f"H{BT}", f"=INDEX({yr_rng},MATCH(MAX({thist}),{thist},0))", fmt="0", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"I{BT}", f"=MAX({thist})", fmt="#,##0", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"J{BT}", f"=MAX({LN})", fmt="[h]:mm:ss", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"K{BT}", f'=IFERROR(_xlfn.MINIFS({LN},{LN},">0"),"")', fmt="[h]:mm:ss", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"L{BT}", f"=SUM({V('Views')})", fmt="#,##0", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"M{BT}", f"=IFERROR(L{BT}/C{AT},0)", fmt="#,##0", al=CENTER, border=BOX, fl="EEF0F5")
    f(ws, f"N{BT}", f'=IFERROR(INDEX({V("Title")},MATCH(MAX({V("Views")}),{V("Views")},0)),"")', border=BOX, fl="EEF0F5", font=fnt(9))
    ws.merge_cells(start_row=BT, start_column=14, end_row=BT, end_column=17)
    put(ws, f"B{BT + 1}", "Views accumulate with age, so older videos are not directly comparable to recent ones.", font=fnt(8, color="6B7385", italic=True))

    r = BT + 4
    section(ws, r, "Pace: uploading vs. watching", "B", "N")
    head(ws, r + 1, 2, ["", "Last 30 days", "Last 90 days", "Last 365 days"], height=22)
    P0 = r + 2
    wins = [("C", 30), ("D", 90), ("E", 365)]
    labels = ["Videos uploaded", "Time uploaded", "Videos watched *", "Time watched *", "Backlog change (videos)", "Backlog change (time)"]
    for k, lab in enumerate(labels):
        put(ws, f"B{P0 + k}", lab, font=fnt(10, True), border=BOX)
    for col, w in wins:
        up = f'{V("Day")},">"&TODAY()-{w}'
        wt = f'{V("Watched On")},">"&TODAY()-{w},{ST},"Watched"'
        f(ws, f"{col}{P0}", f"=COUNTIFS({up})", fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"{col}{P0 + 1}", f"=SUMIFS({LN},{up})", fmt=HM_FMT, al=CENTER, border=BOX)
        f(ws, f"{col}{P0 + 2}", f"=COUNTIFS({wt})", fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"{col}{P0 + 3}", f"=SUMIFS({LN},{wt})", fmt=HM_FMT, al=CENTER, border=BOX)
        f(ws, f"{col}{P0 + 4}", f"={col}{P0}-{col}{P0 + 2}", fmt="+#,##0;-#,##0;0", al=CENTER, border=BOX)
        f(ws, f"{col}{P0 + 5}", f"=ROUND(({col}{P0 + 1}-{col}{P0 + 3})*24,0)", fmt='+#,##0" h";-#,##0" h";0" h"', al=CENTER, border=BOX)
    ws.conditional_formatting.add(f"C{P0 + 4}:E{P0 + 5}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C0392B", bold=True)))
    ws.conditional_formatting.add(f"C{P0 + 4}:E{P0 + 5}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="1E6B34", bold=True)))
    e = P0 + 6
    put(ws, f"B{e}", "Estimated days to clear backlog", font=fnt(10, True), border=BOX)
    f(ws, f"C{e}", f'=IF(D{P0 + 3}/90-D{P0 + 1}/90<=0,"Backlog is growing",K{AT}/(D{P0 + 3}/90-D{P0 + 1}/90))', fmt="#,##0", al=CENTER, border=BOX)
    ws.merge_cells(f"C{e}:D{e}")
    put(ws, f"F{e}", "Uses the last-90-day watch rate minus the upload rate.", font=fnt(8, color="6B7385", italic=True))
    put(ws, f"B{e + 1}", "* Counts only videos with a 'Watched On' date.", font=fnt(8, color="6B7385", italic=True))

    r = e + 4
    section(ws, r, "Records", "B", "N")
    rec = [
        ("Longest video", f"=MAX({LN})", "[h]:mm:ss", f'=INDEX({V("Title")},MATCH(MAX({LN}),{LN},0))'),
        ("Shortest video (>0)", f'=IFERROR(_xlfn.MINIFS({LN},{LN},">0"),"")', "[h]:mm:ss", f'=IFERROR(INDEX({V("Title")},MATCH(C{r + 2},{LN},0)),"")'),
        ("Most viewed video", f"=MAX({V('Views')})", "#,##0", f'=INDEX({V("Title")},MATCH(MAX({V("Views")}),{V("Views")},0))'),
        ("Busiest month (videos)", f"=MAX(History!${H_ROWS['mtotal']}${H_ROWS['m0']}:${H_ROWS['mtotal']}${H_ROWS['m1']})", "#,##0",
         f"=TEXT(INDEX(History!$B${H_ROWS['m0']}:$B${H_ROWS['m1']},MATCH(C{r + 4},History!${H_ROWS['mtotal']}${H_ROWS['m0']}:${H_ROWS['mtotal']}${H_ROWS['m1']},0)),\"mmmm yyyy\")"),
    ]
    for k, (lab, fm, nf, ttl) in enumerate(rec):
        rr = r + 1 + k
        put(ws, f"B{rr}", lab, font=fnt(10, True), border=BOX)
        f(ws, f"C{rr}", fm, fmt=nf, al=CENTER, border=BOX)
        f(ws, f"D{rr}", ttl, font=fnt(9))
    ws.column_dimensions["B"].width = 44
    ws.freeze_panes = "A5"
    protect(ws)
    S_ROWS = dict(A0=A0, AT=AT, B0=B0)

    # =============================================================== DASHBOARD
    ws = W["Dashboard"]
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    for ci in range(2, 14):
        ws.column_dimensions[CL(ci)].width = 12.5
    put(ws, "B2", "SCB Video Library", font=fnt(22, True, NAVY)); ws.row_dimensions[2].height = 32
    f(ws, "B3", '=IF(Settings!$C$5="","","Last updated "&TEXT(Settings!$C$5,"d mmm yyyy h:mm AM/PM")&"   ·   "&TODAY()-INT(Settings!$C$5)&" day(s) ago")',
      font=fnt(9, False, "6B7385", italic=True))
    ws.conditional_formatting.add("B3", FormulaRule(formula=["TODAY()-INT(Settings!$C$5)>7"], font=Font(color="C0392B", bold=True, italic=True)))
    for k, (sh, lab) in enumerate([("Library", "Library"), ("Queue", "Queue"), ("Calendar", "Calendar"), ("History", "History"),
                                    ("Stats", "Stats"), ("Channels", "Channels"), ("Settings", "Settings")]):
        c = f"{CL(2 + k)}4"
        f(ws, c, f'=HYPERLINK("#{sh}!A1","{lab}")', font=fnt(10, True, "FFFFFF", underline=None), fl=TEAL, al=CENTER)
    ws.row_dimensions[4].height = 22
    tiles = [
        ("TOTAL VIDEOS", f"=Stats!C{AT}", "#,##0", f'="across "&{n}&" channels"'),
        ("WATCHED", f"=Stats!D{AT}", "#,##0", f'="of "&TEXT(Stats!C{AT}-Stats!G{AT},"#,##0")&" tracked"'),
        ("BACKLOG", f"=Stats!E{AT}+Stats!F{AT}", "#,##0", f'=TEXT(Stats!E{AT},"#,##0")&" watch-later"'),
        ("% WATCHED", f"=Stats!H{AT}", "0.0%", f'=TEXT(Stats!L{AT},"0.0%")&" of runtime"'),
        ("HOURS WATCHED", f"=Stats!J{AT}*24", '#,##0" h"', f'="of "&TEXT(Stats!I{AT}*24,"#,##0")&" h total"'),
        ("HOURS REMAINING", f"=Stats!K{AT}*24", '#,##0" h"', f'="≈ "&TEXT(Stats!K{AT},"#,##0")&" days of video"'),
    ]
    for k, (lab, fm, nf, sub) in enumerate(tiles):
        c1, c2 = CL(2 + 2 * k), CL(3 + 2 * k)
        for rr in (6, 7, 8):
            ws.merge_cells(f"{c1}{rr}:{c2}{rr}")
            for cc in (c1, c2):
                ws[f"{cc}{rr}"].fill = fill("F3F5F9")
        put(ws, f"{c1}6", lab, font=fnt(8, True, "6B7385"), al=CENTER, fl="F3F5F9")
        f(ws, f"{c1}7", fm, font=fnt(22, True, TEAL if k != 3 else "2E9E6B"), al=CENTER, fmt=nf, fl="F3F5F9")
        f(ws, f"{c1}8", sub, font=fnt(8, color="6B7385"), al=CENTER, fl="F3F5F9")
    ws.row_dimensions[7].height = 34

    section(ws, 10, "By channel")
    head(ws, 11, 2, ["Channel", "", "", "Videos", "Watched", "Backlog", "% Watched", "Time left", "Latest upload", "", "Days since", ""], height=22)
    ws.merge_cells("B11:D11"); ws.merge_cells("J11:K11"); ws.merge_cells("L11:M11")
    for i in range(1, n + 1):
        r = 11 + i; ar = A0 + i - 1; br = B0 + i - 1
        ws.merge_cells(f"B{r}:D{r}"); ws.merge_cells(f"J{r}:K{r}"); ws.merge_cells(f"L{r}:M{r}")
        f(ws, f"B{r}", f"=Stats!B{ar}", font=fnt(10, True), fl=tint(colors[i - 1]), border=BOX)
        for col, src, nf in [("E", "C", "#,##0"), ("F", "D", "#,##0"), ("H", "H", "0.0%"), ("I", "K", HM_FMT)]:
            f(ws, f"{col}{r}", f"=Stats!{src}{ar}", fmt=nf, al=CENTER, border=BOX)
        f(ws, f"G{r}", f"=Stats!E{ar}+Stats!F{ar}", fmt="#,##0", al=CENTER, border=BOX)
        f(ws, f"J{r}", f"=Stats!D{br}", fmt="d mmm yyyy", al=CENTER, border=BOX)
        f(ws, f"L{r}", f'=IF(J{r}="","",TODAY()-J{r})', fmt="#,##0", al=CENTER, border=BOX)
    r = 12 + n
    ws.merge_cells(f"B{r}:D{r}")
    put(ws, f"B{r}", "All channels", font=fnt(10, True), fl="EEF0F5", border=BOX)
    for col, fm, nf in [("E", f"=Stats!C{AT}", "#,##0"), ("F", f"=Stats!D{AT}", "#,##0"), ("G", f"=Stats!E{AT}+Stats!F{AT}", "#,##0"),
                        ("H", f"=Stats!H{AT}", "0.0%"), ("I", f"=Stats!K{AT}", HM_FMT)]:
        f(ws, f"{col}{r}", fm, fmt=nf, font=fnt(10, True), al=CENTER, border=BOX, fl="EEF0F5")
    ws.conditional_formatting.add(f"H12:H{11 + n}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="2E9E6B"))

    cr = r + 2
    bc = BarChart(); bc.type = "bar"; bc.grouping = "stacked"; bc.overlap = 100; bc.gapWidth = 50
    bc.title = "Watched vs. backlog (videos)"; bc.width, bc.height = 13.6, 8.5
    bs = W["Stats"]
    for col, color in [(4, "2E9E6B"), (5, "E0B33A"), (6, "C5CAD6")]:
        bc.add_data(Reference(bs, min_col=col, min_row=7, max_row=A0 + n - 1), titles_from_data=True)
    bc.set_categories(Reference(bs, min_col=2, min_row=A0, max_row=A0 + n - 1))
    for s_, c in zip(bc.series, ["2E9E6B", "E0B33A", "C5CAD6"]):
        s_.graphicalProperties.solidFill = c; s_.graphicalProperties.line.solidFill = c
    bc.x_axis.delete = False; bc.y_axis.delete = False; bc.x_axis.scaling.orientation = "maxMin"
    bc.y_axis.crosses = "max"      # keep the value-axis numbers at the bottom
    tidy(bc, left=0.30, width=0.64)
    ws.add_chart(bc, f"B{cr}")
    yc = BarChart(); yc.type = "col"; yc.gapWidth = 40; yc.title = "Videos uploaded per year"; yc.width, yc.height = 13.6, 8.5
    yc.add_data(Reference(H, min_col=3 + n, min_row=ph, max_row=py1), titles_from_data=True)
    yc.set_categories(Reference(H, min_col=2, min_row=py0, max_row=py1))
    yc.series[0].graphicalProperties.solidFill = TEAL; yc.legend = None
    yc.x_axis.delete = False; yc.y_axis.delete = False
    tidy(yc, legend=False)
    ws.add_chart(yc, f"H{cr}")

    ur = cr + 19
    section(ws, ur, "Up next  -  oldest unwatched video in each channel")
    head(ws, ur + 1, 2, ["Channel", "", "", "Video", "", "", "", "", "", "", "Length", "Published"], height=20)
    ws.merge_cells(f"B{ur + 1}:D{ur + 1}"); ws.merge_cells(f"E{ur + 1}:K{ur + 1}")
    for i in range(1, n + 1):
        r = ur + 1 + i
        ws.merge_cells(f"B{r}:D{r}"); ws.merge_cells(f"E{r}:K{r}")
        f(ws, f"O{r}", f'=IF(_xlfn.MINIFS({V("Seq")},{CI},{idr(i)},{ST},"Unwatched")=0,"",MATCH(_xlfn.MINIFS({V("Seq")},{CI},{idr(i)},{ST},"Unwatched"),{V("Seq")},0))')
        f(ws, f"B{r}", f"={nmr(i)}", font=fnt(10, True), fl=tint(colors[i - 1]), border=BOX)
        f(ws, f"E{r}", f'=IF($O{r}="","All caught up ✓",IF(INDEX({V("Video ID")},$O{r})="",INDEX({V("Title")},$O{r}),HYPERLINK("https://www.youtube.com/watch?v="&INDEX({V("Video ID")},$O{r}),INDEX({V("Title")},$O{r}))))',
          font=fnt(10, False, "1155CC"), border=BOX)
        f(ws, f"L{r}", f'=IF($O{r}="","",INDEX({V("Length")},$O{r}))', fmt="[h]:mm:ss", al=CENTER, border=BOX)
        f(ws, f"M{r}", f'=IF($O{r}="","",INDEX({V("Day")},$O{r}))', fmt="d mmm yyyy", al=CENTER, border=BOX)
    rr0 = ur + n + 4
    section(ws, rr0, "Newest uploads")
    head(ws, rr0 + 1, 2, ["Published", "", "Channel", "", "Title", "", "", "", "", "", "Length", "Status"], height=20)
    for a, b in [("B", "C"), ("D", "E"), ("F", "K")]:
        ws.merge_cells(f"{a}{rr0 + 1}:{b}{rr0 + 1}")
    for k in range(1, 11):
        r = rr0 + 1 + k
        for a, b in [("B", "C"), ("D", "E"), ("F", "K")]:
            ws.merge_cells(f"{a}{r}:{b}{r}")
        f(ws, f"O{r}", f'=IFERROR(MATCH(MAX({V("Seq")})-{k - 1},{V("Seq")},0),"")')
        f(ws, f"B{r}", f'=IF($O{r}="","",INDEX({V("Day")},$O{r}))', fmt="ddd d mmm yyyy", al=CENTER, border=BOX)
        f(ws, f"D{r}", f'=IF($O{r}="","",INDEX({V("Channel")},$O{r}))', border=BOX)
        f(ws, f"F{r}", f'=IF($O{r}="","",IF(INDEX({V("Video ID")},$O{r})="",INDEX({V("Title")},$O{r}),HYPERLINK("https://www.youtube.com/watch?v="&INDEX({V("Video ID")},$O{r}),INDEX({V("Title")},$O{r}))))',
          font=fnt(10, False, "1155CC"), border=BOX)
        f(ws, f"L{r}", f'=IF($O{r}="","",INDEX({V("Length")},$O{r}))', fmt="[h]:mm:ss", al=CENTER, border=BOX)
        f(ws, f"M{r}", f'=IF($O{r}="","",INDEX({V("Status")},$O{r}))', al=CENTER, border=BOX)
    ws.conditional_formatting.add(f"M{rr0 + 2}:M{rr0 + 11}", CellIsRule(operator="equal", formula=['"Watched"'], fill=fill("DDF2E2")))
    ws.column_dimensions.group("O", "O", hidden=True)
    protect(ws)

    # =============================================================== QUEUE
    ws = W["Queue"]
    title_block(ws, "Watch Queue")
    put(ws, "B6", "Your next videos. Click a title to open it on YouTube; mark it watched in the Library.", font=fnt(9, False, "6B7385", italic=True))
    nav(ws, "B4", "Dashboard", "◄ Dashboard")
    for col, w in zip("ABCDEFGHI", [2, 5, 20, 26, 80, 11, 13, 3, 3]):
        ws.column_dimensions[col].width = w
    selector(ws, "C4", sel("Queue", "C4", "Unwatched"), "C3", "Show")
    selector(ws, "D4", sel("Queue", "D4", "All channels"), "D3", "Channel")
    selector(ws, "E4", sel("Queue", "E4", "Oldest first"), "E3", "Order")
    dv_list(ws, "C4", '"Unwatched,Watch Later"'); dv_list(ws, "D4", "=ChannelList"); dv_list(ws, "E4", '"Oldest first,Newest first"')
    f(ws, "K4", crit_formula("$D$4"))
    ws.column_dimensions.group("I", "K", hidden=True)
    qk = V("Queue Key")
    q_rt = hm("SUMIFS(" + LN + "," + qk + ',">0")')
    f(ws, "B5", '=COUNT(' + qk + ')&" videos match   ·   "&' + q_rt + '&" of runtime"',
      font=fnt(9, True, "6B7385"))
    head(ws, 7, 2, ["#", "Published", "Channel", "Title", "Length", "Status"], height=22)
    QN = 30
    for k in range(1, QN + 1):
        r = 7 + k
        f(ws, f"I{r}", f'=IFERROR(MATCH(IF($E$4="Newest first",LARGE({V("Queue Key")},{k}),SMALL({V("Queue Key")},{k})),{V("Seq")},0),"")')
        f(ws, f"B{r}", f'=IF($I{r}="","",{k})', al=CENTER, border=BOX, font=fnt(9, color="6B7385"))
        f(ws, f"C{r}", f'=IF($I{r}="","",INDEX({V("Day")},$I{r}))', fmt="d mmm yyyy", al=CENTER, border=BOX)
        f(ws, f"D{r}", f'=IF($I{r}="","",INDEX({V("Channel")},$I{r}))', border=BOX)
        f(ws, f"E{r}", f'=IF($I{r}="","",IF(INDEX({V("Video ID")},$I{r})="",INDEX({V("Title")},$I{r}),HYPERLINK("https://www.youtube.com/watch?v="&INDEX({V("Video ID")},$I{r}),INDEX({V("Title")},$I{r}))))',
          font=fnt(10, False, "1155CC"), border=BOX)
        f(ws, f"F{r}", f'=IF($I{r}="","",INDEX({V("Length")},$I{r}))', fmt="[h]:mm:ss", al=CENTER, border=BOX)
        f(ws, f"G{r}", f'=IF($I{r}="","",INDEX({ST},$I{r}))', al=CENTER, border=BOX)
    for i in range(n):
        ws.conditional_formatting.add(f"D8:D{7 + QN}", FormulaRule(formula=[f'$D8=Channels!$C${CH_HDR + 1 + i}'], fill=fill(tint(colors[i]))))
    ws.freeze_panes = "A8"
    protect(ws)

    # =============================================================== CALENDAR
    ws = W["Calendar"]
    title_block(ws, "Upload Calendar")
    put(ws, "B5", "Pick a year, month and (optionally) a channel with the dropdowns. Click a title to open it on YouTube.  ✓ = watched, dimmed.", font=fnt(9, False, "6B7385", italic=True))
    nav(ws, "B4", "Dashboard", "◄ Dashboard")
    ws["B4"].alignment = LEFT
    ws.column_dimensions["A"].width = 2
    for ci in range(2, 9):
        ws.column_dimensions[CL(ci)].width = 27
    ws.column_dimensions["I"].width = 3
    lm = latest.month
    selector(ws, "C4", sel("Calendar", "C4", latest.year), "C3", "Year")
    selector(ws, "E4", sel("Calendar", "E4", MONTHS[lm - 1]), "E3", "Month")
    selector(ws, "G4", sel("Calendar", "G4", "All channels"), "G3", "Channel")
    dv_list(ws, "C4", "=YearList"); dv_list(ws, "E4", "=MonthList"); dv_list(ws, "G4", "=ChannelList")
    f(ws, "K2", "=DATE($C$4,MATCH($E$4,MonthList,0),1)", fmt="yyyy-mm-dd")
    f(ws, "K3", "=K2-(WEEKDAY(K2,1)-1)", fmt="yyyy-mm-dd")
    f(ws, "K4", crit_formula("$G$4"))
    put(ws, "J2", "month start"); put(ws, "J3", "grid start"); put(ws, "J4", "channel crit")
    sf = f'{V("Day")},">="&$K$2,{V("Day")},"<"&EDATE($K$2,1),{V("Channel ID")},$K$4'
    cal_rt = hm("SUMIFS(" + LN + "," + sf + ")")
    f(ws, "B6", '="Videos: "&COUNTIFS(' + sf + ')&"      Runtime: "&' + cal_rt + '&"      Watched: "&COUNTIFS(' + sf + ',' + ST + ',"Watched")&" of "&COUNTIFS(' + sf + ')',
      font=fnt(11, True, NAVY))
    legend = []
    for i in range(n):
        rr = 7 if i < 6 else 8
        col = CL(2 + (i % 6))
        cell = f"{col}{rr}"
        f(ws, cell, f'=Channels!$C${CH_HDR + 1 + i}&"  ("&COUNTIFS({V("Day")},">="&$K$2,{V("Day")},"<"&EDATE($K$2,1),{CI},Channels!$E${CH_HDR + 1 + i})&")"',
          font=fnt(9, True, "333333"), fl=tint(colors[i]), al=CENTER, border=BOX)
    for ci, lab in enumerate(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]):
        put(ws, f"{CL(2 + ci)}9", lab, font=fnt(10, True, "FFFFFF"), fl=NAVY, al=CENTER)
    R0 = 10
    slot_ranges, block_cf = [], []
    for w in range(6):
        h = R0 + 6 * w
        for j in range(7):
            cal, g1, g2, g3 = CL(2 + j), CL(11 + j), CL(19 + j), CL(27 + j)
            f(ws, f"{g1}{h}", f"=$K$3+{7 * w + j}", fmt="yyyy-mm-dd")
            f(ws, f"{g2}{h}", f'=IF(MONTH({g1}{h})=MONTH($K$2),COUNTIFS({V("Day")},{g1}{h},{CI},$K$4),0)')
            f(ws, f"{g3}{h}", f'=IF({g2}{h}>0,SUMIFS({LN},{V("Day")},{g1}{h},{CI},$K$4),0)')
            f(ws, f"{cal}{h}",
              f'=IF(MONTH({g1}{h})=MONTH($K$2),DAY({g1}{h})&IF({g2}{h}>0,"   ·   "&{g2}{h}&IF({g2}{h}=1," video"," videos")&"   ·   "&{hm(g3 + str(h))}&IF({g2}{h}>{MAX_SLOTS},"   (+"&({g2}{h}-{MAX_SLOTS})&" more)",""),""),"")',
              font=fnt(10, True, NAVY), fl="E9ECF3", al=LEFT, border=BOX)
            for s in range(1, MAX_SLOTS + 1):
                rs = h + s
                f(ws, f"{g1}{rs}",
                  f'=IF({g2}{h}>={s},MATCH({g1}{h}*100+{s},{V("Cal Key")},0),"")')
                f(ws, f"{g2}{rs}", f'=IF({g1}{rs}="","",IFERROR(MATCH(INDEX({CI},{g1}{rs}),tblChannels[Channel ID],0),""))')
                f(ws, f"{cal}{rs}",
                  f'=IF({g1}{rs}="","",IF(INDEX({V("Video ID")},{g1}{rs})="",IF(INDEX({ST},{g1}{rs})="Watched","✓ ","▸ ")&INDEX({V("Title")},{g1}{rs}),'
                  f'HYPERLINK("https://www.youtube.com/watch?v="&INDEX({V("Video ID")},{g1}{rs}),IF(INDEX({ST},{g1}{rs})="Watched","✓ ","▸ ")&INDEX({V("Title")},{g1}{rs}))))',
                  font=fnt(9), al=Alignment(horizontal="left", vertical="center", wrap_text=False), border=BOX)
            ws.row_dimensions[h].height = 18
        slot_ranges.append(f"B{h + 1}:H{h + MAX_SLOTS}")
        ws.conditional_formatting.add(f"B{h}:H{h + MAX_SLOTS}",
                                      FormulaRule(formula=[f"MONTH(K${h})<>MONTH($K$2)"], fill=fill("D5D9E2"), stopIfTrue=True))
        ws.conditional_formatting.add(f"B{h}:H{h}", FormulaRule(formula=[f"K{h}=TODAY()"], font=Font(bold=True, color="C0392B")))
    sq = " ".join(slot_ranges)
    first_slot = f"B{R0 + 1}"
    for i in range(n):
        ws.conditional_formatting.add(sq, FormulaRule(formula=[f"S{R0 + 1}={i + 1}"], fill=fill(tint(colors[i], 0.72))))
    ws.conditional_formatting.add(sq, FormulaRule(formula=[f'LEFT({first_slot},1)="✓"'], font=Font(color="8A93A6")))
    ws.column_dimensions.group("J", "AG", hidden=True)
    ws.sheet_view.zoomScale = 90
    protect(ws)

    # ---- finishing
    for nm in ("Library", "Channels", "Settings"):
        pass
    if fan:
        ws = W["Carry Over"]
        ws.sheet_properties.tabColor = GOLD
        put(ws, "A1", "STEP 1 - Paste your OLD Library here.  In your old workbook open the Library sheet, click the small square at the very "
                      "top-left (select all), press Ctrl+C. Come back here, click cell A1, then Paste Special > Values.",
            font=fnt(11, True, NAVY))
        put(ws, "A2", "(This note disappears when you paste - that's expected. Both files must come from this site / this tool so the columns line up.)",
            font=fnt(9, False, "6B7385", italic=True))
        last_co = HDR + max(N, 1)
        f(ws, "Y1", f'="CARRY OVER  -  "&COUNT(Y{HDR + 1}:Y{last_co})&" of {N} videos matched with your old file"', font=fnt(12, True, TEAL))
        put(ws, "Y2", "STEP 2 - Copy columns Z:AA (from row 6 down). In the Library sheet click cell C6, then Paste Special > Values.  (Clear any Library filters first.)", font=fnt(10, True, NAVY))
        put(ws, "Y3", "STEP 3 - Copy column AB (Notes). In the Library sheet click cell L6, then Paste Special > Values.", font=fnt(10, True, NAVY))
        put(ws, "Y4", "STEP 4 - On the Channels sheet, re-enter your 'Watched Through' dates. Done - you can delete this sheet.", font=fnt(10, True, NAVY))
        for col, lab in zip(["Y", "Z", "AA", "AB"], ["Matched row in old file", "Mark  (-> Library C)", "Watched On  (-> Library D)", "Notes  (-> Library L)"]):
            put(ws, f"{col}{HDR}", lab, font=fnt(9, True, "FFFFFF"), fl=GOLD, al=CENTER, border=BOX)
        ws.row_dimensions[HDR].height = 30
        for r in range(HDR + 1, last_co + 1):
            f(ws, f"Y{r}", f'=IFERROR(MATCH(Library!M{r},$M$6:$M$40000,0),"")', font=fnt(9, color="6B7385"))
            f(ws, f"Z{r}", f'=IF($Y{r}="","",INDEX($C$6:$C$40000,$Y{r})&"")')
            f(ws, f"AA{r}", f'=IF($Y{r}="","",IF(INDEX($D$6:$D$40000,$Y{r})="","",INDEX($D$6:$D$40000,$Y{r})))', fmt="yyyy-mm-dd")
            f(ws, f"AB{r}", f'=IF($Y{r}="","",INDEX($L$6:$L$40000,$Y{r})&"")')
        for col, w in zip(["Y", "Z", "AA", "AB"], [16, 16, 16, 40]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = f"A{HDR + 1}"

    from openpyxl.worksheet.properties import PageSetupProperties
    for nm in ("Dashboard", "Queue", "Calendar", "History", "Stats", "Channels", "Settings", "Library"):
        w_ = W[nm]
        w_.page_setup.orientation = "landscape"
        w_.page_setup.fitToWidth = 1
        w_.page_setup.fitToHeight = 0
        w_.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    W["Library"].print_title_rows = f"{HDR}:{HDR}"
    wb.active = 0
    wb.save(path)


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--offline", action="store_true", help="don't call YouTube; just rebuild the workbook")
    ap.add_argument("--import-legacy", metavar="XLSX", help="import the old scb_timeline.xlsx export")
    ap.add_argument("--workbook", default=str(WORKBOOK))
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args()
    path = Path(a.workbook)
    now = datetime.now().replace(microsecond=0)

    existing = dict(settings={}, channels=[], videos=[], selectors={})
    if path.exists():
        try:
            existing = read_workbook(path)
        except PermissionError:
            sys.exit("Can't read the workbook - close it in Excel and try again.")
        if not a.no_backup:
            BACKUPS.mkdir(exist_ok=True)
            shutil.copy2(path, BACKUPS / f"SCB_Library_{now:%Y%m%d_%H%M%S}.xlsx")
            for old in sorted(BACKUPS.glob("SCB_Library_*.xlsx"))[:-30]:
                old.unlink()
    channels = existing["channels"] or [dict(name=c["name"], handle=c["handle"], id=c["id"], enabled=True,
                                              through=None, yt=c["yt"], refreshed=None) for c in DEFAULT_CHANNELS]
    videos = existing["videos"]
    if a.import_legacy:
        have = {(v["channel_title"], v["published_utc"], v["title"]) for v in videos}
        legacy = read_legacy(a.import_legacy)
        yt_to_id = {c.get("yt"): c["id"] for c in channels}
        for v in legacy:
            v["channel_id"] = yt_to_id.get(v["channel_title"])
        videos += legacy
        print(f"Imported {len(legacy)} videos from {a.import_legacy}")
    tzname = existing["settings"].get("tz") or DEFAULT_TZ
    tz = get_tz(tzname)

    if not a.offline:
        try:
            yt = YT(load_api_key())
            for c in channels:
                if not c.get("id"):
                    c["id"] = resolve_channel(yt, c.get("handle") or c.get("name"))
                    if not c["id"]:
                        sys.exit(f"Couldn't find a channel for '{c.get('handle') or c.get('name')}'.")
                    print(f"Resolved {c.get('handle')} -> {c['id']}")
            active = [c for c in channels if c["enabled"]]
            info = fetch_channel_info(yt, [c["id"] for c in active])
            fetched, fetched_ch = {}, set()
            for c in active:
                if c["id"] not in info:
                    sys.exit(f"YouTube returned nothing for channel {c['name']} ({c['id']}). Nothing was changed.")
                c["yt"] = info[c["id"]]["title"]
                if not c.get("name"):
                    c["name"] = c["yt"]
                ids = fetch_uploads(yt, info[c["id"]]["uploads"])
                det = fetch_details(yt, ids)
                had = sum(1 for v in videos if v.get("channel_id") == c["id"])
                if not det and had:
                    sys.exit(f"YouTube returned no videos for {c['name']} although the library has {had}. "
                             "Nothing was changed (this is usually a temporary API problem).")
                print(f"  {c['name']}  <-  YouTube channel '{c['yt']}': {len(ids)} uploads, {len(det)} public")
                fetched.update(det); fetched_ch.add(c["id"])
                c["refreshed"] = now
            videos, st = merge(videos, fetched, fetched_ch, now)
            print(f"New: {st['new']}   updated: {st['updated']}   IDs adopted for imported rows: {st['adopted']}   "
                  f"not in last refresh: {st['missing']}   (API calls: {yt.calls})")
        except ApiError as e:
            sys.exit(f"\nUpdate stopped, workbook untouched: {e}")
    for c in channels:
        c.setdefault("name", c.get("yt"))
    videos = finalize(videos, channels, tz)
    tmp = path.with_suffix(".tmp.xlsx")
    build(tmp, videos, channels, tzname, existing["selectors"], now)
    try:
        os.replace(tmp, path)
    except PermissionError:
        sys.exit(f"Close the workbook in Excel, then re-run. (New version saved as {tmp.name})")
    kept = sum(1 for v in videos if v.get("mark") or v.get("watched_on") or v.get("notes"))
    print(f"Done: {len(videos)} videos, {len(channels)} channels, {kept} videos with your marks/notes preserved -> {path.name}")


if __name__ == "__main__":
    main()

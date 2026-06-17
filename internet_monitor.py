"""
Internet Uptime Monitor
-----------------------
Pings 8.8.8.8 at a configurable interval (default: 5 min).
Stores per-minute availability in a local SQLite database.
Exports uptime/downtime reports to styled Excel files.
Runs silently in the background; minimizes to the system tray on close.
"""

import sys
import os
import time
import sqlite3
import threading
import subprocess
import platform
import datetime
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import queue

# ── Optional dependencies (graceful fallback messages) ──────────────────────
try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

# ────────────────────────────────────────────────────────────────────────────
DB_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uptime_data.db")
CONF_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
PING_HOST = "8.8.8.8"

# ── Default config ───────────────────────────────────────────────────────────
DEFAULT_CONFIG = {
    "interval_minutes": 5,
    "ping_count": 1,
    "ping_timeout": 4,
    "start_minimized": False,
}

# ════════════════════════════════════════════════════════════════════════════
#  Config helpers
# ════════════════════════════════════════════════════════════════════════════

def load_config() -> dict:
    if os.path.exists(CONF_PATH):
        try:
            with open(CONF_PATH) as f:
                cfg = {**DEFAULT_CONFIG, **json.load(f)}
            return cfg
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)

def save_config(cfg: dict):
    with open(CONF_PATH, "w") as f:
        json.dump(cfg, f, indent=2)

# ════════════════════════════════════════════════════════════════════════════
#  Database
# ════════════════════════════════════════════════════════════════════════════

def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pings (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            ts        INTEGER NOT NULL,   -- unix timestamp (UTC)
            online    INTEGER NOT NULL    -- 1 = up, 0 = down
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_ts ON pings(ts)")
    con.commit()
    con.close()

def record_ping(online: bool):
    ts = int(time.time())
    con = sqlite3.connect(DB_PATH)
    con.execute("INSERT INTO pings(ts, online) VALUES(?,?)", (ts, 1 if online else 0))
    con.commit()
    con.close()

def fetch_rows(start_ts: int, end_ts: int):
    """Return list of (ts, online) tuples in range [start_ts, end_ts)."""
    con = sqlite3.connect(DB_PATH)
    rows = con.execute(
        "SELECT ts, online FROM pings WHERE ts >= ? AND ts < ? ORDER BY ts",
        (start_ts, end_ts)
    ).fetchall()
    con.close()
    return rows

# ════════════════════════════════════════════════════════════════════════════
#  Ping logic  (uses OS ping only – no external libraries)
# ════════════════════════════════════════════════════════════════════════════

def do_ping(host: str, count: int = 1, timeout: int = 4) -> bool:
    system = platform.system().lower()

    # On Windows, subprocess.run() spawns a brand-new console window for
    # every ping.exe call. Tkinter hides its own window fine, but it has no
    # control over child-process windows, so that console flashes on screen
    # each time we ping unless we explicitly tell Windows not to create one.
    creationflags = 0
    startupinfo = None

    if system == "windows":
        cmd = ["ping", "-n", str(count), "-w", str(timeout * 1000), host]
        creationflags = subprocess.CREATE_NO_WINDOW
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
    else:
        cmd = ["ping", "-c", str(count), "-W", str(timeout), host]

    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout + 2,
            creationflags=creationflags,
            startupinfo=startupinfo,
        )
        return result.returncode == 0
    except Exception:
        return False

# ════════════════════════════════════════════════════════════════════════════
#  Statistics helpers
# ════════════════════════════════════════════════════════════════════════════

def _local_midnight(date: datetime.date) -> datetime.datetime:
    return datetime.datetime(date.year, date.month, date.day, 0, 0, 0)

def _to_ts(dt: datetime.datetime) -> int:
    return int(dt.timestamp())

def compute_stats(rows):
    """Given list of (ts, online), return (up_pings, down_pings, total_pings)."""
    if not rows:
        return 0, 0, 0
    up   = sum(1 for _, o in rows if o)
    down = len(rows) - up
    return up, down, len(rows)

def pct(up, total):
    return (up / total * 100) if total else 0.0

def stats_today():
    now   = datetime.datetime.now()
    start = _local_midnight(now.date())
    rows  = fetch_rows(_to_ts(start), _to_ts(now) + 1)
    return compute_stats(rows)

def stats_this_week():
    now        = datetime.datetime.now()
    week_start = now.date() - datetime.timedelta(days=now.weekday())
    start      = _local_midnight(week_start)
    rows       = fetch_rows(_to_ts(start), _to_ts(now) + 1)
    return compute_stats(rows)

def stats_this_month():
    now   = datetime.datetime.now()
    start = _local_midnight(now.date().replace(day=1))
    rows  = fetch_rows(_to_ts(start), _to_ts(now) + 1)
    return compute_stats(rows)

# ════════════════════════════════════════════════════════════════════════════
#  Excel export
# ════════════════════════════════════════════════════════════════════════════

# Colour palette
C_GREEN_DARK  = "1A7431"
C_GREEN_LIGHT = "C8E6C9"
C_RED_DARK    = "B71C1C"
C_RED_LIGHT   = "FFCDD2"
C_BLUE_DARK   = "0D47A1"
C_BLUE_MED    = "1565C0"
C_BLUE_LIGHT  = "E3F2FD"
C_HEADER_BG   = "1565C0"
C_HEADER_FG   = "FFFFFF"
C_TITLE_BG    = "0D47A1"
C_ALT_ROW     = "F5F9FF"
C_WHITE       = "FFFFFF"
C_GREY        = "ECEFF1"

def _border(style="thin"):
    s = Side(style=style, color="B0BEC5")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr_style(ws, cell_ref, text, bg=C_HEADER_BG, fg=C_HEADER_FG, bold=True, size=11):
    c = ws[cell_ref]
    c.value = text
    c.font  = Font(name="Calibri", bold=bold, color=fg, size=size)
    c.fill  = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = _border()

def _data_cell(ws, row, col, value, bg=C_WHITE, num_fmt=None, bold=False, color="000000", halign="center"):
    c = ws.cell(row=row, column=col, value=value)
    c.font  = Font(name="Calibri", bold=bold, color=color, size=10)
    c.fill  = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=halign, vertical="center")
    c.border = _border()
    if num_fmt:
        c.number_format = num_fmt

def _pct_color(p: float) -> str:
    """Green bg for high uptime, red for low."""
    if p >= 99:   return C_GREEN_LIGHT
    if p >= 95:   return "DCEDC8"
    if p >= 80:   return "FFF9C4"
    if p >= 60:   return "FFE0B2"
    return C_RED_LIGHT

def ping_interval_minutes(cfg):
    return cfg.get("interval_minutes", 5)

def _pings_to_time(pings: int, cfg: dict) -> str:
    """Convert ping count to human-readable duration."""
    interval = ping_interval_minutes(cfg)
    total_minutes = pings * interval
    h, m = divmod(total_minutes, 60)
    return f"{int(h)}h {int(m)}m"

def generate_excel(path: str, cfg: dict):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    now = datetime.datetime.now()

    # ── Build daily data (last 30 days) ─────────────────────────────────────
    daily_data = []
    for i in range(29, -1, -1):
        d = now.date() - datetime.timedelta(days=i)
        s = _local_midnight(d)
        e = s + datetime.timedelta(days=1)
        rows = fetch_rows(_to_ts(s), _to_ts(e))
        up, down, total = compute_stats(rows)
        daily_data.append({
            "date": d, "up": up, "down": down, "total": total,
            "pct": pct(up, total)
        })

    # ── Build weekly data (last 8 weeks) ─────────────────────────────────────
    weekly_data = []
    # Find the start of the current week (Monday)
    today = now.date()
    current_week_start = today - datetime.timedelta(days=today.weekday())
    for i in range(7, -1, -1):
        week_start = current_week_start - datetime.timedelta(weeks=i)
        week_end   = week_start + datetime.timedelta(days=7)
        s = _local_midnight(week_start)
        e = _local_midnight(week_end)
        rows = fetch_rows(_to_ts(s), _to_ts(e))
        up, down, total = compute_stats(rows)
        label = f"{week_start.strftime('%d %b')} – {(week_end - datetime.timedelta(days=1)).strftime('%d %b %Y')}"
        weekly_data.append({
            "label": label, "up": up, "down": down, "total": total,
            "pct": pct(up, total)
        })

    # ── Build monthly data (last 12 months) ──────────────────────────────────
    monthly_data = []
    for i in range(11, -1, -1):
        y = now.year
        m = now.month - i
        while m <= 0:
            m += 12
            y -= 1
        first_day  = datetime.date(y, m, 1)
        last_month = m % 12 + 1
        last_year  = y if m < 12 else y + 1
        last_day   = datetime.date(last_year, last_month, 1)
        s = _local_midnight(first_day)
        e = _local_midnight(last_day)
        rows = fetch_rows(_to_ts(s), _to_ts(e))
        up, down, total = compute_stats(rows)
        monthly_data.append({
            "label": first_day.strftime("%B %Y"),
            "up": up, "down": down, "total": total,
            "pct": pct(up, total)
        })

    # ════════════════════════════════════════════════════════════════════════
    #  Sheet 1 – Summary
    # ════════════════════════════════════════════════════════════════════════
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False

    # Title row
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = "🌐  Internet Uptime Monitor — Summary Report"
    title_cell.font  = Font(name="Calibri", bold=True, size=16, color=C_HEADER_FG)
    title_cell.fill  = PatternFill("solid", fgColor=C_TITLE_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:G2")
    sub = ws["A2"]
    sub.value = f"Generated: {now.strftime('%A, %d %B %Y  %H:%M:%S')}   |   Ping interval: {ping_interval_minutes(cfg)} min   |   Host: {PING_HOST}"
    sub.font  = Font(name="Calibri", italic=True, size=10, color="546E7A")
    sub.fill  = PatternFill("solid", fgColor="E8EAF6")
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 10  # spacer

    # Summary KPI cards (row 4)
    today_up, today_down, today_total = stats_today()
    week_up,  week_down,  week_total  = stats_this_week()
    month_up, month_down, month_total = stats_this_month()

    kpis = [
        ("TODAY", today_up, today_down, today_total, pct(today_up, today_total)),
        ("THIS WEEK", week_up,  week_down,  week_total,  pct(week_up, week_total)),
        ("THIS MONTH", month_up, month_down, month_total, pct(month_up, month_total)),
    ]

    col_offsets = [1, 3, 5]  # A, C, E  (2-col wide each)
    for idx, (label, up, down, total, p) in enumerate(kpis):
        c0 = col_offsets[idx]
        # Merge header
        ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 1)
        hc = ws.cell(row=4, column=c0)
        hc.value = label
        hc.font  = Font(name="Calibri", bold=True, size=11, color=C_HEADER_FG)
        hc.fill  = PatternFill("solid", fgColor=C_BLUE_MED)
        hc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[4].height = 22

        rows_def = [
            ("Uptime %",   f"{p:.2f}%",                     C_GREEN_LIGHT if p >= 95 else C_RED_LIGHT),
            ("Up time",    _pings_to_time(up, cfg),          C_WHITE),
            ("Down time",  _pings_to_time(down, cfg),        C_WHITE),
            ("Checks",     total if total else "No data yet", C_GREY),
        ]
        for r_off, (key, val, bg) in enumerate(rows_def):
            r = 5 + r_off
            ws.cell(row=r, column=c0, value=key).font = Font(name="Calibri", bold=True, size=10)
            ws.cell(row=r, column=c0).fill = PatternFill("solid", fgColor=bg)
            ws.cell(row=r, column=c0).alignment = Alignment(horizontal="right", vertical="center")
            ws.cell(row=r, column=c0).border = _border()
            ws.cell(row=r, column=c0 + 1, value=val).font = Font(name="Calibri", size=10, bold=(r_off == 0))
            ws.cell(row=r, column=c0 + 1).fill = PatternFill("solid", fgColor=bg)
            ws.cell(row=r, column=c0 + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r, column=c0 + 1).border = _border()
            ws.row_dimensions[r].height = 18

    # Column widths
    col_w = [18, 14, 18, 14, 18, 14, 6]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ════════════════════════════════════════════════════════════════════════
    #  Sheet 2 – Daily (last 30 days)
    # ════════════════════════════════════════════════════════════════════════
    def make_data_sheet(wb, title, data_rows, col_headers, row_key):
        ws2 = wb.create_sheet(title)
        ws2.sheet_view.showGridLines = False

        # Title
        ncols = len(col_headers)
        ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        tc = ws2.cell(row=1, column=1)
        tc.value = f"📊  {title}"
        tc.font  = Font(name="Calibri", bold=True, size=14, color=C_HEADER_FG)
        tc.fill  = PatternFill("solid", fgColor=C_TITLE_BG)
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[1].height = 32

        ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        sub2 = ws2.cell(row=2, column=1)
        sub2.value = f"Host: {PING_HOST}  |  Interval: {ping_interval_minutes(cfg)} min  |  Generated: {now.strftime('%d %b %Y %H:%M')}"
        sub2.font  = Font(name="Calibri", italic=True, size=9, color="546E7A")
        sub2.fill  = PatternFill("solid", fgColor="E8EAF6")
        sub2.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[2].height = 16

        ws2.row_dimensions[3].height = 8  # spacer

        # Headers row 4
        for ci, hdr in enumerate(col_headers, 1):
            c = ws2.cell(row=4, column=ci)
            c.value = hdr
            c.font  = Font(name="Calibri", bold=True, size=10, color=C_HEADER_FG)
            c.fill  = PatternFill("solid", fgColor=C_HEADER_BG)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = _border()
        ws2.row_dimensions[4].height = 28

        for ri, row in enumerate(data_rows):
            excel_r = ri + 5
            bg = C_ALT_ROW if ri % 2 == 0 else C_WHITE

            label   = row.get("date", row.get("label", ""))
            if hasattr(label, "strftime"):
                label_str = label.strftime("%a %d %b %Y")
            else:
                label_str = str(label)

            up    = row["up"]
            down  = row["down"]
            total = row["total"]
            p     = row["pct"]
            up_t  = _pings_to_time(up, cfg)
            dn_t  = _pings_to_time(down, cfg)

            pct_bg = _pct_color(p) if total else C_GREY

            values = [
                (label_str,            bg,    None,    False, "000000", "left"),
                (total if total else 0, bg,   "#,##0", False, "000000", "center"),
                (up    if total else 0, bg,   "#,##0", False, C_GREEN_DARK, "center"),
                (down  if total else 0, bg,   "#,##0", False, C_RED_DARK,   "center"),
                (up_t  if total else "—", bg, None,    False, C_GREEN_DARK, "center"),
                (dn_t  if total else "—", bg, None,    False, C_RED_DARK,   "center"),
                (p / 100 if total else None, pct_bg, "0.00%", True,
                 C_GREEN_DARK if p >= 80 else C_RED_DARK, "center"),
            ]

            for ci, (val, cbg, fmt, bld, col, hal) in enumerate(values, 1):
                _data_cell(ws2, excel_r, ci, val, cbg, fmt, bld, col, hal)
            ws2.row_dimensions[excel_r].height = 17

        # Totals row
        last_r  = len(data_rows) + 5
        tot_row = last_r
        total_up   = sum(r["up"]   for r in data_rows)
        total_down = sum(r["down"] for r in data_rows)
        total_all  = sum(r["total"]for r in data_rows)
        total_pct  = pct(total_up, total_all)

        totals = [
            ("TOTAL / AVG",              C_BLUE_LIGHT, None,    True, C_BLUE_DARK, "right"),
            (total_all,                  C_BLUE_LIGHT, "#,##0", True, C_BLUE_DARK, "center"),
            (total_up,                   C_BLUE_LIGHT, "#,##0", True, C_GREEN_DARK,"center"),
            (total_down,                 C_BLUE_LIGHT, "#,##0", True, C_RED_DARK,  "center"),
            (_pings_to_time(total_up,   cfg), C_BLUE_LIGHT, None, True, C_GREEN_DARK,"center"),
            (_pings_to_time(total_down, cfg), C_BLUE_LIGHT, None, True, C_RED_DARK, "center"),
            (total_pct / 100,            C_BLUE_LIGHT, "0.00%", True,
             C_GREEN_DARK if total_pct >= 80 else C_RED_DARK, "center"),
        ]
        for ci, (val, cbg, fmt, bld, col, hal) in enumerate(totals, 1):
            _data_cell(ws2, tot_row, ci, val, cbg, fmt, bld, col, hal)
        ws2.row_dimensions[tot_row].height = 20

        # Column widths
        widths = [20, 10, 10, 10, 12, 12, 11]
        for i, w in enumerate(widths, 1):
            ws2.column_dimensions[get_column_letter(i)].width = w

        return ws2

    col_headers = ["Period", "Total Checks", "Up Checks", "Down Checks",
                   "Uptime (hh mm)", "Downtime (hh mm)", "Uptime %"]

    make_data_sheet(wb, "Daily (Last 30 Days)", daily_data,  col_headers, "date")
    make_data_sheet(wb, "Weekly (Last 8 Weeks)", weekly_data, col_headers, "label")
    make_data_sheet(wb, "Monthly (Last 12 Mo.)", monthly_data,col_headers, "label")

    # ════════════════════════════════════════════════════════════════════════
    #  Sheet 5 – Legend / Help
    # ════════════════════════════════════════════════════════════════════════
    wsl = wb.create_sheet("Help & Legend")
    wsl.sheet_view.showGridLines = False
    wsl.column_dimensions["A"].width = 22
    wsl.column_dimensions["B"].width = 50

    wsl.merge_cells("A1:B1")
    lc = wsl["A1"]
    lc.value = "📘  Help & Legend"
    lc.font  = Font(name="Calibri", bold=True, size=14, color=C_HEADER_FG)
    lc.fill  = PatternFill("solid", fgColor=C_TITLE_BG)
    lc.alignment = Alignment(horizontal="center", vertical="center")
    wsl.row_dimensions[1].height = 30

    legend = [
        ("Colour", "Meaning"),
        ("🟢 Green bg (uptime %)", "≥ 95% uptime — excellent"),
        ("🟡 Yellow-green",        "80–94% uptime — good"),
        ("🟠 Orange",              "60–79% uptime — degraded"),
        ("🔴 Red bg",              "< 60% uptime — poor"),
        ("", ""),
        ("Column", "Description"),
        ("Total Checks",    f"Number of pings performed (1 per {ping_interval_minutes(cfg)} min)"),
        ("Up Checks",       "Pings that received a reply from 8.8.8.8"),
        ("Down Checks",     "Pings with no reply (timeout or unreachable)"),
        ("Uptime (hh mm)",  "Total connected time estimated from up pings"),
        ("Downtime (hh mm)","Total disconnected time estimated from down pings"),
        ("Uptime %",        "Up Checks ÷ Total Checks × 100"),
        ("", ""),
        ("Notes", ""),
        ("Data source",  f"Local SQLite database: {DB_PATH}"),
        ("Ping host",    PING_HOST + "  (Google Public DNS — no tracking)"),
        ("Interval",     f"{ping_interval_minutes(cfg)} minutes between checks"),
        ("Time zone",    "All times are local system time"),
    ]

    for ri, (col_a, col_b) in enumerate(legend, 2):
        a = wsl.cell(row=ri, column=1, value=col_a)
        b = wsl.cell(row=ri, column=2, value=col_b)
        for c in (a, b):
            c.font   = Font(name="Calibri", size=10, bold=(col_b == "Meaning" or col_b == "Description"))
            c.border = _border()
            c.alignment = Alignment(vertical="center", wrap_text=True)
            if col_a in ("Colour", "Column", "Notes"):
                c.fill = PatternFill("solid", fgColor=C_BLUE_LIGHT)
                c.font = Font(name="Calibri", bold=True, size=10, color=C_BLUE_DARK)
            else:
                c.fill = PatternFill("solid", fgColor=C_WHITE if ri % 2 else C_GREY)
        wsl.row_dimensions[ri].height = 18

    # ── Reorder sheets ───────────────────────────────────────────────────────
    ordered = ["Summary", "Daily (Last 30 Days)", "Weekly (Last 8 Weeks)",
               "Monthly (Last 12 Mo.)", "Help & Legend"]
    wb._sheets.sort(key=lambda s: (ordered.index(s.title) if s.title in ordered else 99))

    wb.save(path)

# ════════════════════════════════════════════════════════════════════════════
#  Monitoring thread
# ════════════════════════════════════════════════════════════════════════════

class Monitor:
    def __init__(self, cfg: dict, event_queue: queue.Queue):
        self.cfg   = cfg
        self.q     = event_queue
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True, name="MonitorThread")

    def start(self):
        self._thread.start()

    def stop(self):
        self._stop.set()

    def update_cfg(self, cfg):
        self.cfg = cfg

    def _run(self):
        while not self._stop.is_set():
            online = do_ping(PING_HOST,
                             self.cfg.get("ping_count", 1),
                             self.cfg.get("ping_timeout", 4))
            record_ping(online)
            self.q.put(("PING", online))
            interval_secs = self.cfg.get("interval_minutes", 5) * 60
            # Sleep in short bursts so we can react to stop quickly
            deadline = time.time() + interval_secs
            while time.time() < deadline and not self._stop.is_set():
                time.sleep(2)

# ════════════════════════════════════════════════════════════════════════════
#  System tray icon (pystray)
# ════════════════════════════════════════════════════════════════════════════

def _make_tray_icon(online: bool) -> "Image":
    """Draw a simple coloured circle as the tray icon."""
    img  = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    color = (34, 197, 94) if online else (239, 68, 68)
    draw.ellipse([4, 4, 60, 60], fill=color, outline=(255, 255, 255), width=4)
    return img

# ════════════════════════════════════════════════════════════════════════════
#  GUI
# ════════════════════════════════════════════════════════════════════════════

DARK_BG   = "#1E2330"
PANEL_BG  = "#252B3B"
CARD_BG   = "#2D3447"
ACCENT    = "#3B82F6"
GREEN     = "#22C55E"
RED       = "#EF4444"
FG_WHITE  = "#F1F5F9"
FG_MUTED  = "#94A3B8"
MONO      = "Consolas" if platform.system() == "Windows" else "Courier"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.cfg     = load_config()
        self.q       = queue.Queue()
        self._online = True
        self._tray   = None
        self._tray_thread = None
        self.log_lines = []

        self.title("Internet Uptime Monitor")
        self.configure(bg=DARK_BG)
        self.resizable(False, False)
        self._build_ui()
        self._center()

        # Start monitor
        self.monitor = Monitor(self.cfg, self.q)
        self.monitor.start()

        # Start polling the event queue
        self.after(500, self._poll_queue)
        self.after(1000, self._refresh_stats)

        # Minimise to tray on close
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        if self.cfg.get("start_minimized"):
            self.after(100, self._hide)

    # ── UI construction ──────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ──
        header = tk.Frame(self, bg=ACCENT, height=50)
        header.pack(fill="x")
        tk.Label(header, text="🌐  Internet Uptime Monitor",
                 bg=ACCENT, fg=FG_WHITE,
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=16, pady=12)

        self._status_lbl = tk.Label(header, text="● CHECKING…",
                                    bg=ACCENT, fg=FG_WHITE,
                                    font=("Segoe UI", 10, "bold"))
        self._status_lbl.pack(side="right", padx=16)

        # ── Stats cards row ──────────────────────────────────────────────────
        cards_frame = tk.Frame(self, bg=DARK_BG)
        cards_frame.pack(fill="x", padx=12, pady=(10, 4))

        self._card_today = self._make_card(cards_frame, "TODAY", "—", "—")
        self._card_week  = self._make_card(cards_frame, "THIS WEEK", "—", "—")
        self._card_month = self._make_card(cards_frame, "THIS MONTH", "—", "—")

        # ── Log area ────────────────────────────────────────────────────────
        log_frame = tk.Frame(self, bg=PANEL_BG, bd=0)
        log_frame.pack(fill="both", expand=True, padx=12, pady=4)

        tk.Label(log_frame, text="📋  Ping Log", bg=PANEL_BG,
                 fg=FG_MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=8, pady=(6, 2))

        self._log_text = tk.Text(log_frame, bg=CARD_BG, fg=FG_WHITE,
                                  font=(MONO, 9), height=12, bd=0,
                                  wrap="none", state="disabled",
                                  insertbackground=FG_WHITE)
        self._log_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self._log_text.tag_config("up",   foreground=GREEN)
        self._log_text.tag_config("down", foreground=RED)
        self._log_text.tag_config("info", foreground=FG_MUTED)

        # ── Bottom controls ──────────────────────────────────────────────────
        ctrl = tk.Frame(self, bg=DARK_BG)
        ctrl.pack(fill="x", padx=12, pady=(4, 10))

        # Interval
        tk.Label(ctrl, text="Check interval:", bg=DARK_BG,
                 fg=FG_MUTED, font=("Segoe UI", 9)).pack(side="left")
        self._interval_var = tk.StringVar(value=str(self.cfg.get("interval_minutes", 5)))
        interval_spin = tk.Spinbox(ctrl, from_=1, to=60,
                                   textvariable=self._interval_var,
                                   width=4, font=("Segoe UI", 9),
                                   bg=CARD_BG, fg=FG_WHITE,
                                   buttonbackground=CARD_BG,
                                   insertbackground=FG_WHITE)
        interval_spin.pack(side="left", padx=(4, 2))
        tk.Label(ctrl, text="min", bg=DARK_BG,
                 fg=FG_MUTED, font=("Segoe UI", 9)).pack(side="left", padx=(0, 10))

        btn_apply = tk.Button(ctrl, text="Apply", command=self._apply_interval,
                              bg=ACCENT, fg=FG_WHITE, font=("Segoe UI", 9, "bold"),
                              relief="flat", padx=8, cursor="hand2")
        btn_apply.pack(side="left", padx=(0, 12))

        btn_export = tk.Button(ctrl, text="⬇ Export Excel", command=self._export,
                               bg="#16A34A", fg=FG_WHITE, font=("Segoe UI", 9, "bold"),
                               relief="flat", padx=10, cursor="hand2")
        btn_export.pack(side="right")

        btn_hide = tk.Button(ctrl, text="Hide to Tray", command=self._hide,
                             bg=PANEL_BG, fg=FG_MUTED, font=("Segoe UI", 9),
                             relief="flat", padx=8, cursor="hand2")
        btn_hide.pack(side="right", padx=(0, 8))

    def _make_card(self, parent, title, up_val, pct_val):
        f = tk.Frame(parent, bg=CARD_BG, bd=0, padx=14, pady=10)
        f.pack(side="left", expand=True, fill="both", padx=4)

        tk.Label(f, text=title, bg=CARD_BG, fg=FG_MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        pct_lbl = tk.Label(f, text=pct_val, bg=CARD_BG, fg=GREEN,
                            font=("Segoe UI", 18, "bold"))
        pct_lbl.pack(anchor="w")
        sub_lbl = tk.Label(f, text=up_val, bg=CARD_BG, fg=FG_MUTED,
                            font=("Segoe UI", 8))
        sub_lbl.pack(anchor="w")
        return pct_lbl, sub_lbl

    def _center(self):
        self.update_idletasks()
        w, h = 640, 520
        sw   = self.winfo_screenwidth()
        sh   = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ── Event loop ───────────────────────────────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "PING":
                    self._online = data
                    self._update_status(data)
        except queue.Empty:
            pass
        self.after(500, self._poll_queue)

    def _update_status(self, online: bool):
        ts   = datetime.datetime.now().strftime("%H:%M:%S")
        text = f"[{ts}]  ●  {'ONLINE' if online else 'OFFLINE'}  —  {PING_HOST}"
        tag  = "up" if online else "down"
        self._append_log(text, tag)

        if online:
            self._status_lbl.config(text="● ONLINE", fg="#86EFAC")
        else:
            self._status_lbl.config(text="● OFFLINE", fg=RED)

        # Update tray icon
        if TRAY_AVAILABLE and self._tray:
            try:
                self._tray.icon = _make_tray_icon(online)
            except Exception:
                pass

    def _append_log(self, text: str, tag: str = "info"):
        self._log_text.config(state="normal")
        self._log_text.insert("end", text + "\n", tag)
        self._log_text.see("end")
        # Keep last 500 lines
        lines = int(self._log_text.index("end-1c").split(".")[0])
        if lines > 500:
            self._log_text.delete("1.0", f"{lines-500}.0")
        self._log_text.config(state="disabled")

    def _refresh_stats(self):
        def fmt(up, down, total):
            p = pct(up, total)
            interval = self.cfg.get("interval_minutes", 5)
            up_m  = up * interval
            dn_m  = down * interval
            up_str = f"{up_m//60}h {up_m%60}m up  /  {dn_m//60}h {dn_m%60}m down"
            return f"{p:.1f}%", up_str

        for fn, card in [(stats_today, self._card_today),
                         (stats_this_week,  self._card_week),
                         (stats_this_month, self._card_month)]:
            up, down, total = fn()
            pct_str, detail = fmt(up, down, total) if total else ("—", "No data yet")
            pct_lbl, sub_lbl = card
            pct_lbl.config(text=pct_str,
                           fg=GREEN if total and pct(up, total) >= 80 else RED)
            sub_lbl.config(text=detail)

        self.after(30_000, self._refresh_stats)  # refresh every 30 s

    # ── Actions ──────────────────────────────────────────────────────────────
    def _apply_interval(self):
        try:
            val = int(self._interval_var.get())
            if val < 1 or val > 60:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid", "Interval must be 1–60 minutes.")
            return
        self.cfg["interval_minutes"] = val
        save_config(self.cfg)
        self.monitor.update_cfg(self.cfg)
        self._append_log(f"[CONFIG]  Interval set to {val} min.", "info")

    def _export(self):
        if not XLSX_AVAILABLE:
            messagebox.showerror(
                "Missing library",
                "openpyxl is not installed.\n\nRun:\n  pip install openpyxl"
            )
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            initialfile=f"uptime_report_{datetime.date.today()}.xlsx",
            title="Save Uptime Report"
        )
        if not path:
            return
        try:
            generate_excel(path, self.cfg)
            self._append_log(f"[EXPORT]  Report saved → {path}", "info")
            messagebox.showinfo("Exported", f"Report saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    # ── Tray / hide ──────────────────────────────────────────────────────────
    def _hide(self):
        self.withdraw()
        if TRAY_AVAILABLE and self._tray is None:
            self._start_tray()

    def _show(self):
        self.deiconify()
        self.lift()
        self.focus_force()

    def _on_close(self):
        self._hide()

    def _start_tray(self):
        icon_img = _make_tray_icon(self._online)
        menu = pystray.Menu(
            pystray.MenuItem("Open", self._show, default=True),
            pystray.MenuItem("Export Excel…", lambda _i, _it: self.after(0, self._export)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit_app),
        )
        self._tray = pystray.Icon("UptimeMonitor", icon_img,
                                  "Internet Uptime Monitor", menu)
        self._tray_thread = threading.Thread(target=self._tray.run,
                                             daemon=True, name="TrayThread")
        self._tray_thread.start()

    def _quit_app(self, *_):
        self.monitor.stop()
        if self._tray:
            try:
                self._tray.stop()
            except Exception:
                pass
        self.after(0, self.destroy)

# ════════════════════════════════════════════════════════════════════════════
#  Entry point
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # Warn about optional deps
    missing = []
    if not TRAY_AVAILABLE:
        missing.append("pystray  Pillow   (system tray)")
    if not XLSX_AVAILABLE:
        missing.append("openpyxl          (Excel export)")
    if missing:
        print("⚠  Optional packages not installed — some features disabled.")
        print("   Install with:")
        for m in missing:
            print(f"     pip install {m.split()[0].lower()}")
        print()

    init_db()
    app = App()
    app.mainloop()

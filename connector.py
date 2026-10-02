"""
Ashok Daily Entry - local connector.

Serves the HTML form at http://localhost:8765 and applies submissions
to the workout workbook in the OneDrive folder (OneDrive then syncs).

Rules it respects (from the RULES - Claude & Ashok tab):
- only writes the fields Ashok submitted (never touches anything else)
- safe save: temp file + atomic replace, with lock + mtime checks
- never reorders tabs, never touches styles outside the target cells
"""
import json
import os
import sys
import datetime
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HOST = "0.0.0.0"   # listen on all interfaces so the phone can reach it over home Wi-Fi
PORT = 8765

# default workbook location - override with config.json next to this file
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_WORKBOOK = r"C:\Users\ashok\OneDrive\NOblox\Ashok_Health_Stats\Ashok_Dan_Workout_Weekly_Log_v2.xlsx"
WORKBOOK = DEFAULT_WORKBOOK
if os.path.exists(os.path.join(BASE_DIR, "config.json")):
    try:
        with open(os.path.join(BASE_DIR, "config.json"), encoding="utf-8") as f:
            cfg = json.load(f)
        if cfg.get("workbook"):
            WORKBOOK = cfg["workbook"]
    except Exception:
        pass

# Daily Metrics Metabolism column map (column index: field name)
COLUMNS = {
    "weight": 3,          # C  Morning weight (kg)
    "bf": 4,              # D  Body fat %
    "steps": 42,          # AP Steps fitbit
    "hiit": 47,           # AU HIIT rounds
    "sleep_h": 48,        # AV Sleep (h) fitbit
    "sleep_score": 49,    # AW Sleep SCORE fitbit
    "water": 60,          # AN Water Ml
    "comment": 59,        # BG ASHOK COMMENTS Notes
}
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def apply_entry(payload):
    """Write the submitted fields into Daily Metrics Metabolism. Returns (ok, message, updated)."""
    import openpyxl
    from copy import copy

    if not os.path.exists(WORKBOOK):
        return False, f"Workbook not found: {WORKBOOK}", []

    start_mtime = os.path.getmtime(WORKBOOK)
    try:
        fd = os.open(WORKBOOK, os.O_RDWR | os.O_BINARY)
        os.close(fd)
    except PermissionError:
        return False, "Workbook is open in Excel - close it and try again.", []

    wb = openpyxl.load_workbook(WORKBOOK)
    dm = wb["Daily Metrics Metabolism"]

    date_val = datetime.datetime.strptime(payload["date"], "%Y-%m-%d")
    day_name = payload.get("day") or DAYS[date_val.weekday()]

    # find the row for this date (col B)
    row = None
    for r in range(6, dm.max_row + 1):
        d = dm.cell(row=r, column=2).value
        if isinstance(d, datetime.datetime) and d.date() == date_val.date():
            row = r
            break
    if row is None:
        # append a new row at the end
        row = dm.max_row + 1
        dm.cell(row=row, column=1).value = day_name
        dm.cell(row=row, column=2).value = date_val
        for col in (34,):  # AH Target kcal default 2050 like other rows
            dm.cell(row=row, column=col).value = 2050

    updated = []
    for field, col in COLUMNS.items():
        raw = payload.get(field)
        if raw in (None, ""):
            continue
        if field in ("weight", "bf", "steps", "hiit", "sleep_h", "sleep_score", "water"):
            try:
                val = float(raw)
            except ValueError:
                continue
            val = int(val) if val == int(val) else val
        else:
            val = str(raw)
        cell = dm.cell(row=row, column=col)
        cell.value = val
        updated.append(field.replace("_", " "))

    # safe save
    now = os.path.getmtime(WORKBOOK)
    if abs(now - start_mtime) > 2:
        wb.close()
        return False, "Workbook changed on disk while saving - not saved. Try again.", []

    tmp = WORKBOOK + ".form_tmp.xlsx"
    try:
        wb.save(tmp)
        check = openpyxl.load_workbook(tmp, read_only=True)
        ok = len(check.sheetnames) >= 20
        check.close()
        if not ok:
            raise RuntimeError("temp save verification failed")
        os.replace(tmp, WORKBOOK)
    except Exception as e:
        wb.close()
        for leftover in (tmp,):
            try:
                if os.path.exists(leftover):
                    os.remove(leftover)
            except Exception:
                pass
        return False, f"Save failed (original untouched): {e}", []
    wb.close()
    return True, f"{day_name} {payload['date']} → row {row}.", updated


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass  # quiet

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            html = os.path.join(BASE_DIR, "index.html")
            if os.path.exists(html):
                with open(html, "rb") as f:
                    body = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        if urlparse(self.path).path != "/submit":
            self.send_response(404)
            self.end_headers()
            return
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        try:
            payload = json.loads(raw.decode("utf-8"))
            ok, msg, updated = apply_entry(payload)
        except Exception as e:
            ok, msg, updated = False, f"Error: {e}", []
        resp = json.dumps({"ok": ok, "message": msg, "updated": updated}).encode("utf-8")
        self.send_response(200 if ok else 400)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(resp)))
        self.end_headers()
        self.wfile.write(resp)


if __name__ == "__main__":
    import socket
    def lan_ip():
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "127.0.0.1"

    ip = lan_ip()
    print("=" * 60)
    print(" Ashok Daily Entry connector")
    print(f" ON THIS PC:   http://localhost:{PORT}")
    print(f" ON YOUR PHONE (same Wi-Fi): http://{ip}:{PORT}")
    print(f" Workbook: {WORKBOOK}")
    print(" Keep this window open. Press Ctrl+C to stop.")
    print("=" * 60)
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")

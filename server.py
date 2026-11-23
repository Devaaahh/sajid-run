"""Sajid Run - game server.

Serves the game in ./game and keeps the high-score leaderboard.

Entry points:
  * `app`  - a plain WSGI application. Vercel's Python runtime picks it up
             (see vercel.json); any other WSGI host (gunicorn, waitress...) works too.
  * main() - local development server:
        python server.py            # then open http://localhost:8000
        python server.py 9000       # use another port

Storage:
  * Supabase table `scores` (when SUPABASE_URL and SUPABASE_KEY are set in .env or
    in the Vercel project's environment variables)
  * Excel workbook scores.xlsx - rewritten after every saved score (locally), and
    generated on demand at /api/scores.xlsx (everywhere)
  * scores.json - fallback when Supabase is not configured / unreachable
    (on Vercel this lives in /tmp and is wiped regularly, so use Supabase there)
"""
import io
import json
import mimetypes
import os
import sys
import threading
import urllib.error
import urllib.request
import webbrowser
from datetime import datetime
from pathlib import Path
from socketserver import ThreadingMixIn
from wsgiref.simple_server import WSGIRequestHandler, WSGIServer, make_server

ROOT = Path(__file__).resolve().parent
GAME_DIR = ROOT / "game"
ON_VERCEL = bool(os.environ.get("VERCEL"))
DATA_DIR = Path("/tmp") if ON_VERCEL else ROOT  # Vercel's filesystem is read-only except /tmp
SCORES_JSON = DATA_DIR / "scores.json"
SCORES_XLSX = ROOT / "scores.xlsx"
TOP_N = 10
MAX_BODY = 4096
_lock = threading.Lock()

mimetypes.add_type("application/javascript", ".js")
mimetypes.add_type("text/css", ".css")
mimetypes.add_type("audio/mpeg", ".mp3")


# ---------------------------------------------------------------- config
def load_env():
    env_file = ROOT / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


load_env()
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY") or os.environ.get("SUPABASE_PUBLISHABLE_KEY", "")
SUPABASE_TABLE = os.environ.get("SUPABASE_TABLE", "scores")


# ---------------------------------------------------------------- supabase
def supabase_enabled() -> bool:
    return bool(SUPABASE_URL and SUPABASE_KEY)


def _supabase(method: str, query: str = "", body=None):
    req = urllib.request.Request(
        f"{SUPABASE_URL}/rest/v1/{SUPABASE_TABLE}{query}",
        method=method,
        data=json.dumps(body).encode("utf-8") if body is not None else None,
        headers={
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json",
            "Prefer": "return=representation",
        },
    )
    with urllib.request.urlopen(req, timeout=8) as resp:
        return json.loads(resp.read() or b"[]")


def supabase_rows(limit: int):
    return _supabase("GET", f"?select=name,score,coins,distance,created_at&order=score.desc&limit={limit}")


# ---------------------------------------------------------------- local json
def json_all() -> list:
    try:
        data = json.loads(SCORES_JSON.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (OSError, ValueError):
        return []


def json_add(entry: dict):
    scores = json_all()
    scores.append(entry)
    scores.sort(key=lambda s: s["score"], reverse=True)
    SCORES_JSON.write_text(json.dumps(scores, indent=2), encoding="utf-8")


# ---------------------------------------------------------------- excel
def excel_bytes(rows: list, source: str) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "High Scores"
    ws.append(["Rank", "Name", "Score", "Coins", "Distance (m)", "Date"])
    head_fill = PatternFill("solid", fgColor="FF7A00")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")
    for i, r in enumerate(rows, 1):
        ws.append([i, r.get("name"), r.get("score"), r.get("coins"), r.get("distance"),
                   (r.get("created_at") or "")[:19].replace("T", " ")])
    for col, width in zip("ABCDEF", (7, 20, 12, 10, 14, 22)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    info = wb.create_sheet("Info")
    info.append(["Game", "Sajid Run - can he escape from Pranav and Ashwin?"])
    info.append(["Source", source])
    info.append(["Updated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def write_excel_file(rows: list, source: str):
    if ON_VERCEL:
        return  # read-only filesystem; /api/scores.xlsx builds it on demand instead
    try:
        SCORES_XLSX.write_bytes(excel_bytes(rows, source))
    except ImportError:
        print("openpyxl not installed - skipping scores.xlsx (pip install openpyxl)")
    except PermissionError:
        print("scores.xlsx is open in Excel - close it so it can be updated")


# ---------------------------------------------------------------- leaderboard
def all_rows():
    """Every score, best first, plus a label saying where they came from."""
    if supabase_enabled():
        try:
            return supabase_rows(10000), f"Supabase table '{SUPABASE_TABLE}'"
        except (urllib.error.URLError, ValueError, OSError) as e:
            print("Supabase read failed, using local scores:", e)
    return json_all(), "local scores.json"


def get_top() -> list:
    if supabase_enabled():
        try:
            return supabase_rows(TOP_N)
        except (urllib.error.URLError, ValueError, OSError) as e:
            print("Supabase read failed, using local scores:", e)
    return json_all()[:TOP_N]


def add_score(entry: dict) -> list:
    with _lock:
        rows, source = None, "local scores.json"
        if supabase_enabled():
            try:
                _supabase("POST", "", entry)
                rows, source = supabase_rows(10000), f"Supabase table '{SUPABASE_TABLE}'"
            except (urllib.error.URLError, ValueError, OSError) as e:
                print("Supabase write failed, saving locally:", e)
        if rows is None:
            json_add(entry)
            rows = json_all()
        write_excel_file(rows, source)
        return rows[:TOP_N]


def parse_entry(raw: bytes) -> dict:
    data = json.loads(raw or b"{}")
    if not isinstance(data, dict):
        raise ValueError("expected an object")
    return {
        "name": str(data.get("name", "")).strip()[:16] or "Anonymous",
        "score": max(0, int(data.get("score", 0))),
        "coins": max(0, int(data.get("coins", 0))),
        "distance": max(0, int(data.get("distance", 0))),
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }


# ---------------------------------------------------------------- WSGI app (Vercel entry point)
STATUS = {200: "200 OK", 400: "400 Bad Request", 404: "404 Not Found",
          405: "405 Method Not Allowed", 500: "500 Internal Server Error"}


def wsgi_app(environ, start_response):
    method = environ.get("REQUEST_METHOD", "GET").upper()
    path = environ.get("PATH_INFO") or "/"

    def reply(code, body: bytes, ctype, extra=()):
        headers = [("Content-Type", ctype), ("Content-Length", str(len(body)))] + list(extra)
        start_response(STATUS[code], headers)
        return [b"" if method == "HEAD" else body]

    def reply_json(payload, code=200):
        return reply(code, json.dumps(payload).encode("utf-8"), "application/json",
                     [("Cache-Control", "no-store")])

    try:
        if path == "/api/scores":
            if method in ("GET", "HEAD"):
                return reply_json(get_top())
            if method == "POST":
                try:
                    length = min(int(environ.get("CONTENT_LENGTH") or 0), MAX_BODY)
                    entry = parse_entry(environ["wsgi.input"].read(length))
                except (ValueError, TypeError):
                    return reply_json({"error": "bad request"}, 400)
                return reply_json(add_score(entry))
            return reply_json({"error": "method not allowed"}, 405)

        if path == "/api/scores.xlsx":
            rows, source = all_rows()
            if not rows:
                return reply_json({"error": "no scores yet"}, 404)
            return reply(200, excel_bytes(rows, source),
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         [("Content-Disposition", 'attachment; filename="sajid-run-scores.xlsx"'),
                          ("Cache-Control", "no-store")])

        # static game files
        if method not in ("GET", "HEAD"):
            return reply_json({"error": "method not allowed"}, 405)
        rel = path.lstrip("/") or "index.html"
        target = (GAME_DIR / rel).resolve()
        if target.is_dir():
            target = target / "index.html"
        if GAME_DIR.resolve() not in target.parents or not target.is_file():
            return reply(404, b"Not found", "text/plain; charset=utf-8")
        ctype = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype == "application/javascript":
            ctype += "; charset=utf-8"
        cache = "no-cache" if target.suffix in (".html", ".js", ".css") else "public, max-age=86400"
        return reply(200, target.read_bytes(), ctype, [("Cache-Control", cache)])
    except Exception as e:  # never leak a traceback to the browser
        print("Server error:", repr(e))
        return reply_json({"error": "server error"}, 500)


# Vercel looks for a top-level variable named "app" (or "application" / "handler").
app = application = wsgi_app


# ---------------------------------------------------------------- local dev server
class _ThreadingWSGIServer(ThreadingMixIn, WSGIServer):
    daemon_threads = True


class _QuietHandler(WSGIRequestHandler):
    def log_message(self, fmt, *args):
        if args and "/api/" in str(args[0]):
            super().log_message(fmt, *args)


def ensure_assets():
    needed = ["runner.png", "pranav.png", "ashwin.png", "start.mp3", "run.mp3", "caught.mp3"]
    if all((GAME_DIR / "assets" / n).exists() for n in needed):
        return
    print("Preparing game assets...")
    import prepare_assets  # needs Pillow

    prepare_assets.main()


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    port = int(args[0]) if args else 8000
    ensure_assets()
    for p in range(port, port + 20):
        try:
            server = make_server("127.0.0.1", p, app,
                                 server_class=_ThreadingWSGIServer, handler_class=_QuietHandler)
            break
        except OSError:
            continue
    else:
        sys.exit("No free port found")
    url = f"http://localhost:{server.server_address[1]}"
    print(f"SAJID RUN is live at {url}  (Ctrl+C to stop)")
    print("High scores ->", f"Supabase ({SUPABASE_URL})" if supabase_enabled() else "scores.json",
          "+ scores.xlsx")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nBye!")


if __name__ == "__main__":
    main()

"""
CTFN Excel Server â REST API for the =CTFN() Excel UDF.
Deployed on Render. Bridges to the Pressrisk/CTFN WebSocket API for
M&A news, Active Situations, and DMA data.

Architecture:
  =CTFN("headline","ROKU")  â  VBA UDF  â  HTTPS GET ctfn.onrender.com  â  this server  â  wss://api.pressrisk.io/

Environment variables:
  CTFN_USERS           Comma-separated user:password pairs (e.g. deane:pass1,fahim:pass2)
  CTFN_TOKEN_HOURS     Token expiry in hours (default 24)
  CTFN_WS_USERNAME     Pressrisk WebSocket username
  CTFN_WS_PASSWORD     Pressrisk WebSocket password
  HOST                 Bind address (default 0.0.0.0)
  PORT                 Server port (set by Render)
"""

import os
import sys
import json
import time
import datetime
import hashlib
import secrets
import logging
import threading
import re
from typing import Any, Optional

from fastapi import FastAPI, Query, Header, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("ctfn")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
PORT = int(os.environ.get("PORT", 8002))

# ---------------------------------------------------------------------------
# User authentication (mirrors CTFNDATA architecture)
# ---------------------------------------------------------------------------
USERS_RAW = os.environ.get("CTFN_USERS", "")
USERS: dict = {}  # username â password hash
if USERS_RAW:
    for pair in USERS_RAW.split(","):
        pair = pair.strip()
        if ":" in pair:
            user, pw = pair.split(":", 1)
            USERS[user.strip().lower()] = hashlib.sha256(pw.strip().encode()).hexdigest()
    log.info(f"Auth enabled: {len(USERS)} user(s) configured")
else:
    log.info("Auth disabled (no CTFN_USERS env var set)")

AUTH_ENABLED = len(USERS) > 0

TOKEN_EXPIRY_HOURS = int(os.environ.get("CTFN_TOKEN_HOURS", 24))

# Active sessions: token â {"user": str, "expires": datetime}
SESSIONS: dict = {}


def _hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


def _create_token(username: str) -> str:
    token = secrets.token_hex(32)
    SESSIONS[token] = {
        "user": username,
        "expires": datetime.datetime.now() + datetime.timedelta(hours=TOKEN_EXPIRY_HOURS),
    }
    now = datetime.datetime.now()
    expired = [t for t, s in SESSIONS.items() if s["expires"] < now]
    for t in expired:
        del SESSIONS[t]
    return token


def _validate_token(token: str) -> Optional[str]:
    if not token or token not in SESSIONS:
        return None
    session = SESSIONS[token]
    if session["expires"] < datetime.datetime.now():
        del SESSIONS[token]
        return None
    return session["user"]


def _check_auth(authorization: str) -> Optional[JSONResponse]:
    """Check auth header. Returns error response if invalid, None if OK."""
    if not AUTH_ENABLED:
        return None
    token = authorization.replace("Bearer ", "").strip()
    user = _validate_token(token)
    if user is None:
        return JSONResponse(
            {"status": "error", "value": "#ERR: Not authenticated. Use /login first."},
            status_code=401,
        )
    return None


# ---------------------------------------------------------------------------
# Pressrisk WebSocket API
# ---------------------------------------------------------------------------
WS_URL = "wss://api.pressrisk.io/"
WS_USERNAME = os.environ.get("CTFN_WS_USERNAME", "")
WS_PASSWORD = os.environ.get("CTFN_WS_PASSWORD", "")
WS_HEADERS = {
    "Username": WS_USERNAME,
    "Password": WS_PASSWORD,
}

WS_PAGE_TIMEOUT = 12
WS_PAGE_DELAY = 3

# ---------------------------------------------------------------------------
# In-memory cache (key -> (timestamp, data))
# ---------------------------------------------------------------------------
_cache: dict[str, tuple[float, Any]] = {}
CACHE_TTL = 600  # 10 minutes

# ---------------------------------------------------------------------------
# WebSocket fetcher
# ---------------------------------------------------------------------------
try:
    import websocket
    WS_AVAILABLE = True
except ImportError:
    WS_AVAILABLE = False
    log.warning("websocket-client not installed â WebSocket fetching disabled")


def _ws_fetch_page(query_type: str, page: int = 1, timeout: int = WS_PAGE_TIMEOUT) -> dict:
    """Fetch a single page from the Pressrisk WebSocket API."""
    if not WS_AVAILABLE:
        return {"error": "websocket-client not installed"}

    envelopes = []
    connected = threading.Event()
    error_msg = [None]

    def on_open(ws):
        connected.set()
        query = {"query": {"type": query_type, "page": page}}
        ws.send(json.dumps(query))

    def on_message(ws, message):
        try:
            data = json.loads(message)
            if data.get("type") == "Keep-Alive":
                return
            envelopes.append(data)
        except json.JSONDecodeError:
            envelopes.append({"_raw": message})

    def on_error(ws, error):
        error_msg[0] = str(error)

    def on_close(ws, code, msg):
        pass

    ws = websocket.WebSocketApp(
        WS_URL,
        header=[f"{k}: {v}" for k, v in WS_HEADERS.items()],
        on_open=on_open,
        on_message=on_message,
        on_error=on_error,
        on_close=on_close,
    )

    ws_thread = threading.Thread(
        target=lambda: ws.run_forever(ping_interval=10), daemon=True
    )
    ws_thread.start()

    if not connected.wait(timeout=10):
        if error_msg[0]:
            return {"error": f"Connection failed: {error_msg[0]}"}
        return {"error": "Connection timed out"}

    time.sleep(timeout)
    ws.close()
    ws_thread.join(timeout=5)

    if error_msg[0] and not envelopes:
        return {"error": error_msg[0]}
    if not envelopes:
        return {"error": "No response received"}

    return envelopes[0]


def _ws_fetch_pages(query_type: str, max_pages: int | None = None,
                    timeout: int = WS_PAGE_TIMEOUT) -> tuple[list, dict]:
    """Fetch multiple pages from the Pressrisk WebSocket API."""
    env = _ws_fetch_page(query_type, page=1, timeout=timeout)
    if "error" in env:
        return [], env

    items = env.get("items", [])
    pagination = env.get("pagination", {})
    total_pages = pagination.get("pages", 1)

    pages_to_fetch = min(max_pages, total_pages) if max_pages else total_pages

    for p in range(2, pages_to_fetch + 1):
        time.sleep(WS_PAGE_DELAY)
        env = _ws_fetch_page(query_type, page=p, timeout=timeout)
        if "error" not in env:
            items.extend(env.get("items", []))

    return items, pagination


def _grep_filter(items: list, pattern: str) -> list:
    """Client-side case-insensitive filter across all string fields."""
    pat = re.compile(re.escape(pattern), re.IGNORECASE)

    def _search(obj):
        if isinstance(obj, str):
            return pat.search(obj) is not None
        elif isinstance(obj, dict):
            return any(_search(v) for v in obj.values())
        elif isinstance(obj, list):
            return any(_search(v) for v in obj)
        return False

    return [item for item in items if _search(item)]


# ---------------------------------------------------------------------------
# High-level data functions
# ---------------------------------------------------------------------------
def _search_news(search: str = "", pages: str = "1", limit: int = 20) -> dict:
    now = time.time()
    cache_key = f"news:{search}:{pages}:{limit}"
    if cache_key in _cache and now - _cache[cache_key][0] < CACHE_TTL:
        return _cache[cache_key][1]

    max_pages = None if pages == "all" else int(pages)
    items, pagination = _ws_fetch_pages("news", max_pages=max_pages)

    if isinstance(pagination, dict) and "error" in pagination:
        return {"meta": {"error": pagination["error"]}, "items": []}

    if search:
        items = _grep_filter(items, search)

    result = {
        "meta": {
            "query_type": "news",
            "api_total": pagination.get("total", 0) if isinstance(pagination, dict) else 0,
            "filter": search,
            "returned": min(limit, len(items)),
        },
        "items": items[:limit]
    }
    _cache[cache_key] = (now, result)
    return result


def _get_situations(search: str = "", pages: str = "all", limit: int = 20) -> dict:
    now = time.time()
    cache_key = f"situation:{search}:{pages}:{limit}"
    if cache_key in _cache and now - _cache[cache_key][0] < CACHE_TTL:
        return _cache[cache_key][1]

    max_pages = None if pages == "all" else int(pages)
    items, pagination = _ws_fetch_pages("situation", max_pages=max_pages)

    if isinstance(pagination, dict) and "error" in pagination:
        return {"meta": {"error": pagination["error"]}, "items": []}

    if search:
        items = _grep_filter(items, search)

    result = {
        "meta": {
            "query_type": "situation",
            "api_total": pagination.get("total", 0) if isinstance(pagination, dict) else 0,
            "filter": search,
            "returned": min(limit, len(items)),
        },
        "items": items[:limit]
    }
    _cache[cache_key] = (now, result)
    return result


def _get_dma(search: str = "", pages: str = "1", limit: int = 20) -> dict:
    now = time.time()
    cache_key = f"dma:{search}:{pages}:{limit}"
    if cache_key in _cache and now - _cache[cache_key][0] < CACHE_TTL:
        return _cache[cache_key][1]

    max_pages = None if pages == "all" else int(pages)
    items, pagination = _ws_fetch_pages("dma", max_pages=max_pages)

    if isinstance(pagination, dict) and "error" in pagination:
        return {"meta": {"error": pagination["error"]}, "items": []}

    if search:
        items = _grep_filter(items, search)

    result = {
        "meta": {
            "query_type": "dma",
            "api_total": pagination.get("total", 0) if isinstance(pagination, dict) else 0,
            "filter": search,
            "returned": min(limit, len(items)),
        },
        "items": items[:limit]
    }
    _cache[cache_key] = (now, result)
    return result


# ---------------------------------------------------------------------------
# Parameter dispatcher (unchanged from original)
# ---------------------------------------------------------------------------
def _resolve(param: str, ticker: str, search: str, limit: int,
             pages: str, extra: str | None) -> Any:
    p = param.lower().strip().replace(" ", "_")
    query = ticker or search

    # === NEWS PARAMETERS ===
    if p in ("latest_news", "news_latest", "recent_news"):
        result = _search_news(search=query, pages="1", limit=limit or 10)
        items = result.get("items", [])
        if not items:
            return "No news found"
        lines = []
        for item in items:
            title = item.get("title", "")
            date = item.get("published", "")[:10]
            tickers = ", ".join(item.get("tickers", []))
            lines.append(f"[{date}] {title} ({tickers})")
        return " | ".join(lines)

    if p in ("news_count", "story_count"):
        result = _search_news(search=query, pages="all", limit=1)
        return result.get("meta", {}).get("returned", 0)

    if p in ("news_headline", "latest_headline", "headline"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        return items[0].get("title", "N/A") if items else "No news found"

    if p in ("news_headlines", "headlines"):
        result = _search_news(search=query, pages="1", limit=limit or 5)
        items = result.get("items", [])
        if not items:
            return "No news found"
        return " | ".join(item.get("title", "") for item in items)

    if p in ("news_date", "latest_news_date"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        return items[0].get("published", "N/A")[:10] if items else "N/A"

    if p in ("news_author", "latest_author"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        return items[0].get("author", "N/A") if items else "N/A"

    if p in ("news_url", "latest_news_url", "article_url"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        return items[0].get("url", "N/A") if items else "N/A"

    if p in ("news_tags", "latest_tags"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        if items:
            tags = items[0].get("tags", [])
            return ", ".join(tags) if isinstance(tags, list) else str(tags)
        return "N/A"

    if p in ("news_tickers", "latest_tickers", "related_tickers"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        if items:
            tickers = items[0].get("tickers", [])
            return ", ".join(tickers) if isinstance(tickers, list) else str(tickers)
        return "N/A"

    if p in ("news_views", "latest_views", "view_count"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        return items[0].get("viewsCount", 0) if items else 0

    if p in ("news_slug", "latest_slug"):
        result = _search_news(search=query, pages="1", limit=1)
        items = result.get("items", [])
        return items[0].get("slug", "N/A") if items else "N/A"

    if p in ("news_search", "search_news"):
        result = _search_news(search=query, pages=pages or "all", limit=limit or 10)
        items = result.get("items", [])
        if not items:
            return f"No CTFN news matching '{query}'"
        lines = []
        for item in items:
            title = item.get("title", "")
            date = item.get("published", "")[:10]
            lines.append(f"[{date}] {title}")
        return " | ".join(lines)

    # === Nth NEWS ITEM (news_1, news_2, ...) ===
    if p.startswith("news_") and p[5:].isdigit():
        n = int(p[5:])
        result = _search_news(search=query, pages="1", limit=n)
        items = result.get("items", [])
        if n <= len(items):
            item = items[n - 1]
            title = item.get("title", "")
            date = item.get("published", "")[:10]
            tickers = ", ".join(item.get("tickers", []))
            return f"[{date}] {title} ({tickers})"
        return f"No news item #{n} found"

    # === SITUATION PARAMETERS ===
    if p in ("situations", "active_situations", "situation_list"):
        result = _get_situations(search=query, limit=limit or 10)
        items = result.get("items", [])
        if not items:
            return "No situations found" if query else "No active situations"
        lines = []
        for item in items:
            title = item.get("title", "")
            tickers = ", ".join(item.get("tickers", []))
            lines.append(f"{title} ({tickers})" if tickers else title)
        return " | ".join(lines)

    if p in ("situation_count", "num_situations"):
        result = _get_situations(search=query, limit=1)
        meta = result.get("meta", {})
        if query:
            result2 = _get_situations(search=query, limit=9999)
            return len(result2.get("items", []))
        return meta.get("api_total", 0)

    if p in ("situation", "latest_situation"):
        result = _get_situations(search=query, limit=1)
        items = result.get("items", [])
        if items:
            item = items[0]
            title = item.get("title", "")
            tickers = ", ".join(item.get("tickers", []))
            return f"{title} ({tickers})" if tickers else title
        return "No situation found"

    if p in ("situation_title",):
        result = _get_situations(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("title", "N/A") if items else "N/A"

    if p in ("situation_tickers",):
        result = _get_situations(search=query, limit=1)
        items = result.get("items", [])
        if items:
            tickers = items[0].get("tickers", [])
            return ", ".join(tickers) if isinstance(tickers, list) else str(tickers)
        return "N/A"

    if p in ("situation_date", "situation_published"):
        result = _get_situations(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("published", "N/A")[:10] if items else "N/A"

    if p in ("situation_url", "situation_link"):
        result = _get_situations(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("url", "N/A") if items else "N/A"

    if p in ("situation_slug",):
        result = _get_situations(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("slug", "N/A") if items else "N/A"

    if p in ("situation_search", "search_situations"):
        result = _get_situations(search=query, limit=limit or 10)
        items = result.get("items", [])
        if not items:
            return f"No situations matching '{query}'"
        lines = []
        for item in items:
            title = item.get("title", "")
            tickers = ", ".join(item.get("tickers", []))
            lines.append(f"{title} ({tickers})" if tickers else title)
        return " | ".join(lines)

    # === Nth SITUATION (situation_1, situation_2, ...) ===
    if p.startswith("situation_") and p[10:].isdigit():
        n = int(p[10:])
        result = _get_situations(search=query, limit=n)
        items = result.get("items", [])
        if n <= len(items):
            item = items[n - 1]
            title = item.get("title", "")
            tickers = ", ".join(item.get("tickers", []))
            return f"{title} ({tickers})" if tickers else title
        return f"No situation #{n} found"

    # === DMA PARAMETERS ===
    if p in ("dma", "dma_latest", "latest_dma"):
        result = _get_dma(search=query, limit=limit or 5)
        items = result.get("items", [])
        if not items:
            return "No DMA filings found"
        lines = []
        for item in items:
            company = item.get("companyNameLong", "")
            t = item.get("ticker", "")
            date = item.get("filingDate", "")
            lines.append(f"{company} [{t}] {date}" if t else f"{company} {date}")
        return " | ".join(lines)

    if p in ("dma_count",):
        result = _get_dma(search=query, limit=1)
        return result.get("meta", {}).get("api_total", 0)

    if p in ("dma_company", "dma_filer"):
        result = _get_dma(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("companyNameLong", "N/A") if items else "N/A"

    if p in ("dma_ticker",):
        result = _get_dma(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("ticker", "N/A") if items else "N/A"

    if p in ("dma_filing_date", "dma_date"):
        result = _get_dma(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("filingDate", "N/A") if items else "N/A"

    if p in ("dma_cfius",):
        result = _get_dma(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("CFIUS", "N/A") if items else "N/A"

    if p in ("dma_enforceability", "dma_rating"):
        result = _get_dma(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("enforceabilityRating", "N/A") if items else "N/A"

    if p in ("dma_filing_url", "dma_url"):
        result = _get_dma(search=query, limit=1)
        items = result.get("items", [])
        return items[0].get("linkToFilingDetails", "N/A") if items else "N/A"

    if p in ("dma_search", "search_dma"):
        result = _get_dma(search=query, pages=pages or "1", limit=limit or 10)
        items = result.get("items", [])
        if not items:
            return f"No DMA filings matching '{query}'"
        lines = []
        for item in items:
            company = item.get("companyNameLong", "")
            t = item.get("ticker", "")
            date = item.get("filingDate", "")
            lines.append(f"{company} [{t}] {date}" if t else f"{company} {date}")
        return " | ".join(lines)

    # === Nth DMA (dma_1, dma_2, ...) ===
    if p.startswith("dma_") and p[4:].isdigit():
        n = int(p[4:])
        result = _get_dma(search=query, limit=n)
        items = result.get("items", [])
        if n <= len(items):
            item = items[n - 1]
            company = item.get("companyNameLong", "")
            t = item.get("ticker", "")
            rating = item.get("enforceabilityRating", "")
            return f"{company} [{t}] {rating}" if t else f"{company} {rating}"
        return f"No DMA filing #{n} found"

    # === META / STATS ===
    if p in ("total_news", "news_total"):
        result = _search_news(search="", pages="1", limit=1)
        return result.get("meta", {}).get("api_total", 0)

    if p in ("total_situations", "situations_total"):
        result = _get_situations(search="", limit=1)
        return result.get("meta", {}).get("api_total", 0)

    if p in ("total_dma", "dma_total"):
        result = _get_dma(search="", limit=1)
        return result.get("meta", {}).get("api_total", 0)

    if p in ("has_coverage", "covered"):
        result = _search_news(search=query, pages="all", limit=1)
        has_news = len(result.get("items", [])) > 0
        result2 = _get_situations(search=query, limit=1)
        has_sit = len(result2.get("items", [])) > 0
        if has_news and has_sit:
            return "News + Situation"
        elif has_news:
            return "News only"
        elif has_sit:
            return "Situation only"
        return "No coverage"

    return f"#ERR: Unknown parameter '{param}'"


# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------
app = FastAPI(title="CTFN Excel Server", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Authentication endpoints
# ---------------------------------------------------------------------------
@app.post("/login")
async def login(request: Request):
    """Authenticate with username/password, receive a session token."""
    if not AUTH_ENABLED:
        return JSONResponse({"status": "ok", "token": "AUTH_DISABLED",
                             "message": "Authentication is not enabled on this server"})

    try:
        body = await request.json()
    except Exception:
        return JSONResponse({"status": "error", "message": "Invalid JSON body"}, status_code=400)

    username = str(body.get("username", "")).strip().lower()
    password = str(body.get("password", ""))

    if not username or not password:
        return JSONResponse({"status": "error", "message": "Username and password required"}, status_code=400)

    pw_hash = _hash_pw(password)
    if username not in USERS or USERS[username] != pw_hash:
        log.warning(f"Failed login attempt for user '{username}'")
        return JSONResponse({"status": "error", "message": "Invalid credentials"}, status_code=401)

    token = _create_token(username)
    log.info(f"User '{username}' logged in (token expires in {TOKEN_EXPIRY_HOURS}h)")
    return JSONResponse({
        "status": "ok",
        "token": token,
        "user": username,
        "expires_in_hours": TOKEN_EXPIRY_HOURS,
    })


@app.post("/logout")
async def logout(authorization: str = Header("", alias="Authorization")):
    """Invalidate a session token."""
    token = authorization.replace("Bearer ", "").strip()
    if token in SESSIONS:
        user = SESSIONS[token]["user"]
        del SESSIONS[token]
        log.info(f"User '{user}' logged out")
    return JSONResponse({"status": "ok"})


# ---------------------------------------------------------------------------
# API endpoint
# ---------------------------------------------------------------------------
@app.get("/api/ctfn")
async def ctfn_api(
    param: str = Query(..., description="Data parameter to retrieve"),
    ticker: str = Query("", description="Ticker symbol to search for"),
    search: str = Query("", description="Search text (company name, keyword)"),
    limit: int = Query(10, description="Max items to return"),
    pages: str = Query("1", description="Pages to fetch: number or 'all'"),
    extra: Optional[str] = Query(None, description="Extra parameter"),
    authorization: str = Header("", alias="Authorization"),
):
    auth_err = _check_auth(authorization)
    if auth_err:
        return auth_err

    try:
        result = _resolve(param, ticker, search, limit, pages, extra)
        return JSONResponse({"status": "ok", "value": result})
    except Exception as e:
        log.exception(f"Error resolving {param}")
        return JSONResponse({"status": "error", "value": f"#ERR: {str(e)}"}, status_code=200)


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "server": "ctfn-excel-server",
        "version": "2.0.0",
        "auth_enabled": AUTH_ENABLED,
        "users": len(USERS),
        "active_sessions": len(SESSIONS),
        "cache_entries": len(_cache),
        "ws_configured": bool(WS_USERNAME),
    }


@app.get("/api/params")
async def list_params():
    """Return all supported parameters."""
    return {
        "News": [
            "latest_news", "news_headline", "news_headlines",
            "news_date", "news_author", "news_url", "news_tags",
            "news_tickers", "news_views", "news_slug", "news_count",
            "news_search", "news_1 / news_2 / news_N"
        ],
        "Active Situations": [
            "situations", "situation", "situation_title",
            "situation_tickers", "situation_date", "situation_url",
            "situation_slug", "situation_count", "situation_search",
            "situation_1 / situation_2 / situation_N"
        ],
        "DMA (Definitive Merger Agreements)": [
            "dma", "dma_company", "dma_ticker", "dma_filing_date",
            "dma_cfius", "dma_enforceability", "dma_filing_url",
            "dma_count", "dma_search", "dma_1 / dma_2 / dma_N"
        ],
        "Meta / Coverage": [
            "total_news", "total_situations", "total_dma", "has_coverage"
        ],
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    log.info(f"Starting CTFN Excel Server on {host}:{PORT}...")
    uvicorn.run(app, host=host, port=PORT, log_level="info")

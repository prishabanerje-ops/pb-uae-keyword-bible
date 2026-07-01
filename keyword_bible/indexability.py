import re
import time
import urllib.request
import urllib.robotparser
from urllib.parse import urlparse
from . import config

_META_ROBOTS = re.compile(r'<meta[^>]+name=["\']robots["\'][^>]+content=["\']([^"\']+)["\']', re.I)
_CANON = re.compile(r'<link[^>]+rel=["\']canonical["\'][^>]+href=["\']([^"\']+)["\']', re.I)

def parse_signals(html, headers, final_url, requested_url):
    meta = _META_ROBOTS.search(html or "")
    noindex = bool(meta and "noindex" in meta.group(1).lower())
    canon_m = _CANON.search(html or "")
    canonical = canon_m.group(1).strip() if canon_m else ""
    xr = headers.get("X-Robots-Tag", "") if headers else ""
    return {
        "noindex": noindex,
        "canonical": canonical,
        "self_canonical": (canonical.rstrip("/") == (requested_url or "").rstrip("/")) if canonical else False,
        "x_robots_noindex": "noindex" in str(xr).lower(),
    }

def indexable_verdict(status, robots_allowed, signals):
    if status != 200:
        return (False, f"http {status}")
    if not robots_allowed:
        return (False, "blocked by robots.txt")
    if signals.get("noindex"):
        return (False, "meta robots noindex")
    if signals.get("x_robots_noindex"):
        return (False, "x-robots-tag noindex")
    return (True, "indexable")

_ROBOTS_CACHE = {}  # netloc -> RobotFileParser, or None when robots.txt is unreachable

def _robots_allowed(url):
    parts = urlparse(url)
    netloc = parts.netloc
    if netloc not in _ROBOTS_CACHE:
        rp = urllib.robotparser.RobotFileParser()
        rp.set_url(f"{parts.scheme}://{netloc}/robots.txt")
        try:
            rp.read()
        except Exception:
            rp = None  # if robots.txt unreachable, treat as allowed (documented)
        _ROBOTS_CACHE[netloc] = rp
    rp = _ROBOTS_CACHE[netloc]
    if rp is None:
        return True
    return rp.can_fetch(config.USER_AGENT, url)

def check_url(url):
    req = urllib.request.Request(url, headers={"User-Agent": config.USER_AGENT})
    status, final_url, html, headers = 0, url, "", {}
    try:
        with urllib.request.urlopen(req, timeout=config.HTTP_TIMEOUT) as resp:
            status = resp.status
            final_url = resp.geturl()
            headers = dict(resp.headers)
            html = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        status = e.code
        headers = dict(e.headers or {})
    except Exception as e:
        return {"http_status": "ERROR", "final_url": url, "robots_allowed": "",
                "noindex": "", "canonical": "", "self_canonical": "",
                "indexable": False, "reason": f"fetch error: {e}"}
    allowed = _robots_allowed(url)
    sig = parse_signals(html, headers, final_url, url)
    ok, reason = indexable_verdict(status, allowed, sig)
    time.sleep(0.3)
    return {"http_status": status, "final_url": final_url, "robots_allowed": allowed,
            "noindex": sig["noindex"], "canonical": sig["canonical"],
            "self_canonical": sig["self_canonical"], "indexable": ok, "reason": reason}

def gsc_status(url):
    # GSC URL Inspection adapter. Wired at execution time if credentials exist.
    # Until then, honest fallback (never fabricated).
    return {"coverage_state": "GSC not connected", "last_crawl": "GSC not connected",
            "google_canonical": "GSC not connected"}

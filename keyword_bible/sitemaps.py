import re
import urllib.request
from . import config

_LOC = re.compile(r"<loc>\s*([^<]+?)\s*</loc>", re.IGNORECASE)

def parse(xml_text):
    """Extract <loc> values from sitemap XML."""
    return [m.group(1) for m in _LOC.finditer(xml_text)]

def fetch(url):
    """Fetch sitemap XML from URL. Raises RuntimeError on non-200 status."""
    req = urllib.request.Request(url, headers={"User-Agent": config.USER_AGENT})
    with urllib.request.urlopen(req, timeout=config.HTTP_TIMEOUT) as resp:
        if resp.status != 200:
            raise RuntimeError(f"Sitemap fetch failed {resp.status} for {url}")
        return resp.read().decode("utf-8", errors="replace")

def load_all():
    """Load and parse all sitemaps (health, travel) from config.SITEMAPS."""
    return {vert: parse(fetch(url)) for vert, url in config.SITEMAPS.items()}

"""Pull power sentences out of a shareholder letter or 20-F excerpt.

The live fetcher stores raw text under data/cache and is optional.
Tests run on fixtures so a filing URL does not have to be up.
"""

from __future__ import annotations

import re
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CACHE = ROOT / "data" / "cache"

POWER_TERMS = ("contracted power", "connected power", "active power")
USER_AGENT = "NBIS-Scrapes/0.1 (personal research; aaditprayag@users.noreply.github.com)"


def power_sentences(text: str) -> list[str]:
    flat = " ".join(text.split())
    parts = re.split(r"(?<=[.!?])\s+", flat)
    hits = []
    for sentence in parts:
        low = sentence.lower()
        if any(term in low for term in POWER_TERMS):
            cleaned = sentence.strip()
            if cleaned and cleaned not in hits:
                hits.append(cleaned)
    return hits


def fetch_text(url: str, timeout: int = 45) -> str:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        raw = response.read()
    CACHE.mkdir(parents=True, exist_ok=True)
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", url)[-120:]
    (CACHE / name).write_bytes(raw)
    return raw.decode("utf-8", errors="replace")

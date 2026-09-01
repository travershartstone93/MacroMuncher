#!/usr/bin/env python3
"""
Offline-first Microsoft Learn docs harvester for:
  - Excel VBA API reference (learn.microsoft.com/en-us/office/vba/api/...)
  - Power Query M reference (learn.microsoft.com/en-us/powerquery-m/...)

Hardened fixes applied (from brittleness review):
- ✅ Case-insensitive domain allow-list (Learn.Microsoft.Com works)
- ✅ Content-Type validation (only cache/parse HTML; reject unexpected 200 bodies)
- ✅ Charset-aware decoding (uses HTTP headers when available)
- ✅ Uses API-page heuristic for crawl prioritization AND controlled pruning
- ✅ Deterministic crawl ordering (secondary sort key is URL, not insertion order)
- ✅ Scrapes richer fields: signature, param_docs (tables/dl), return_value, remarks
- ✅ Header search includes h2/h3/h4 (more resilient to doc heading changes)
- ✅ Adds cache schema_version + scraper_version + validation metrics
- ✅ Validation step warns/fails when extraction quality collapses (doc layout change detector)
- ✅ Manifest includes mapping URL -> cached filename for debugging
- ✅ URL canonicalization normalizes query parameter ordering and strips fragments

Dependencies:
  python3 -m pip install requests beautifulsoup4 lxml

Usage:
  python3 harvest_docs.py harvest --out-dir doc_cache
  python3 harvest_docs.py ensure  --out-dir doc_cache
  python3 harvest_docs.py validate --out-dir doc_cache

Notes:
- This crawler is scope-bounded by domain + path prefixes.
- Respectful crawling: throttle + caching. (Robots/ToS considerations are up to you.)
"""

from __future__ import annotations

import argparse
import dataclasses
import hashlib
import json
import os
import re
import sys
import time
import urllib.parse
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Set, Tuple

import requests
from bs4 import BeautifulSoup

# -----------------------------
# Versions / schema
# -----------------------------

SCHEMA_VERSION = 2
SCRAPER_VERSION = 3  # bump when extraction logic changes materially

# -----------------------------
# Config
# -----------------------------

DEFAULT_SEEDS = [
    # Excel VBA reference landing pages / indexes
    "https://learn.microsoft.com/en-us/office/vba/api/overview/excel",
    "https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model",

    # Power Query M reference indexes
    "https://learn.microsoft.com/en-us/powerquery-m/power-query-m-function-reference",
    "https://learn.microsoft.com/en-us/powerquery-m/table-functions",
    "https://learn.microsoft.com/en-us/powerquery-m/",
]

TOC_SEEDS = [
    "https://learn.microsoft.com/en-us/office/vba/api/toc.json",
    "https://learn.microsoft.com/en-us/powerquery-m/toc.json",
]
ALLOWED_DOMAINS = {"learn.microsoft.com"}  # checked lowercased
ALLOWED_PATH_PREFIXES = (
    "/en-us/office/vba/api/",
    "/en-us/powerquery-m/",
)

LIKELY_API_PATH_PATTERNS = (
    re.compile(r"^/en-us/office/vba/api/(excel|office)\.[a-z0-9_.-]+$", re.I),
    re.compile(r"^/en-us/powerquery-m/[a-z0-9_.-]+$", re.I),
)

# Overview pages are useful “bridges” even when pruning
OVERVIEW_PATH_HINTS = (
    "/overview/",
    "function-reference",
    "table-functions",
    "list-functions",
    "record-functions",
    "text-functions",
    "number-functions",
    "date-functions",
    "duration-functions",
    "binary-functions",
)

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
REQUEST_TIMEOUT_S = 25


# -----------------------------
# Data model for cache
# -----------------------------

@dataclass
class DocRecord:
    id: str
    language: str                   # "vba_excel" or "powerquery_m"
    symbol: str                     # "Excel.Range.Sort" or "Table.SelectRows"
    kind: str                       # "method" | "function" | "object" | "unknown"
    summary: str
    signature: str
    parameters: List[str]
    param_docs: Dict[str, str]
    return_value: str
    remarks: str
    source_url: str
    retrieved_at_utc: str
    source_hash: str
    parse_warnings: List[str]

    def to_dict(self) -> Dict[str, Any]:
        return dataclasses.asdict(self)


# -----------------------------
# Utilities
# -----------------------------

def utc_now_iso() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())

def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def url_to_cache_name(url: str) -> str:
    return sha256_bytes(url.encode("utf-8")) + ".html"

def normalize_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def is_allowed_url(url: str) -> bool:
    try:
        u = urllib.parse.urlparse(url)
    except Exception:
        return False
    if u.scheme not in ("http", "https"):
        return False
    if u.netloc.lower() not in ALLOWED_DOMAINS:
        return False
    if not any(u.path.startswith(pfx) for pfx in ALLOWED_PATH_PREFIXES):
        return False
    return True

def canonicalize_url(url: str) -> str:
    """
    - strips fragments
    - normalizes query parameter ordering
    """
    u = urllib.parse.urlparse(url)
    u = u._replace(fragment="")
    if u.query:
        q = urllib.parse.parse_qsl(u.query, keep_blank_values=True)
        q.sort(key=lambda kv: (kv[0], kv[1]))
        u = u._replace(query=urllib.parse.urlencode(q))
    return u.geturl()

def looks_like_api_page(url: str) -> bool:
    u = urllib.parse.urlparse(url)
    return any(p.match(u.path) for p in LIKELY_API_PATH_PATTERNS)

def looks_like_overview_page(url: str) -> bool:
    path = urllib.parse.urlparse(url).path.lower()
    return any(hint in path for hint in OVERVIEW_PATH_HINTS)

def extract_links(base_url: str, html: str) -> List[str]:
    soup = BeautifulSoup(html, "lxml")
    links: List[str] = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href:
            continue
        abs_url = urllib.parse.urljoin(base_url, href)
        abs_url = canonicalize_url(abs_url)
        if is_allowed_url(abs_url):
            links.append(abs_url)
    return links

def extract_links_from_toc_json(base_url: str, toc_json: Any) -> List[str]:
    """
    Extract URLs from Learn TOC JSON structures.
    """
    out: List[str] = []

    def visit(node: Any):
        if isinstance(node, dict):
            href = node.get("href") or node.get("url")
            if isinstance(href, str) and href:
                abs_url = urllib.parse.urljoin(base_url, href)
                abs_url = canonicalize_url(abs_url)
                if is_allowed_url(abs_url):
                    out.append(abs_url)
            for key in ("items", "children"):
                if isinstance(node.get(key), list):
                    for child in node[key]:
                        visit(child)
        elif isinstance(node, list):
            for child in node:
                visit(child)

    visit(toc_json)
    return out

def decode_html_bytes(content: bytes, content_type_header: str) -> str:
    """
    Decode bytes using charset from Content-Type if present; otherwise UTF-8 fallback.
    """
    charset = None
    m = re.search(r"charset\s*=\s*([A-Za-z0-9_\-]+)", content_type_header or "", re.I)
    if m:
        charset = m.group(1).strip()
    if charset:
        try:
            return content.decode(charset, errors="replace")
        except Exception:
            pass
    return content.decode("utf-8", errors="replace")

def content_type_is_html(content_type_header: str) -> bool:
    ct = (content_type_header or "").lower()
    return ("text/html" in ct) or ("application/xhtml+xml" in ct)


# -----------------------------
# Crawler (deterministic priority queue + pruning)
# -----------------------------

class Harvester:
    def __init__(
        self,
        out_dir: str,
        throttle_s: float = 1.0,
        max_pages: int = 8000,
        max_depth: int = 7,
        verbose: bool = True,
        debug: bool = False,
    ):
        self.out_dir = out_dir
        self.throttle_s = throttle_s
        self.max_pages = max_pages
        self.max_depth = max_depth
        self.verbose = verbose
        self.debug = debug

        self.raw_dir = os.path.join(out_dir, "raw")
        os.makedirs(self.raw_dir, exist_ok=True)

        self.session = requests.Session()
        self.session.headers.update({"User-Agent": USER_AGENT})

        # url -> cache filename (debuggability)
        self.url_cache_map: Dict[str, str] = {}

    def fetch(self, url: str) -> Tuple[Optional[bytes], Optional[str], str]:
        """
        Returns (content_bytes, error, content_type)
        """
        cache_name = url_to_cache_name(url)
        cache_path = os.path.join(self.raw_dir, cache_name)
        self.url_cache_map[url] = cache_name

        if os.path.exists(cache_path):
            try:
                with open(cache_path, "rb") as f:
                    # cached files are assumed to be HTML we previously accepted
                    if self.debug:
                        print(f"[CACHE] {url}")
                    return f.read(), None, "text/html (cached)"
            except Exception as e:
                return None, f"Failed to read cache: {e}", ""

        time.sleep(self.throttle_s)

        if self.debug:
            print(f"[FETCH] {url}")
        try:
            resp = self.session.get(url, timeout=REQUEST_TIMEOUT_S)
        except Exception as e:
            return None, f"Request failed: {e}", ""

        if resp.status_code != 200:
            if self.debug:
                print(f"[RESP] {url} -> {resp.status_code}")
            return None, f"HTTP {resp.status_code}", resp.headers.get("Content-Type", "")

        ct = resp.headers.get("Content-Type", "")
        if self.debug:
            print(f"[RESP] {url} -> 200 {ct}")
        if not content_type_is_html(ct):
            # Don't cache non-HTML successful responses
            return None, f"Unexpected Content-Type (not HTML): {ct}", ct

        content = resp.content
        # basic sanity: ensure it looks like HTML
        head = content[:500].lower()
        if b"<html" not in head and b"<!doctype html" not in head:
            # still maybe HTML, but likely an error payload
            # do not cache, report warning
            return None, "Body did not look like HTML (missing <html/doctype in first 500 bytes)", ct

        try:
            with open(cache_path, "wb") as f:
                f.write(content)
        except Exception as e:
            return None, f"Failed to write cache: {e}", ct

        return content, None, ct

    def fetch_json(self, url: str) -> Tuple[Optional[Any], Optional[str]]:
        """
        Fetch JSON and return (obj, error).
        """
        try:
            resp = self.session.get(url, timeout=REQUEST_TIMEOUT_S)
        except Exception as e:
            return None, f"Request failed: {e}"
        if resp.status_code != 200:
            return None, f"HTTP {resp.status_code}"
        try:
            return resp.json(), None
        except Exception as e:
            return None, f"JSON decode failed: {e}"

    def crawl(self, seeds: List[str]) -> Tuple[List[str], Dict[str, str]]:
        """
        Deterministic crawl ordering:
        priority tuple is:
          (is_not_api, depth, url)
        This makes traversal reproducible across runs given same seed set.
        """
        # Use a list as a heap: each item is (pri, depth, url)
        heap: List[Tuple[int, int, str]] = []
        seen: Set[str] = set()
        fetched: List[str] = []
        errors: Dict[str, str] = {}

        def push(url: str, depth: int):
            pri = 0 if looks_like_api_page(url) else 1
            heap.append((pri, depth, url))

        for s in seeds:
            s = canonicalize_url(s)
            if is_allowed_url(s):
                push(s, 0)

        # Attempt TOC-based discovery for richer coverage
        for toc_url in TOC_SEEDS:
            toc_obj, err = self.fetch_json(toc_url)
            if err:
                if self.verbose:
                    print(f"[WARN] TOC fetch failed: {toc_url} -> {err}")
                continue
            toc_links = extract_links_from_toc_json(toc_url, toc_obj)
            if self.verbose:
                print(f"[INFO] TOC links discovered: {len(toc_links)} from {toc_url}")
            for link in toc_links:
                push(link, 1)

        # heapify repeatedly in a deterministic way (small heap sizes are fine)
        while heap and len(fetched) < self.max_pages:
            heap.sort()  # deterministic
            pri, depth, url = heap.pop(0)

            if url in seen:
                continue
            seen.add(url)

            if depth > self.max_depth:
                continue

            content, err, ct = self.fetch(url)
            if err:
                errors[url] = err
                if self.verbose:
                    print(f"[ERR] {url} -> {err}")
                continue

            fetched.append(url)
            if self.verbose and len(fetched) % 50 == 0:
                print(f"[OK] fetched {len(fetched)} pages... (last: {url})")

            html = decode_html_bytes(content, ct)

            leaf_api = looks_like_api_page(url)

            # Controlled expansion:
            # - from API pages: follow other API pages, plus overview/bridge pages
            # - from non-API pages: follow all allowed
            for link in extract_links(url, html):
                if link in seen:
                    continue
                if depth + 1 > self.max_depth:
                    continue
                if leaf_api:
                    if looks_like_api_page(link) or looks_like_overview_page(link):
                        push(link, depth + 1)
                else:
                    push(link, depth + 1)

        return fetched, errors


# -----------------------------
# Extraction helpers (more resilient)
# -----------------------------

def find_heading(soup: BeautifulSoup, names: List[str]) -> Optional[Any]:
    names_lc = {n.lower() for n in names}
    for hx in soup.find_all(["h2", "h3", "h4"]):
        t = normalize_ws(hx.get_text()).lower()
        if t in names_lc:
            return hx
    return None

def next_meaningful_paragraph(node: Any, max_hops: int = 80) -> str:
    cur = node
    hops = 0
    while cur and hops < max_hops:
        cur = cur.find_next()
        hops += 1
        if not cur:
            break
        if cur.name == "p":
            txt = normalize_ws(cur.get_text())
            if txt:
                return txt
    return ""

def extract_codeblock_under_heading(soup: BeautifulSoup, heading_texts: List[str]) -> str:
    hx = find_heading(soup, heading_texts)
    if not hx:
        return ""
    node = hx
    for _ in range(120):
        node = node.find_next()
        if not node:
            break
        if node.name == "pre":
            return normalize_ws(node.get_text("\n"))
    return ""

def extract_paragraph_under_heading(soup: BeautifulSoup, heading_texts: List[str]) -> str:
    hx = find_heading(soup, heading_texts)
    if not hx:
        return ""
    node = hx
    for _ in range(120):
        node = node.find_next()
        if not node:
            break
        if node.name == "p":
            txt = normalize_ws(node.get_text(" "))
            if txt:
                return txt
        if node.name == "pre":
            # Prefer paragraph syntax; avoid grabbing example code.
            break
    return ""

def extract_table_after_heading(soup: BeautifulSoup, heading_names: List[str]) -> List[List[str]]:
    hx = find_heading(soup, heading_names)
    if not hx:
        return []
    node = hx
    for _ in range(160):
        node = node.find_next()
        if not node:
            break
        if node.name == "table":
            rows = []
            for tr in node.find_all("tr"):
                cells = [normalize_ws(td.get_text(" ")) for td in tr.find_all(["th", "td"])]
                if cells:
                    rows.append(cells)
            return rows
        if node.name == "dl":
            rows = []
            for dt in node.find_all("dt"):
                dd = dt.find_next_sibling("dd")
                name = normalize_ws(dt.get_text(" "))
                desc = normalize_ws(dd.get_text(" ")) if dd else ""
                if name:
                    rows.append([name, desc])
            return rows
    return []

def table_rows_to_param_docs(rows: List[List[str]]) -> Dict[str, str]:
    if not rows:
        return {}

    # Detect header row by column names
    headers = [c.lower() for c in rows[0]]
    header_like = any(h in ("name", "parameter", "argument") for h in headers)
    data_rows = rows[1:] if header_like else rows

    param_docs: Dict[str, str] = {}
    for r in data_rows:
        if not r:
            continue
        name = (r[0] or "").strip()
        desc = (r[1] if len(r) > 1 else "").strip()
        name_m = re.match(r"^([A-Za-z_]\w*)", name)
        if name_m:
            pname = name_m.group(1)
            if pname not in param_docs:
                param_docs[pname] = desc
            elif desc and not param_docs[pname]:
                param_docs[pname] = desc
    return param_docs

def _is_callout_paragraph(p: Any) -> bool:
    for parent in p.parents:
        classes = parent.get("class", []) if hasattr(parent, "get") else []
        classes_lc = {c.lower() for c in classes}
        if any(c in classes_lc for c in ("note", "tip", "important", "warning", "caution", "alert")):
            return True
    return False

def first_meaningful_paragraph(soup: BeautifulSoup) -> str:
    main = soup.find("main") or soup.find("article") or soup
    skip_phrases = (
        "this browser is no longer supported",
        "upgrade to microsoft edge",
        "access to this page requires authorization",
    )
    for p in main.find_all("p"):
        if _is_callout_paragraph(p):
            continue
        txt = normalize_ws(p.get_text())
        if not txt:
            continue
        low = txt.lower()
        # avoid boilerplate
        if low == "in this article":
            continue
        if low.startswith("read more"):
            continue
        if low == "note":
            continue
        if any(phrase in low for phrase in skip_phrases):
            continue
        return txt
    return ""

def parse_params_from_signature_best_effort(signature: str) -> List[str]:
    """
    Best-effort parameter name extraction from a signature string.
    Uses balanced parens rather than regex so nested parens don't truncate.
    """
    if not signature:
        return []

    # Find first '(' and matching ')'
    l = signature.find("(")
    if l == -1:
        return []
    r = find_matching_rparen(signature, l)
    if r is None:
        return []

    inside = signature[l + 1:r]
    parts = split_top_level_commas_text(inside)

    params: List[str] = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        part = re.sub(r"\b(Optional|ByVal|ByRef)\b", "", part, flags=re.I).strip()
        # VBA: "x As Long" => x
        m = re.match(r"^([A-Za-z_]\w*)", part)
        if m:
            params.append(m.group(1))
    return params

def split_top_level_commas_text(s: str) -> List[str]:
    """
    Split on commas at depth 0, respecting parentheses and strings (VBA-style doubled quotes).
    Used for signature parsing.
    """
    out: List[str] = []
    cur: List[str] = []
    depth = 0
    in_str = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '"':
            if in_str and i + 1 < len(s) and s[i + 1] == '"':
                cur.append('""')
                i += 2
                continue
            in_str = not in_str
            cur.append(ch)
            i += 1
            continue

        if not in_str:
            if ch in "([{":
                depth += 1
            elif ch in ")]}":
                depth = max(0, depth - 1)
            elif ch == "," and depth == 0:
                seg = "".join(cur).strip()
                if seg:
                    out.append(seg)
                cur = []
                i += 1
                continue

        cur.append(ch)
        i += 1

    tail = "".join(cur).strip()
    if tail:
        out.append(tail)
    return out

def find_matching_rparen(s: str, lpar_idx: int) -> Optional[int]:
    depth = 0
    in_str = False
    i = lpar_idx
    while i < len(s):
        ch = s[i]
        if ch == '"':
            if in_str and i + 1 < len(s) and s[i + 1] == '"':
                i += 2
                continue
            in_str = not in_str
            i += 1
            continue
        if not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return i
        i += 1
    return None


# -----------------------------
# Page parsers
# -----------------------------

def parse_doc_record(url: str, html_text: str, html_bytes: bytes) -> Optional[DocRecord]:
    soup = BeautifulSoup(html_text, "lxml")
    title = normalize_ws(soup.title.get_text()) if soup.title else ""
    h1 = soup.find("h1")
    h1_text = normalize_ws(h1.get_text()) if h1 else ""

    path = urllib.parse.urlparse(url).path

    if path.startswith("/en-us/powerquery-m/"):
        return parse_powerquery_m(url, soup, title=title, h1=h1_text, html_bytes=html_bytes)

    if path.startswith("/en-us/office/vba/api/"):
        return parse_office_vba_excel(url, soup, title=title, h1=h1_text, html_bytes=html_bytes)

    return None

def parse_powerquery_m(url: str, soup: BeautifulSoup, title: str, h1: str, html_bytes: bytes) -> Optional[DocRecord]:
    symbol = ""
    if h1:
        symbol = normalize_ws(h1.split("-")[0]).strip()
    else:
        symbol = normalize_ws(title.split("-")[0]).strip()

    # Keep function pages that look like M functions (usually namespaced like Table.SelectRows)
    if "." not in symbol:
        return None

    warnings: List[str] = []
    summary = first_meaningful_paragraph(soup)
    if not summary:
        warnings.append("No summary paragraph detected.")

    signature = extract_codeblock_under_heading(soup, ["Syntax", "Usage"])
    if not signature:
        warnings.append("No Syntax/Usage code block detected; signature may be missing.")

    # M pages often have "Arguments"/"Parameters"
    rows = extract_table_after_heading(soup, ["Arguments", "Parameters"])
    param_docs = table_rows_to_param_docs(rows)

    params = parse_params_from_signature_best_effort(signature)
    if not params and param_docs:
        params = list(param_docs.keys())

    # Return value / return type
    return_value = ""
    hx_returns = find_heading(soup, ["Returns", "Return value"])
    if hx_returns:
        return_value = next_meaningful_paragraph(hx_returns)
    if not return_value and signature:
        m = re.search(r"\bas\s+([A-Za-z_]\w*)\b", signature, re.I)
        if m:
            return_value = f"Returns: {m.group(1)}"
    if not return_value:
        warnings.append("No return value detected.")

    remarks = ""
    hx_remarks = find_heading(soup, ["Remarks"])
    if hx_remarks:
        remarks = next_meaningful_paragraph(hx_remarks)

    rec_id = f"m.{symbol.lower()}"
    return DocRecord(
        id=rec_id,
        language="powerquery_m",
        symbol=symbol,
        kind="function",
        summary=summary or "",
        signature=signature or "",
        parameters=params,
        param_docs=param_docs,
        return_value=return_value or "",
        remarks=remarks or "",
        source_url=url,
        retrieved_at_utc=utc_now_iso(),
        source_hash=sha256_bytes(html_bytes),
        parse_warnings=warnings,
    )

def parse_office_vba_excel(url: str, soup: BeautifulSoup, title: str, h1: str, html_bytes: bytes) -> Optional[DocRecord]:
    path = urllib.parse.urlparse(url).path

    # Keep only likely reference pages like .../excel.range.sort
    if not re.search(r"/en-us/office/vba/api/(excel|office)\.", path, re.I):
        return None

    warnings: List[str] = []

    raw = normalize_ws(h1) if h1 else ""
    kind = "unknown"
    symbol_core = ""

    if raw:
        raw2 = re.sub(r"\s*\(.*?\)\s*$", "", raw).strip()  # drop "(Excel)"
        if re.search(r"\bmethod\b", raw2, re.I):
            kind = "method"
            raw2 = re.sub(r"\bmethod\b", "", raw2, flags=re.I).strip()
        elif re.search(r"\bobject\b", raw2, re.I):
            kind = "object"
            raw2 = re.sub(r"\bobject\b", "", raw2, flags=re.I).strip()
        symbol_core = raw2

    if not symbol_core:
        # URL fallback without title-casing
        base = os.path.basename(path)  # e.g. excel.range.sort
        if base.lower().startswith("excel."):
            symbol_core = base[len("excel."):]
        elif base.lower().startswith("office."):
            symbol_core = base[len("office."):]

    symbol = f"Excel.{symbol_core}" if symbol_core else ""
    if not symbol:
        return None

    summary = first_meaningful_paragraph(soup)
    if not summary:
        warnings.append("No summary paragraph detected.")

    signature = extract_paragraph_under_heading(soup, ["Syntax"])
    if not signature:
        signature = extract_codeblock_under_heading(soup, ["Syntax"])
    if not signature:
        warnings.append("No Syntax code block detected; signature may be missing.")

    params = parse_params_from_signature_best_effort(signature)

    rows = extract_table_after_heading(soup, ["Parameters"])
    param_docs = table_rows_to_param_docs(rows)
    if not params and param_docs:
        params = list(param_docs.keys())

    return_value = ""
    hx_ret = find_heading(soup, ["Return value", "Returns"])
    if hx_ret:
        return_value = next_meaningful_paragraph(hx_ret)
    if not return_value and signature:
        m = re.search(r"\bAs\s+([A-Za-z_]\w*)\b", signature, re.I)
        if m:
            return_value = f"Returns: {m.group(1)}"
    if not return_value and kind == "method":
        warnings.append("No return value detected.")

    remarks = ""
    hx_remarks = find_heading(soup, ["Remarks"])
    if hx_remarks:
        remarks = next_meaningful_paragraph(hx_remarks)

    rec_id = f"vba.{symbol.lower()}"
    return DocRecord(
        id=rec_id,
        language="vba_excel",
        symbol=symbol,
        kind=kind,
        summary=summary or "",
        signature=signature or "",
        parameters=params,
        param_docs=param_docs,
        return_value=return_value or "",
        remarks=remarks or "",
        source_url=url,
        retrieved_at_utc=utc_now_iso(),
        source_hash=sha256_bytes(html_bytes),
        parse_warnings=warnings,
    )


# -----------------------------
# Cache builder + validation
# -----------------------------

def compute_quality_metrics(records: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    def pct(n: int, d: int) -> float:
        return 0.0 if d == 0 else round(100.0 * n / d, 2)

    total = len(records)
    with_summary = sum(1 for r in records.values() if (r.get("summary") or "").strip())
    with_signature = sum(1 for r in records.values() if (r.get("signature") or "").strip())
    with_params = sum(1 for r in records.values() if (r.get("parameters") or []))
    with_param_docs = sum(1 for r in records.values() if (r.get("param_docs") or {}))
    with_return = sum(1 for r in records.values() if (r.get("return_value") or "").strip())

    return {
        "total_records": total,
        "with_summary": with_summary,
        "with_signature": with_signature,
        "with_parameters": with_params,
        "with_param_docs": with_param_docs,
        "with_return_value": with_return,
        "pct_with_summary": pct(with_summary, total),
        "pct_with_signature": pct(with_signature, total),
        "pct_with_parameters": pct(with_params, total),
        "pct_with_param_docs": pct(with_param_docs, total),
        "pct_with_return_value": pct(with_return, total),
    }

def validate_quality(metrics: Dict[str, Any]) -> Tuple[bool, List[str]]:
    """
    Basic sanity thresholds. Tune as you learn your crawl.
    Goal: detect sudden doc layout changes.
    """
    total = metrics.get("total_records", 0)
    warnings: List[str] = []
    ok = True

    if total < 200:
        warnings.append("Record count is quite low; crawl may have been too shallow or blocked.")
        # not necessarily fatal

    # If these drop drastically, something is wrong
    if metrics.get("pct_with_summary", 0.0) < 70.0:
        ok = False
        warnings.append("Too many records missing summary (<70%). Doc layout may have changed or extraction broke.")
    if metrics.get("pct_with_signature", 0.0) < 40.0:
        warnings.append("Many records missing signature (<40%). This may be normal for some pages, but could indicate extraction issues.")
    if metrics.get("pct_with_parameters", 0.0) < 40.0:
        warnings.append("Many records missing parameter list (<40%). This may be normal for objects but watch for sudden drops.")

    return ok, warnings

def build_docs_cache(raw_dir: str, urls: List[str], url_cache_map: Dict[str, str], out_json: str, errors_json: str, report_json: str, verbose: bool = True) -> Tuple[int, int]:
    records: Dict[str, Dict[str, Any]] = {}
    parse_errors: Dict[str, str] = {}

    for idx, url in enumerate(urls, 1):
        cache_name = url_cache_map.get(url) or url_to_cache_name(url)
        cache_path = os.path.join(raw_dir, cache_name)
        if not os.path.exists(cache_path):
            parse_errors[url] = "Missing raw cache file."
            continue

        try:
            with open(cache_path, "rb") as f:
                html_bytes = f.read()
        except Exception as e:
            parse_errors[url] = f"Failed reading raw: {e}"
            continue

        # Cached files are HTML we previously accepted; decode as UTF-8 fallback
        html_text = html_bytes.decode("utf-8", errors="replace")

        try:
            rec = parse_doc_record(url, html_text, html_bytes)
        except Exception as e:
            parse_errors[url] = f"Parse exception: {e}"
            continue

        if rec is None:
            continue

        # Deduplicate by id; keep richer record
        if rec.id in records:
            existing = records[rec.id]

            def richness(r: Dict[str, Any]) -> Tuple[int, int, int, int]:
                return (
                    1 if (r.get("summary") or "").strip() else 0,
                    1 if (r.get("signature") or "").strip() else 0,
                    len(r.get("parameters", []) or []),
                    len(r.get("param_docs", {}) or {}),
                )

            if richness(rec.to_dict()) > richness(existing):
                records[rec.id] = rec.to_dict()
            continue

        records[rec.id] = rec.to_dict()

        if verbose and idx % 500 == 0:
            print(f"[PARSE] processed {idx}/{len(urls)} urls, kept {len(records)} records")

    metrics = compute_quality_metrics(records)
    ok, warn_list = validate_quality(metrics)

    # Write docs cache
    os.makedirs(os.path.dirname(out_json) or ".", exist_ok=True)
    header = {
        "generated_at_utc": utc_now_iso(),
        "record_count": len(records),
        "schema_version": SCHEMA_VERSION,
        "scraper_version": SCRAPER_VERSION,
        "quality_metrics": metrics,
        "quality_warnings": warn_list,
        "records": records,
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(header, f, ensure_ascii=False, indent=2, sort_keys=True)

    with open(errors_json, "w", encoding="utf-8") as f:
        json.dump(
            {
                "generated_at_utc": utc_now_iso(),
                "parse_error_count": len(parse_errors),
                "errors": parse_errors,
            },
            f,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )

    with open(report_json, "w", encoding="utf-8") as f:
        json.dump(
            {
                "generated_at_utc": utc_now_iso(),
                "ok": ok,
                "quality_metrics": metrics,
                "quality_warnings": warn_list,
            },
            f,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )

    return len(records), len(parse_errors)


# -----------------------------
# CLI commands
# -----------------------------

def cmd_harvest(args: argparse.Namespace) -> int:
    out_dir = args.out_dir
    os.makedirs(out_dir, exist_ok=True)

    seeds = args.seeds or DEFAULT_SEEDS
    if args.seeds_file:
        with open(args.seeds_file, "r", encoding="utf-8") as f:
            seeds = [ln.strip() for ln in f if ln.strip() and not ln.strip().startswith("#")]

    norm_seeds: List[str] = []
    for s in seeds:
        s = canonicalize_url(s)
        if is_allowed_url(s):
            norm_seeds.append(s)
        else:
            print(f"[WARN] seed not allowed by scope filters: {s}")

    h = Harvester(
        out_dir=out_dir,
        throttle_s=args.throttle,
        max_pages=args.max_pages,
        max_depth=args.max_depth,
        verbose=not args.quiet,
        debug=args.debug,
    )

    print(f"[INFO] crawling {len(norm_seeds)} seeds, max_pages={args.max_pages}, max_depth={args.max_depth}")
    urls, fetch_errors = h.crawl(norm_seeds)

    manifest_path = os.path.join(out_dir, "crawl_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "generated_at_utc": utc_now_iso(),
                "fetched_count": len(urls),
                "error_count": len(fetch_errors),
                "urls": urls,
                "fetch_errors": fetch_errors,
                "url_cache_map": h.url_cache_map,  # debug friendliness
            },
            f,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )

    print(f"[INFO] fetched {len(urls)} pages, {len(fetch_errors)} fetch errors (manifest: {manifest_path})")

    out_json = os.path.join(out_dir, "docs_cache.json")
    errors_json = os.path.join(out_dir, "docs_cache.errors.json")
    report_json = os.path.join(out_dir, "docs_cache.report.json")

    kept, parse_errs = build_docs_cache(
        raw_dir=os.path.join(out_dir, "raw"),
        urls=urls,
        url_cache_map=h.url_cache_map,
        out_json=out_json,
        errors_json=errors_json,
        report_json=report_json,
        verbose=not args.quiet,
    )

    print(f"[INFO] wrote docs cache: {out_json}")
    print(f"[INFO] wrote report: {report_json}")
    print(f"[INFO] records kept: {kept} | parse errors: {parse_errs} (see {errors_json})")

    # If validation is requested (or strict), enforce
    if args.strict:
        with open(report_json, "r", encoding="utf-8") as f:
            rep = json.load(f)
        if not rep.get("ok", True):
            print("[ERROR] Validation failed (strict mode). See docs_cache.report.json for details.")
            return 2

    return 0

def cmd_ensure(args: argparse.Namespace) -> int:
    out_dir = args.out_dir
    cache_path = os.path.join(out_dir, "docs_cache.json")
    if os.path.exists(cache_path):
        print(f"[OK] docs cache exists: {cache_path}")
        return 0

    print(f"[INFO] docs cache not found at {cache_path}; running harvester...")
    harvest_args = argparse.Namespace(
        out_dir=out_dir,
        throttle=args.throttle,
        max_pages=args.max_pages,
        max_depth=args.max_depth,
        seeds=args.seeds,
        seeds_file=args.seeds_file,
        quiet=args.quiet,
        debug=getattr(args, "debug", False),
        strict=args.strict,
    )
    return cmd_harvest(harvest_args)

def cmd_validate(args: argparse.Namespace) -> int:
    out_dir = args.out_dir
    cache_path = os.path.join(out_dir, "docs_cache.json")
    if not os.path.exists(cache_path):
        print(f"[ERROR] docs cache missing: {cache_path}")
        return 2

    with open(cache_path, "r", encoding="utf-8") as f:
        blob = json.load(f)

    metrics = blob.get("quality_metrics", {})
    warnings = blob.get("quality_warnings", [])
    ok, warn_list = validate_quality(metrics if isinstance(metrics, dict) else {})

    print("[INFO] Cache validation:")
    print(json.dumps({"ok": ok, "metrics": metrics, "warnings": warnings or warn_list}, indent=2))

    if args.strict and not ok:
        return 2
    return 0

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Microsoft Learn docs harvester (Excel VBA + Power Query M)")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add_common(sp):
        sp.add_argument("--out-dir", default="doc_cache", help="Output directory (raw html + docs_cache.json)")
        sp.add_argument("--throttle", type=float, default=1.0, help="Seconds to sleep between HTTP requests")
        sp.add_argument("--max-pages", type=int, default=8000, help="Maximum pages to fetch")
        sp.add_argument("--max-depth", type=int, default=7, help="Maximum crawl depth from seeds")
        sp.add_argument("--seeds-file", default=None, help="File with seed URLs (one per line)")
        sp.add_argument("--seeds", nargs="*", default=None, help="Seed URLs (overrides defaults if provided)")
        sp.add_argument("--quiet", action="store_true", help="Reduce logging")
        sp.add_argument("--debug", action="store_true", help="Enable per-request debug logging")
        sp.add_argument("--strict", action="store_true", help="Fail if validation indicates extraction collapse")

    sp_harvest = sub.add_parser("harvest", help="Crawl + build docs_cache.json")
    add_common(sp_harvest)
    sp_harvest.set_defaults(func=cmd_harvest)

    sp_ensure = sub.add_parser("ensure", help="Ensure docs_cache.json exists (harvest if missing)")
    add_common(sp_ensure)
    sp_ensure.set_defaults(func=cmd_ensure)

    sp_validate = sub.add_parser("validate", help="Validate existing docs_cache.json quality metrics")
    add_common(sp_validate)
    sp_validate.set_defaults(func=cmd_validate)

    return p

def main(argv: List[str]) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    return args.func(args)

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

"""Serper.dev Google Search API wrapper."""

import json
import ssl
import urllib.request

import config

BASE_URL = "https://google.serper.dev/search"

# Proxy for accessing Google services from China
PROXY = "127.0.0.1:51081"


def _get_opener():
    """Create an opener with proxy and relaxed SSL (for MITM proxy)."""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    proxy_handler = urllib.request.ProxyHandler({
        "https": PROXY,
    })
    https_handler = urllib.request.HTTPSHandler(context=ctx)
    return urllib.request.build_opener(proxy_handler, https_handler)


_opener = None


def google_search(query, num=10, page=1, gl=None, hl=None):
    """Search Google via Serper.dev. Returns [{title, snippet, link}, ...].
    num: 10, 20, 40, 60, 80, or 100 (max per call).
    page: page number for pagination.
    """
    global _opener
    if _opener is None:
        _opener = _get_opener()

    body = {
        "q": query,
        "num": min(num, 100),
        "page": page,
    }
    if gl:
        body["gl"] = gl
    if hl:
        body["hl"] = hl

    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        BASE_URL,
        data=data,
        headers={
            "X-API-KEY": config.get_service_key("serper"),
            "Content-Type": "application/json",
        },
        method="POST",
    )

    with _opener.open(req, timeout=180) as resp:
        result = json.loads(resp.read())

    results = []
    for item in result.get("organic", []):
        results.append({
            "title": item.get("title", ""),
            "snippet": item.get("snippet", ""),
            "link": item.get("link", ""),
        })
    return results


def google_search_all(query, num=100, gl=None, hl=None):
    """Search Google for up to `num` results (max 100, 10 per page).
    Returns [{title, snippet, link}, ...].
    """
    results = []
    for page in range(1, 12):  # Google max 10 pages
        if len(results) >= num:
            break
        try:
            batch = google_search(query, num=10, page=page, gl=gl, hl=hl)
            if not batch:
                break
            results.extend(batch)
        except Exception as e:
            print(f"[serper] error at page={page}: {e}")
            break
    return results[:num]


# ---- Google Scholar Search (Serper.dev) ----

SCHOLAR_URL = "https://google.serper.dev/scholar"


def _parse_publication_info(pub_info):
    """Parse Google Scholar publicationInfo into (authors, publication).
    Typical format: 'Author1, Author2 - Journal Name, Year'
    Returns (authors, publication) — both may be empty strings.
    """
    if not pub_info:
        return "", ""
    # Split on ' - ' (space-hyphen-space) to separate authors from venue
    parts = pub_info.split(" - ", 1)
    if len(parts) == 2:
        left, right = parts
        # If left part looks like a journal name (contains common journal words
        # and no person-name patterns), treat the whole string as publication
        journal_words = ["journal", "review", "press", "conference", "proceedings",
                         "university", "institute", "international", "research",
                         "science", "technology", "annual", "symposium", "society"]
        left_lower = left.lower()
        looks_like_journal = any(w in left_lower for w in journal_words)
        # Also check: if left has 5+ words, likely not authors
        if looks_like_journal and len(left.split()) >= 4:
            return "", pub_info
        return left.strip(), right.strip()
    return "", pub_info.strip()


def scholar_search(query, num=10, page=1):
    """Search Google Scholar via Serper.dev.
    Returns [{title, authors, year, publication, snippet, pdfUrl, citedBy, link}, ...].
    """
    global _opener
    if _opener is None:
        _opener = _get_opener()

    body = {
        "q": query,
        "num": min(num, 100),
        "page": page,
    }

    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        SCHOLAR_URL,
        data=data,
        headers={
            "X-API-KEY": config.get_service_key("serper"),
            "Content-Type": "application/json",
        },
        method="POST",
    )

    with _opener.open(req, timeout=180) as resp:
        result = json.loads(resp.read())

    results = []
    for item in result.get("organic", []):
        pub_info = item.get("publicationInfo", "")
        authors, publication = _parse_publication_info(pub_info)
        results.append({
            "title": item.get("title", ""),
            "authors": authors,
            "year": item.get("year", ""),
            "publication": publication or pub_info,
            "snippet": item.get("snippet", ""),
            "pdfUrl": item.get("pdfUrl", ""),
            "citedBy": item.get("citedBy", 0),
            "link": item.get("link", ""),
        })
    return results


def scholar_search_all(query, num=50):
    """Search Google Scholar for up to `num` results.
    Returns [{title, snippet, year, publicationInfo, pdfUrl, citedBy, link}, ...].
    """
    results = []
    for page in range(1, 12):
        if len(results) >= num:
            break
        try:
            batch = scholar_search(query, num=10, page=page)
            if not batch:
                break
            results.extend(batch)
        except Exception as e:
            print(f"[scholar] error at page={page}: {e}")
            break
    return results[:num]


# ---- Baidu Search (Qianfan AppBuilder) ----

BAIDU_BASE_URL = "https://qianfan.baidubce.com/v2/ai_search/web_search"

_baidu_ssl_ctx = None


def baidu_search(query, top_k=50):
    """Search Baidu via Qianfan AppBuilder API. Returns [{title, snippet, link}, ...].
    top_k: max 50 per call (API limit).
    """
    global _baidu_ssl_ctx
    if _baidu_ssl_ctx is None:
        _baidu_ssl_ctx = ssl.create_default_context()
        _baidu_ssl_ctx.check_hostname = False
        _baidu_ssl_ctx.verify_mode = ssl.CERT_NONE

    body = {
        "messages": [{"content": query, "role": "user"}],
        "search_source": "baidu_search_v2",
        "resource_type_filter": [{"type": "web", "top_k": min(top_k, 50)}],
    }
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        BAIDU_BASE_URL,
        data=data,
        headers={
            "X-Appbuilder-Authorization": f"Bearer {config.get_service_key('baidu')}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    with urllib.request.urlopen(req, context=_baidu_ssl_ctx, timeout=180) as resp:
        result = json.loads(resp.read())

    results = []
    for item in result.get("references", []):
        results.append({
            "title": item.get("title", ""),
            "snippet": item.get("content", ""),
            "link": item.get("url", ""),
        })
    return results[:top_k]

"""Smart search: paginated Google search with AI filtering and early stopping.

Stops when N consecutive pages return zero relevant results. No hard page cap.
"""

import database as db
import config
from utils.web_search import google_search
from utils.ai_extractor import _call_yunwu

MAX_PAGES = 50  # safety cap, should never be reached with early stopping


def ai_filter_batch(results, query, api_key):
    """Filter a batch of search results with AI. Returns [(index, result, reason), ...]."""
    if not results:
        return []

    prompt = (
        f"以下是搜索「{query}」的{len(results)}条结果。"
        "判断标准：同时涉及「中国/中式/中华/华人/唐人街」和「建筑/寺庙/园林/住宅/民居/宫殿/亭台楼阁」的条目，视为相关。\n\n"
        "返回JSON格式：{\"relevant\": [1, 3, 5], \"reason\": {\"1\": \"一句话理由\", ...}}\n"
        "只返回JSON，不要其他内容。\n\n"
    )
    for i, r in enumerate(results, 1):
        prompt += f"{i}. 标题: {r['title']}\n   摘要: {r['snippet']}\n\n"

    try:
        response = _call_yunwu(prompt, api_key, temperature=0.0, max_tokens=2000)
    except Exception as e:
        print(f"[ai_filter] API error: {e}")
        return [(i, r, "") for i, r in enumerate(results)]

    if isinstance(response, str):
        import json
        import re
        m = re.search(r'\{[\s\S]*\}', response)
        if m:
            try:
                response = json.loads(m.group())
            except json.JSONDecodeError:
                print(f"[ai_filter] failed to parse: {response[:200]}")
                return []

    if not isinstance(response, dict):
        return []

    relevant_ids = response.get("relevant", [])
    reasons = response.get("reason", {})

    out = []
    for rid in relevant_ids:
        idx = rid - 1
        if 0 <= idx < len(results):
            reason = reasons.get(str(rid), "")
            out.append((rid, results[idx], reason))
    return out


def smart_search(query, api_key=None, empty_stop=1, gl=None, hl=None):
    """Search Google with AI filtering and early stopping. No hard page cap.

    Paginates through results page by page. Stops when `empty_stop`
    consecutive pages have zero relevant results.

    Returns (results_list, stats_dict).
    """
    if api_key is None:
        api_key = config.get_llm("search_parse")[2]
    all_relevant = []
    all_raw = []
    empty_count = 0
    stats = {"pages_fetched": 0, "total_raw": 0, "total_relevant": 0}

    page = 1
    while page <= MAX_PAGES:
        print(f"  [page {page}] fetching...", end=" ")
        try:
            raw = google_search(query, num=10, page=page, gl=gl, hl=hl)
        except Exception as e:
            print(f"error: {e}")
            break

        stats["pages_fetched"] += 1
        stats["total_raw"] += len(raw)
        print(f"{len(raw)} raw", end=" ")

        if not raw:
            empty_count += 1
            print("-> empty")
            if empty_count >= empty_stop:
                print(f"  -> {empty_count} consecutive empty pages, stopping.")
                break
            page += 1
            continue

        all_raw.extend(raw)
        relevant = ai_filter_batch(raw, query, api_key)
        stats["total_relevant"] += len(relevant)
        print(f"-> {len(relevant)} relevant")

        if len(relevant) == 0:
            empty_count += 1
            if empty_count >= empty_stop:
                print(f"  -> {empty_count} consecutive empty pages, stopping.")
                break
        else:
            empty_count = 0
            for rid, r, reason in relevant:
                all_relevant.append({
                    "title": r["title"],
                    "snippet": r["snippet"],
                    "link": r["link"],
                    "relevance_reason": reason,
                })

        page += 1

    stats["total_relevant"] = len(all_relevant)
    if all_raw:
        try:
            db.insert_raw_results(query, "serper", all_raw)
        except Exception as e:
            print(f"[serper] save raw error: {e}")
    return all_relevant, stats


def baidu_smart_search(query, api_key=None):
    """Search Baidu with AI filtering. One call gets up to 50 results."""
    from utils.web_search import baidu_search

    if api_key is None:
        api_key = config.get_llm("search_parse")[2]
    all_relevant = []
    stats = {"total_raw": 0, "total_relevant": 0}

    print(f"  [baidu] fetching...", end=" ")
    try:
        raw = baidu_search(query, top_k=100)
    except Exception as e:
        print(f"error: {e}")
        return all_relevant, stats

    stats["total_raw"] = len(raw)
    print(f"{len(raw)} raw", end=" ")

    if not raw:
        return all_relevant, stats

    try:
        db.insert_raw_results(query, "baidu", raw)
    except Exception as e:
        print(f"[baidu] save raw error: {e}")

    # Split into batches of 25 for more stable AI filtering
    batch_size = 25
    for i in range(0, len(raw), batch_size):
        batch = raw[i:i + batch_size]
        relevant = ai_filter_batch(batch, query, api_key)
        stats["total_relevant"] += len(relevant)
        for rid, r, reason in relevant:
            all_relevant.append({
                "title": r["title"],
                "snippet": r["snippet"],
                "link": r["link"],
                "relevance_reason": reason,
            })

    print(f"-> {stats['total_relevant']} relevant")
    stats["total_relevant"] = len(all_relevant)
    return all_relevant, stats


def ai_confirm_duplicates(candidates, api_key=None):
    """Let AI confirm whether fuzzy-matched title pairs are truly duplicates.

    candidates: [(title_a, snippet_a, title_b, snippet_b), ...]
    Returns: set of indices (0-based) that ARE duplicates.
    """
    if api_key is None:
        api_key = config.get_llm("search_parse")[2]
    if not candidates:
        return set()

    prompt = (
        "以下是几对标题相似的搜索结果。请判断每对是否指的是同一个建筑/事件/内容。\n"
        "返回JSON: {\"duplicates\": [0, 2], \"reasons\": {\"0\": \"理由\", ...}}\n"
        "只返回JSON。\n\n"
    )
    for i, (ta, sa, tb, sb) in enumerate(candidates):
        prompt += f"{i}. A: {ta}\n   摘要A: {sa[:200]}\n   B: {tb}\n   摘要B: {sb[:200]}\n\n"

    try:
        response = _call_yunwu(prompt, api_key, temperature=0.0, max_tokens=1000)
    except Exception as e:
        print(f"[ai_confirm_dup] API error: {e}")
        return set(range(len(candidates)))  # on error, treat all as duplicates

    if isinstance(response, str):
        import re
        m = re.search(r'\{[\s\S]*\}', response)
        if m:
            try:
                import json
                response = json.loads(m.group())
            except json.JSONDecodeError:
                return set(range(len(candidates)))

    if not isinstance(response, dict):
        return set(range(len(candidates)))

    return set(response.get("duplicates", []))

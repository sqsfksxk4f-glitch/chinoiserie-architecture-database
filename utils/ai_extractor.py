"""
AI-powered Chinese architecture case extraction and verification using Yunwu API.
"""

import json
import re
import ssl
import urllib.request
from urllib.parse import quote
import openai

import config

API_BASE_URL = "https://api.deepseek.com/v1"
YUNWU_BASE_URL = "https://yunwu.ai/v1"
DEFAULT_MODEL = "deepseek-chat"

# ---- Phase 2: Upgraded extraction prompt - extracts all raw_cases fields ----

EXTRACTION_PROMPT = """你是一个建筑学研究的助手。请从以下文本中，提取所有提到的海外中式建筑案例名称及其详细信息。

文本可能来自多个网页，每个网页以 "--- [来源] http://..." 标记分隔。

重要提示：文本可能是任何语言（中文、英文、罗马尼亚语、法语、德语、日语、泰语、俄语等），你都能理解。请在各种语言文本中识别中国风格建筑（Chinese-style architecture）的案例，包括但不限于以下关键词的各种语言表述：Chinese temple、Chinese pagoda、Chinese garden、Chinese pavilion、Chinese house、Chinese palace、Chinese bridge、Chinatown、chinezăresc、chinois、chinesisch、cinese、китайский 等。

要求：
1. 识别文中提到的任何与中国风格相关的建筑（不管文本是哪种语言，只要描述的是中国式建筑就提取）
2. 提取每个建筑的以下信息（如果文中提到）：
   - name: 建筑的中文名称（如能从文中推断；如果文中只有当地语言名称，则用当地语言名称填入此字段）
   - name_en: 建筑的英文名称（如 Wikipedia 上的英文名，没有则填 null）
   - name_official: 建筑在当地官方语言的名称（原文中的写法，没有则填 null）
   - country: 所在国家
   - city: 所在城市
   - construction_year: 建造时间
   - demolition_year: 拆除时间
   - restoration_year: 修复时间
   - quantity: 数量规模（单体=1，建筑群=10，街区=50）
   - function: 建筑功能（寺庙/园林/民居/宫殿/衙署/塔/桥/祠堂/书院/现代中式/其他）
   - designer: 设计师或设计单位
   - latitude: 纬度（数字）
   - longitude: 经度（数字）
   - search_keywords: 用于搜索该建筑的最佳检索词（用原语言）
   - image_urls: 文中提到的图片链接（数组）
   - source_url: 来源URL（从最近的 "--- [来源]" 标记获取）
3. 以 JSON 数组返回，未提到的字段填 null
4. 只返回 JSON 数组，不要其他内容
5. 即使原文中没有中文名称，也请提取该建筑——只要原文描述的是中国风格建筑，就应该被提取

示例输出格式：
[{"name": "旧金山唐人街龙门", "name_en": "Dragon Gate of San Francisco Chinatown", "name_official": null, "country": "美国", "city": "旧金山", "construction_year": "1970", "demolition_year": null, "restoration_year": null, "quantity": 1, "function": "牌坊", "designer": null, "latitude": 37.79, "longitude": -122.41, "search_keywords": "San Francisco Chinatown Dragon Gate", "image_urls": ["http://example.com/img1.jpg"], "source_url": "http://example.com/article1"}]

以下是需要分析的文本：
"""

# ---- Verification prompt ----

VERIFICATION_PROMPT = """你是一个建筑学研究的审核助手。请评估以下海外中式建筑案例的真实性和可信度。

案例信息：
{case_data}

该案例的来源列表：
{sources}

请分析：
1. **各字段一致性**：各字段之间是否有矛盾？（如建造时间早于设计师出生年份，经纬度不在声称的城市等）
2. **来源可信度**：综合所有来源的可信度评级
3. **真实性评分**：0-1之间的分数
4. **风险字段**：哪些字段可能不准确，为什么
5. **推荐结论**：verified（可以入库）/ needs_review（需要人工复核）/ rejected（明显错误）

请以 JSON 格式返回：
{{"consistency_check": "分析文字", "credibility_summary": "分析文字", "authenticity_score": 0.85, "risk_fields": ["field1", "field2"], "risk_reasons": {{"field1": "原因", "field2": "原因"}}, "recommendation": "needs_review", "notes": "总体备注"}}
只返回 JSON，不要其他内容。
"""

# ---- Cross-validation prompt ----

CROSS_VALIDATION_PROMPT = """你是一个建筑学研究的审核助手。以下是从不同来源提取的可能指向同一建筑的多条记录。请判断它们是否是同一个建筑。

记录列表：
{cases_json}

请分析：
1. **名称相似度**：名称是否指同一建筑（考虑翻译差异、别名等）
2. **位置一致性**：国家和城市是否一致
3. **时间一致性**：建造时间等是否吻合
4. **字段互补性**：不同来源的信息是否可以互补
5. **判断结论**：same（是同一建筑，应合并）/ related（相关但不同，如扩建前后）/ different（完全不同）

请以 JSON 格式返回：
{{"name_similarity": "分析", "location_match": true/false, "time_match": true/false, "complementary_fields": ["field1"], "conclusion": "same", "merged_data": {{合并后的最佳数据}}, "notes": "备注"}}
只返回 JSON，不要其他内容。
"""

# ---- Field verification by web search ----

FIELD_VERIFY_PROMPT = """你是一个建筑学研究的审核助手。请利用下方【网络搜索结果】，逐个核验以下建筑案例的字段值是否正确。

当前案例数据：
{case_data}

【网络搜索结果】（来自百度 + Bing 实时搜索）：
{web_context}

核验要求：
对每个字段逐一分析：
1. 该字段的「当前值」是否能在搜索结果中找到印证？
2. 如果能印证 → status: "confirmed"，列出印证来源的编号
3. 如果搜索结果显示不同的值 → status: "conflict"，给出网上找到的值 + 建议替换为 + 来源编号
4. 如果搜索结果中完全没提到该字段 → status: "unverified"
5. 如果字段本身为空但网上能找到 → status: "found"，给出网上找到的值 + 来源编号

需要核验的字段：
- name (建筑名称): 搜索结果中这个建筑通常被称为什么？我们的名称准确吗？
- country (国家): 搜索结果是否确认建筑在该国？
- city (城市): 搜索结果是否确认建筑在该城市？
- construction_year (建造时间): 搜索结果中有没有提到建造年份？
- function (建筑类型/功能): 搜索结果如何描述该建筑的类型？
- designer (设计师): 搜索结果中有没有提到设计师或设计单位？
- latitude / longitude (经纬度): 搜索结果中有没有位置信息？
- description (描述): 搜索结果对该建筑的描述与我们的是否一致？

请以 JSON 格式返回：
{{
  "field_results": {{
    "name": {{"status": "confirmed/conflict/unverified/found", "current_value": "当前值", "web_value": "网上值或null", "suggestion": "建议或null", "source_indices": [1, 3], "confidence": "高/中/低"}},
    "country": {{...}},
    "city": {{...}},
    ... 每个需要核验的字段都返回
  }},
  "overall_assessment": "整体评估",
  "reliability_score": 0.85
}}
只返回 JSON，不要其他内容。"""


# ---- Completion prompt ----

COMPLETION_PROMPT = """你是一个建筑学研究的助手。以下案例信息不完整，请根据提供的原文内容，尝试补全缺失的字段。

当前案例数据：
{case_data}

缺失的字段：{missing_fields}

相关原文内容：
{source_texts}

请对每个缺失字段：
1. 判断能否从原文中找到信息
2. 如果能，提取具体值
3. 标注该信息的来源（从哪段原文）
4. 给出置信度（0-1）

请以 JSON 格式返回：
{{"completions": {{"field_name": {{"value": "补全的值", "source_text_index": 0, "confidence": 0.8}}, ...}}, "notes": "总体备注"}}
只返回 JSON，不要其他内容。
"""


# ---- Research & complete prompt (代替原文搜索，AI基于知识库研究) ----

RESEARCH_COMPLETE_PROMPT = """你是一个海外中式建筑研究的专家。请利用下方【网络搜索结果】，研究并补全该建筑的缺失字段。

当前案例数据：
{case_data}

缺失的字段：{missing_fields}

置信度判断标准：
- "高"：搜索结果中多个独立来源一致确认该值，或有官方/学术来源明确记载
- "中"：搜索结果中至少有一个来源提到该值，但缺乏交叉验证
- "低"：搜索结果中未直接提到该值，基于相关背景知识推测

研究要求：
1. 仔细阅读下方【网络搜索结果】，从中提取每个缺失字段的信息
2. 对每个缺失字段，必须给出：
   - value: 补全的值（如果搜索结果中完全找不到，填null）
   - confidence: 置信度（严格按上述标准判断："高"/"中"/"低"）
   - rationale: 推理依据（简要说明为什么给出这个值和这个置信度）
   - verification_hint: 建议用户去哪里核验
   - sources: 支撑该值的来源列表，每个来源包含：
     * index: 搜索结果编号（对应下方 [结果 X]）
     * description: 用一句话描述该来源如何支撑这个值（如"维基百科条目写明该寺建于1880年"）
     * 【重要】置信度为"高"时至少列出2个来源，置信度为"中"时至少列出1个来源，置信度为"低"时可为空数组[]
3. 注意字段之间的逻辑一致性（如建造时间不能晚于设计师去世年份等）

请以 JSON 格式返回：
{{"completions": {{"country": {{"value": "美国", "confidence": "高", "rationale": "多个来源一致确认", "verification_hint": "可在Wikipedia搜索该建筑确认", "sources": [{{"index": 1, "description": "维基百科将该建筑列为旧金山中式地标"}}, {{"index": 3, "description": "学术论文明确定位该建筑在美国加州"}}]}}, "city": {{...}}, ...}}, "notes": "总体研究备注"}}
只返回 JSON，不要其他内容。
"""

# ---- Verification of completed fields prompt ----

VERIFY_COMPLETION_PROMPT = """你是一个建筑学研究的审核助手。请评估以下补全后的案例数据的可信度。

原始案例数据：
{original_data}

AI 补全的字段：
{completed_fields}

请分析：
1. **字段一致性**：补全字段之间是否有矛盾？补全字段与原字段是否一致？
2. **整体可信度评分**：0-1之间的分数，反映补全数据整体的可信程度
3. **风险字段**：哪些补全字段可信度较低，存在风险
4. **推荐**："verified"（可以采纳）/ "needs_review"（需要人工复核部分字段）/ "low_confidence"（可信度低，建议仅作参考）

请以 JSON 格式返回：
{{"consistency_check": "分析文字", "overall_score": 0.75, "risk_fields": ["field1"], "risk_reasons": {{"field1": "原因"}}, "recommendation": "needs_review", "notes": "总体备注"}}
只返回 JSON，不要其他内容。
"""

# ============================================================
#  API calling functions
# ============================================================

def _resolve_channel(api_key):
    """Return (base_url, model) for the config channel whose key matches api_key.

    Falls back to (None, None) when api_key is empty or matches no channel.
    """
    if not api_key:
        return None, None
    cfg = config.load_config()
    for ch in config.LLM_CHANNELS:
        c = cfg.get(ch["id"], {})
        k = c.get("api_key", "")
        if k and k == api_key:
            return c.get("base_url", ""), c.get("model", "")
    return None, None


def _channel_client(channel_id, api_key=None, timeout=60):
    """OpenAI client + model for a config channel; api_key overrides if given."""
    base_url, model, cfg_key = config.get_llm(channel_id)
    return openai.OpenAI(base_url=base_url, api_key=api_key or cfg_key, timeout=timeout), model


def _call_yunwu(prompt, api_key, model=DEFAULT_MODEL, base_url=None, temperature=0.3, max_tokens=4000, timeout=60):
    """Call the LLM API and return parsed JSON result.

    base_url/model resolve from the config channel matching api_key when set;
    otherwise fall back to the model-name heuristic (gemini -> Yunwu).
    """
    cfg_base, cfg_model = _resolve_channel(api_key)
    if base_url is None:
        base_url = cfg_base or (YUNWU_BASE_URL if "gemini" in model.lower() else API_BASE_URL)
    if cfg_model and model == DEFAULT_MODEL:
        model = cfg_model
    client = openai.OpenAI(base_url=base_url, api_key=api_key, timeout=timeout)
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    content = response.choices[0].message.content.strip()
    return _parse_response(content)


def _extract_balanced_json(text):
    """Extract the first balanced JSON object or array from text.

    Returns (json_str, start, end) or (None, -1, -1) if not found.
    Unlike greedy regex, this correctly handles multiple JSON blocks in the same text.
    """
    # Find the first { or [
    first_brace = text.find('{')
    first_bracket = text.find('[')
    if first_brace == -1 and first_bracket == -1:
        return None, -1, -1
    if first_brace == -1:
        start = first_bracket
        open_char, close_char = '[', ']'
    elif first_bracket == -1:
        start = first_brace
        open_char, close_char = '{', '}'
    else:
        start = min(first_brace, first_bracket)
        open_char, close_char = ('{', '}') if start == first_brace else ('[', ']')

    depth = 0
    in_string = False
    escape = False
    for i, ch in enumerate(text[start:], start):
        if escape:
            escape = False
            continue
        if ch == '\\' and in_string:
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == open_char:
            depth += 1
            open_char = ch  # lock to the first opener type
            close_char = '}' if ch == '{' else ']'
        elif ch == ('}' if open_char == '{' else ']'):
            depth -= 1
            if depth == 0:
                return text[start:i + 1], start, i + 1
    return None, -1, -1


def _sanitize_json(json_str):
    """Remove common JSON formatting issues from LLM output."""
    # Remove trailing commas before } or ]
    json_str = re.sub(r',\s*}', '}', json_str)
    json_str = re.sub(r',\s*]', ']', json_str)
    return json_str


def _parse_response(content):
    """Parse JSON from LLM response. Handles markdown code blocks."""
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)```', content)
    if json_match:
        content = json_match.group(1).strip()

    json_str, _, _ = _extract_balanced_json(content)
    if json_str:
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            # Try sanitizing and re-parsing
            try:
                return json.loads(_sanitize_json(json_str))
            except json.JSONDecodeError:
                pass

    return content


def _parse_response_with_text(content):
    """Parse a response that should contain both text description and JSON.

    Returns (parsed_dict, text_description) tuple.
    The text_description is everything that is not the JSON block.
    """
    text_description = ""
    json_str = None

    # Try ```json ... ``` code block first
    md_json = re.search(r'```json\s*([\s\S]*?)\s*```', content)
    if md_json:
        text_description = content[:md_json.start()].strip() + "\n" + content[md_json.end():].strip()
        text_description = text_description.strip()
        json_str = md_json.group(1)
    else:
        # Find the outermost JSON object with known fields
        for pattern in [r'"status"', r'"correct_value"', r'"chinese_name"']:
            field_match = re.search(pattern, content)
            if field_match:
                # Search backwards from the field to find the opening brace
                before = content[:field_match.start()]
                # Find the last { before this field (the opening brace of the JSON)
                brace_pos = before.rfind('{')
                if brace_pos >= 0:
                    json_str, _, _ = _extract_balanced_json(content[brace_pos:])
                    if json_str:
                        text_description = (content[:brace_pos] + " " + content[brace_pos + len(json_str):]).strip()
                break

    result = {}
    if json_str:
        try:
            result = json.loads(json_str)
        except json.JSONDecodeError:
            try:
                result = json.loads(_sanitize_json(json_str))
            except json.JSONDecodeError:
                pass

    if not result:
        # Fallback: try normal _parse_response
        result = _parse_response(content)
        if isinstance(result, dict):
            text_description = ""
        else:
            text_description = content
            result = {}

    return result, text_description


def _count_chinese_chars(text):
    """Count the proportion of Chinese characters in text."""
    chinese = sum(1 for c in text if '一' <= c <= '鿿' or '㐀' <= c <= '䶿')
    total = len(text.strip())
    return chinese / max(total, 1)


def _translate_to_chinese(text, api_key, model=DEFAULT_MODEL):
    """Translate text to Chinese using the API."""
    prompt = f"""请将以下文本翻译成中文。保留所有地名、建筑名、人名的原文拼写（在括号中标注）。
只返回翻译后的中文文本，不要其他内容。

原文：
{text[:8000]}"""
    result = _call_yunwu(prompt, api_key, model=model, temperature=0.1, max_tokens=4000)
    if isinstance(result, str):
        return result
    if isinstance(result, list) and len(result) > 0:
        return str(result[0])
    return text


# ============================================================
#  Public functions
# ============================================================

def extract_case_names(text, api_key, model=DEFAULT_MODEL):
    """Extract Chinese architecture case names and details from text.

    Returns a list of dicts matching raw_cases fields.
    """
    if not text or not text.strip():
        return []

    # If text has very few Chinese characters, translate first
    cn_ratio = _count_chinese_chars(text)
    if cn_ratio < 0.05:
        try:
            translated = _translate_to_chinese(text, api_key, model=model)
            # Use translated text as primary, keep original for reference
            combined_text = f"[中文翻译]\n{translated}\n\n[原文]\n{text[:2000]}"
        except Exception:
            combined_text = text
    else:
        combined_text = text

    result = _call_yunwu(EXTRACTION_PROMPT + combined_text, api_key, model=model)
    if isinstance(result, list):
        return result
    return [{"name": str(result), "country": None, "city": None}]


def verify_authenticity(case_data, sources, api_key, model=DEFAULT_MODEL):
    """Evaluate the authenticity and credibility of a raw case.

    case_data: dict of case fields
    sources: list of dicts with keys: url, credibility, title
    Returns: dict with authenticity_score, recommendation, risk_fields, etc.
    """
    prompt = VERIFICATION_PROMPT.format(
        case_data=json.dumps(case_data, ensure_ascii=False, indent=2),
        sources=json.dumps(sources, ensure_ascii=False, indent=2),
    )
    result = _call_yunwu(prompt, api_key, model=model)
    return result if isinstance(result, dict) else {"notes": str(result)}


def cross_validate_cases(cases, api_key, model=DEFAULT_MODEL):
    """Compare multiple raw cases and determine if they refer to the same building.

    cases: list of dicts (raw case data)
    Returns: dict with conclusion (same/related/different) and merged_data.
    """
    prompt = CROSS_VALIDATION_PROMPT.format(
        cases_json=json.dumps(cases, ensure_ascii=False, indent=2),
    )
    result = _call_yunwu(prompt, api_key, model=model)
    return result if isinstance(result, dict) else {"conclusion": "different", "notes": str(result)}


def verify_case_by_search(case_data, api_key, model=DEFAULT_MODEL, on_progress=None):
    """Field-by-field verification using real web search (Bing).

    Step 1: Build search queries, search Bing CN + EN in parallel.
    Step 2: Pass search results + case data to GPT-4o to verify each field.

    on_progress(msg) — optional callback for status updates.

    Returns: dict with field_results (per-field status), overall_assessment, reliability_score.
    """
    if on_progress:
        on_progress("(1/2) 正在 Bing 搜索...")

    name = case_data.get("name", "")
    country = case_data.get("country", "")
    city = case_data.get("city", "")

    # Build search queries
    cn_query = " ".join(p for p in [name, country, city] if p)
    en_keywords = _parse_keywords(case_data.get("search_keywords", ""))
    en_query = " ".join(en_keywords[:3]) if en_keywords else cn_query

    # Parallel search: Bing CN + Bing EN
    search_results = []
    from concurrent.futures import ThreadPoolExecutor, as_completed
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = {
            pool.submit(_search_bing, cn_query, 6): "bing_cn",
        }
        if en_query and en_query != cn_query:
            futures[pool.submit(_search_bing, en_query, 6)] = "bing_en"

        for future in as_completed(futures):
            lang = futures[future]
            try:
                results = future.result()
                for r in results:
                    r["search_source"] = lang
                search_results.extend(results)
            except Exception:
                pass

    # Deduplicate by URL
    seen_urls = set()
    deduped = []
    for r in search_results:
        url = r.get("url", "")
        if url and url in seen_urls:
            continue
        seen_urls.add(url)
        deduped.append(r)

    if on_progress:
        on_progress(f"(2/2) 已找到 {len(deduped)} 条结果，正在 AI 分析...")

    # Format search results as numbered context
    search_sources = []
    web_parts = []
    for i, r in enumerate(deduped[:10]):
        idx = i + 1
        search_sources.append({
            "index": idx,
            "url": r.get("url", ""),
            "title": r.get("title", ""),
            "source": r.get("search_source", "unknown"),
        })
        tag = {"bing_cn": "[Bing-中文]", "bing_en": "[Bing-English]"}.get(
            r.get("search_source", ""), "")
        web_parts.append(
            f"[结果 {idx}] {tag} {r.get('title', '')}\n"
            f"URL: {r.get('url', '')}\n"
            f"摘要: {r.get('snippet', '')}"
        )

    web_context = "\n\n---\n\n".join(web_parts) if web_parts else "（搜索引擎未返回结果）"

    # Call AI to verify each field against search results
    prompt = FIELD_VERIFY_PROMPT.format(
        case_data=json.dumps(case_data, ensure_ascii=False, indent=2),
        web_context=web_context,
    )

    result = _call_yunwu(prompt, api_key, model=model, temperature=0.3, max_tokens=3000)
    if isinstance(result, dict):
        result["search_sources"] = search_sources
        result["search_queries"] = {"cn": cn_query, "en": en_query}
        return result
    return {"notes": str(result), "search_sources": search_sources,
            "search_queries": {"cn": cn_query, "en": en_query}}


def complete_missing_fields(case_data, missing_fields, source_texts, api_key, model=DEFAULT_MODEL):
    """Analyze source texts and suggest completions for missing fields.

    case_data: dict of current case fields
    missing_fields: list of field names to try to fill
    source_texts: list of (text, url) tuples containing original source content
    Returns: dict with completions keyed by field_name
    """
    formatted_sources = []
    for i, (text, url) in enumerate(source_texts):
        formatted_sources.append(f"[原文 {i}] 来源: {url}\n{text[:2000]}")
    sources_block = "\n\n---\n\n".join(formatted_sources)

    prompt = COMPLETION_PROMPT.format(
        case_data=json.dumps(case_data, ensure_ascii=False, indent=2),
        missing_fields=json.dumps(missing_fields, ensure_ascii=False),
        source_texts=sources_block,
    )
    result = _call_yunwu(prompt, api_key, model=model)
    return result if isinstance(result, dict) else {"notes": str(result)}


# ---- Real web search (Bing) ----

_BING_SSL = ssl.create_default_context()
_BING_SSL.check_hostname = False
_BING_SSL.verify_mode = ssl.CERT_NONE
_BING_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}


def _search_bing(query, max_results=8):
    """Real web search via Bing. Returns list of {title, url, snippet} dicts."""
    url = f"https://www.bing.com/search?q={quote(query)}&count={max_results}"
    try:
        req = urllib.request.Request(url, headers=_BING_HEADERS)
        with urllib.request.urlopen(req, timeout=12, context=_BING_SSL) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
    except Exception:
        return []

    results = []
    # Bing wraps organic results in <li class="b_algo"> blocks
    blocks = re.findall(r'<li class="b_algo"[^>]*?>(.*?)</li>', html, re.DOTALL)

    for block in blocks[:max_results]:
        # --- URL ---
        url_match = re.search(r'<a[^>]*href="(https?://[^"]+)"', block)
        if not url_match:
            continue
        raw_url = url_match.group(1)

        # --- Title: <h2>...<a href="...">title text</a>...</h2> ---
        h2_match = re.search(r'<h2[^>]*>(.*?)</h2>', block, re.DOTALL)
        title = ""
        if h2_match:
            title = re.sub(r'<[^>]+>', ' ', h2_match.group(1))
            title = re.sub(r'\s+', ' ', title).strip()

        # --- Snippet: text in the block after the heading area ---
        snippet = ""
        # Try <p> elements in the block
        p_match = re.search(r'<p[^>]*>(.*?)</p>', block, re.DOTALL)
        if p_match:
            snippet = re.sub(r'<[^>]+>', ' ', p_match.group(1))
        else:
            # Fallback: strip all HTML and use remaining text
            snippet = re.sub(r'<[^>]+>', ' ', block)
        snippet = re.sub(r'\s+', ' ', snippet).strip()
        # Trim if too long
        if len(snippet) > 300:
            snippet = snippet[:300] + "..."

        if title and raw_url:
            results.append({"title": title, "url": raw_url, "snippet": snippet})

    return results


def _search_baidu(query, max_results=8):
    """Real web search via Baidu. Returns list of {title, url, snippet} dicts."""
    url = f"https://www.baidu.com/s?wd={quote(query)}&rn={max_results}"
    try:
        req = urllib.request.Request(url, headers=_BING_HEADERS)
        with urllib.request.urlopen(req, timeout=12, context=_BING_SSL) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
    except Exception:
        return []

    results = []
    # Baidu wraps results in <div class="result c-container"> or similar
    blocks = re.findall(r'<div[^>]*class="[^"]*result[^"]*c-container[^"]*"[^>]*>(.*?)</div>\s*</div>', html, re.DOTALL)
    if not blocks:
        # Fallback: look for h3 with links
        blocks = re.findall(r'<div[^>]*class="[^"]*result[^"]*"[^>]*>(.*?)</div>\s*</div>', html, re.DOTALL)
    for block in blocks[:max_results]:
        # Title + URL from <h3><a href="...">Title</a></h3>
        link_match = re.search(r'<a[^>]*href="(https?://[^"]+)"[^>]*>(.*?)</a>', block, re.DOTALL)
        if not link_match:
            # Baidu sometimes wraps URLs in data-showurl or other attributes
            link_match = re.search(r'href="(https?://[^"]+)"', block)
        if not link_match:
            continue
        raw_url = link_match.group(1)
        title = re.sub(r'<[^>]+>', ' ', link_match.group(2) if link_match.lastindex >= 2 else '')
        title = re.sub(r'\s+', ' ', title).strip()
        if not title:
            # Try to find text content
            text_match = re.search(r'<h3[^>]*>(.*?)</h3>', block, re.DOTALL)
            if text_match:
                title = re.sub(r'<[^>]+>', ' ', text_match.group(1))
                title = re.sub(r'\s+', ' ', title).strip()

        # Snippet
        snippet = re.sub(r'<[^>]+>', ' ', block)
        snippet = re.sub(r'\s+', ' ', snippet).strip()
        if len(snippet) > 300:
            snippet = snippet[:300] + "..."

        if title and raw_url:
            results.append({"title": title, "url": raw_url, "snippet": snippet})

    return results


# ---- Keyword helpers ----

def _parse_keywords(raw_keywords):
    """Parse search_keywords from JSON string or list into a flat list of strings."""
    if not raw_keywords:
        return []
    if isinstance(raw_keywords, list):
        return [str(k) for k in raw_keywords if k]
    try:
        parsed = json.loads(raw_keywords)
        if isinstance(parsed, list):
            return [str(k) for k in parsed if k]
        if isinstance(parsed, str) and parsed.strip():
            return [parsed.strip()]
    except (json.JSONDecodeError, TypeError):
        pass
    # Fallback: comma-separated plain text
    return [k.strip() for k in str(raw_keywords).replace("\n", ",").split(",") if k.strip()]


def _build_search_context(case_data, missing_fields):
    """Assemble rich search queries and context from all available case fields.

    Returns a dict with everything the search step needs.
    """
    name = case_data.get("name", "")
    country = case_data.get("country", "")
    city = case_data.get("city", "")
    function = case_data.get("function") or case_data.get("building_type", "")
    designer = case_data.get("designer") or case_data.get("architect", "")
    construction_year = case_data.get("construction_year") or case_data.get("year_built", "")
    keywords = _parse_keywords(case_data.get("search_keywords", ""))

    has_chinese = any('一' <= c <= '鿿' for c in name)

    # --- Chinese query ---
    cn_parts = [name]
    if country: cn_parts.append(country)
    if city: cn_parts.append(city)
    if function: cn_parts.append(function)
    if designer: cn_parts.append(designer)
    if construction_year: cn_parts.append(construction_year)
    # Add Chinese-looking keywords not already in the query
    cn_text = " ".join(cn_parts)
    for kw in keywords:
        if any('一' <= c <= '鿿' for c in kw) and kw not in cn_text:
            cn_parts.append(kw)
    cn_query = " ".join(cn_parts)

    # --- English query ---
    en_keywords = [k for k in keywords if not any('一' <= c <= '鿿' for c in k)]
    if en_keywords:
        en_query = " ".join(en_keywords[:5])
    elif not has_chinese:
        en_parts = [p for p in [name, country, city, function, designer, construction_year] if p]
        en_query = " ".join(en_parts)
    else:
        # Chinese name, no English keywords — include name+pinyin context
        en_parts = [name, country, city, "Chinese architecture"]
        en_parts = [p for p in en_parts if p]
        en_query = " ".join(en_parts)

    # --- Context block for the prompt ---
    context_lines = [f"建筑名称：{name}"]
    if country: context_lines.append(f"所在国家：{country}")
    if city: context_lines.append(f"所在城市：{city}")
    if function: context_lines.append(f"建筑类型：{function}")
    if designer: context_lines.append(f"设计师：{designer}")
    if construction_year: context_lines.append(f"建造年份：{construction_year}")
    if keywords: context_lines.append(f"已有检索词：{', '.join(keywords[:8])}")
    context_lines.append(f"需要补全的字段：{', '.join(missing_fields)}")

    return {
        "cn_query": cn_query,
        "en_query": en_query,
        "context_block": "\n".join(context_lines),
        "name": name,
        "country": country,
        "city": city,
    }


def _search_with_ai_model(search_context, missing_fields, api_key):
    """Use an AI model with web search capability to search and extract building info.

    search_context: dict from _build_search_context
    Returns: list of {title, snippet, url, source} dicts from AI search.
    """
    context_block = search_context["context_block"]

    prompt = f"""请搜索以下建筑的相关信息：

{context_block}

请搜索并返回：
1. 该建筑的基本信息（建造时间、设计师、建筑风格、功能、保护级别、面积等）
2. 信息来源的标题和URL

请以 JSON 格式返回搜索结果：
{{"findings": [{{"title": "来源标题", "url": "https://...", "snippet": "相关信息摘要", "relevance": "high/medium/low"}}], "summary": "综合搜索结果摘要"}}
只返回 JSON，不要其他内容。"""

    try:
        client, model = _channel_client("search_parse", api_key)
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=3000,
        )
        content = response.choices[0].message.content.strip()
        result = _parse_response(content)
        if isinstance(result, dict):
            findings = result.get("findings", [])
            return [{"title": f.get("title", ""), "snippet": f.get("snippet", ""),
                     "url": f.get("url", ""), "source": "ai_search"}
                    for f in findings if isinstance(f, dict) and f.get("title")]
    except Exception:
        pass
    return []


def research_and_complete(case_data, missing_fields, api_key, model=DEFAULT_MODEL):
    """Real web search to find info, then GPT-4o to complete missing fields.

    Step 1: Build rich search context from all available fields + search_keywords.
    Step 2: Real Bing searches — Chinese + English in parallel.
    Step 3: Merge & deduplicate results from both languages.
    Step 4: Completion model produces structured field completions, citing source indices.
    Step 5 (fallback): If Bing returns no results, try AI-based search.

    Returns: dict with completions, search_sources, notes.
    """
    search_context = _build_search_context(case_data, missing_fields)
    cn_query = search_context["cn_query"]
    en_query = search_context["en_query"]

    # Step 1-2: parallel bilingual REAL web search
    search_results = []
    from concurrent.futures import ThreadPoolExecutor, as_completed
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = {}
        if cn_query:
            futures[pool.submit(_search_bing, cn_query, 8)] = "cn"
        if en_query and en_query != cn_query:
            futures[pool.submit(_search_bing, en_query, 8)] = "en"

        for future in as_completed(futures):
            lang = futures[future]
            try:
                results = future.result()
                for r in results:
                    r["search_lang"] = lang
                    r["source"] = "bing"
                search_results.extend(results)
            except Exception:
                pass

    # Fallback: if Bing returned nothing, try AI-based search
    if not search_results:
        if cn_query:
            try:
                ai_results = _search_with_ai_model(search_context, missing_fields, api_key)
                for r in ai_results:
                    r["search_lang"] = "cn"
                    r["source"] = "ai_search"
                search_results.extend(ai_results)
            except Exception:
                pass
        if en_query and en_query != cn_query:
            try:
                en_context = dict(search_context)
                en_context["context_block"] = search_context["context_block"] + f"\nEnglish search query: {en_query}"
                ai_results = _search_with_ai_model(en_context, missing_fields, api_key)
                for r in ai_results:
                    r["search_lang"] = "en"
                    r["source"] = "ai_search"
                search_results.extend(ai_results)
            except Exception:
                pass

    # Deduplicate by URL (keep first occurrence)
    seen_urls = set()
    deduped = []
    for r in search_results:
        url = r.get("url", "")
        if url and url in seen_urls:
            continue
        seen_urls.add(url)
        deduped.append(r)

    # Step 3: format search results as numbered context for the AI
    search_sources = []  # index -> {url, title}
    web_context_parts = []
    for i, r in enumerate(deduped[:10]):
        idx = i + 1
        search_sources.append({
            "index": idx,
            "url": r.get("url", ""),
            "title": r.get("title", ""),
            "lang": r.get("search_lang", "cn"),
        })
        lang_tag = "[EN]" if r.get("search_lang") == "en" else "[CN]"
        context = (
            f"[结果 {idx}] {lang_tag} {r.get('title', '')}\n"
            f"URL: {r.get('url', '')}\n"
            f"摘要: {r.get('snippet', '')}"
        )
        web_context_parts.append(context)

    web_context = (
        "\n\n---\n\n".join(web_context_parts)
        if web_context_parts
        else "（搜索引擎和AI均未返回结果，请基于训练数据尝试补全）"
    )

    # Step 4: AI completion with search context
    prompt = RESEARCH_COMPLETE_PROMPT.format(
        case_data=json.dumps(case_data, ensure_ascii=False, indent=2),
        missing_fields=json.dumps(missing_fields, ensure_ascii=False),
    )
    source_label = "Bing 实时搜索" if any(r.get("source") == "bing" for r in deduped) else "AI辅助搜索"
    prompt += (
        f"\n\n【网络搜索结果（{source_label}，中英文双语检索，共 {len(deduped)} 条）】"
        f"以下是为该建筑搜索到的网络信息，请引用结果编号（如 [结果 1]）：\n\n{web_context}"
    )

    result = _call_yunwu(prompt, api_key, model=model, temperature=0.4, max_tokens=4000)
    if isinstance(result, dict):
        result["search_sources"] = search_sources
        result["search_queries"] = {"cn": cn_query, "en": en_query}
        return result
    return {"notes": str(result), "search_sources": search_sources,
            "search_queries": {"cn": cn_query, "en": en_query}}


def verify_completed_case(case_data, completions, api_key, model=DEFAULT_MODEL):
    """Verify the authenticity of AI-completed fields against the original case data.

    case_data: dict of original case fields
    completions: dict of AI-completed fields (from research_and_complete)
    api_key: Yunwu API key

    Returns: dict with consistency_check, overall_score, risk_fields, recommendation, notes.
    """
    completed_summary = {}
    for field_name, info in completions.items():
        if isinstance(info, dict):
            completed_summary[field_name] = {
                "value": info.get("value"),
                "confidence": info.get("confidence"),
                "rationale": info.get("rationale", ""),
            }

    if not completed_summary:
        return {"notes": "无补全字段可核验", "overall_score": 0, "recommendation": "low_confidence"}

    prompt = VERIFY_COMPLETION_PROMPT.format(
        original_data=json.dumps(case_data, ensure_ascii=False, indent=2),
        completed_fields=json.dumps(completed_summary, ensure_ascii=False, indent=2),
    )
    result = _call_yunwu(prompt, api_key, model=model, temperature=0.3, max_tokens=3000)
    return result if isinstance(result, dict) else {"notes": str(result)}


# ---- Per-field verification from pasted search results ----

FIELD_VERIFY_FROM_PASTE_PROMPT = """你是一个建筑学研究的审核助手。用户用百度或谷歌搜索了一个建筑案例的某个字段，通过Chrome插件提取了搜索结果页面的文字内容。

【当前核验的字段】
字段名：{field_label}
当前值：{current_value}

【案例其他字段（供上下文参考）】
{case_context}

【搜索结果页面文字内容】
（注意：这是从浏览器直接提取的搜索结果页文字，夹杂了百度/谷歌的导航菜单、广告、相关搜索推荐等噪音。请从中筛选出真正的搜索结果条目进行分析。）
{pasted_content}

请仔细分析搜索结果中真正与这个建筑相关的条目，针对「{field_label}」这个字段进行核验：

1. **提取相关结果**：从所有文字中，筛选出与这个建筑真正相关的搜索结果条目（忽略导航栏、广告、热搜推荐等无关内容）
2. **找出该字段的所有说法**：在相关结果中，搜索这个字段的值。例如核验名称时，找出所有用到的名称变体
3. **统计频率**：各说法分别被多少个来源使用
4. **与当前值对比**：主流说法与当前值「{current_value}」是否一致？

请以 JSON 格式返回：
{{{{
  "status": "confirmed/conflict/found/unverified",
  "web_value": "搜索结果中最主流的说法",
  "all_variations": ["变体1", "变体2"],
  "variation_counts": {{{{"变体1": 5, "变体2": 3}}}},
  "suggestion": "具体建议，如：当前名称与主流名称一致，保持不变 / 建议将名称从X改为Y，因为Z个来源中使用Y / 网上未找到该字段信息，建议补充搜索",
  "analysis": "分析过程：列出找到的关键来源及其说法，说明判断依据",
  "confidence": "高/中/低"
}}}}

置信度判断标准：
- 高：3个以上独立来源一致确认
- 中：1-2个来源确认
- 低：仅在推测或未直接提到

只返回 JSON，不要其他内容。"""


# ---- Name-specific prompt for paste/search result analysis ----
FIELD_VERIFY_FROM_PASTE_NAME_PROMPT = """你是一个建筑学研究的助手。用户搜索了一个海外中式建筑的名称，通过Chrome插件提取了搜索结果页面的文字内容。

【当前数据库中记录的名称】
{current_value}

【案例其他字段（供上下文参考）】
{case_context}

【搜索结果页面文字内容】
（注意：这是从浏览器直接提取的搜索结果页文字，夹杂了导航菜单、广告、相关搜索推荐等噪音。请从中筛选出真正的搜索结果条目进行分析。）
{pasted_content}

请仔细分析搜索结果内容，找出这个建筑最常用、最准确的名称：

1. **提取所有名称变体**：从搜索结果中，找出这个建筑所有被提到的名称写法（包括中文名、英文名、当地语言名、简称、全称等）
2. **统计出现频率**：每个名称变体在多少个独立来源中被使用
3. **确定最常用名称**：找出出现频率最高的名称作为推荐名称
4. **与当前值对比**：当前数据库中的名称「{current_value}」是否就是最常用的名称？如果不是，应该替换为最常用的名称

请以 JSON 格式返回：
{{{{
  "status": "confirmed/conflict",
  "current_value": "数据库中的当前名称",
  "correct_value": "搜索结果中最常用、最准确的名称",
  "chinese_name": "中文名称（如适用）",
  "english_name": "英文名称（如适用）",
  "official_language_name": "当地官方语言名称（如适用）",
  "common_names": ["其他常用名称列表"],
  "all_variations": ["搜索结果中出现的所有名称变体"],
  "variation_counts": {{{{"变体1": 5, "变体2": 3}}}},
  "conclusion": "一句话总结：当前名称是否最常用，以及建议的名称",
  "reasoning": "详细分析：列出关键来源及其使用的名称，说明判断依据",
  "references": ["参考来源（包括描述和链接）"],
  "confidence": "高/中/低"
}}}}

状态说明：
- confirmed: 当前名称就是最常用的名称，无需修改
- conflict: 搜索结果中最常用的名称与当前名称不一致，建议替换为最常用名称

重要原则：完全依据搜索结果内容来判断，搜索结果中使用最多的名称就是正确的名称。

只返回 JSON，不要其他内容。"""


def verify_field_from_paste(field_name, field_label, current_value, case_context, pasted_content, api_key, model=DEFAULT_MODEL):
    """Verify a single field against search results pasted by the user.

    field_name: internal field name (e.g. "name", "country")
    field_label: display label (e.g. "建筑名称", "国家")
    current_value: current value of the field in the case
    case_context: dict of other case fields for context
    pasted_content: text content pasted from browser search results
    api_key: Yunwu API key

    Returns: dict with status, web_value, suggestion, analysis, confidence.
    """
    if not pasted_content or not pasted_content.strip():
        return {"status": "unverified", "web_value": None, "suggestion": "未提供搜索结果内容",
                "analysis": "", "confidence": "低"}

    # Build case context string (exclude the field being verified)
    context_parts = []
    for k, v in case_context.items():
        if v and k != field_name:
            label = k  # use raw field name as label
            context_parts.append(f"{label}: {v}")
    context_str = "\n".join(context_parts) if context_parts else "（无其他字段信息）"

    # Truncate pasted content if too long
    max_paste_len = 8000
    if len(pasted_content) > max_paste_len:
        pasted_content = pasted_content[:max_paste_len] + "\n\n... (内容过长，已截断)"

    prompt = FIELD_VERIFY_FROM_PASTE_PROMPT.format(
        field_label=field_label,
        current_value=current_value if current_value else "（空）",
        case_context=context_str,
        pasted_content=pasted_content,
    )

    result = _call_yunwu(prompt, api_key, model=model, temperature=0.3, max_tokens=2000)
    if isinstance(result, dict):
        return result
    return {"status": "unverified", "web_value": None, "suggestion": "AI分析失败",
            "analysis": str(result), "confidence": "低"}


def verify_field_name_from_paste(current_value, case_context, pasted_content, api_key, model=DEFAULT_MODEL):
    """Analyze search results to extract the most commonly used name for a building.

    Unlike verify_field_from_paste which does general field verification,
    this focuses on extracting the correct name from search results and
    determining whether the current name should be replaced.

    Returns: dict with status, correct_value, chinese_name, english_name,
             official_language_name, common_names, all_variations, etc.
    """
    if not pasted_content or not pasted_content.strip():
        return {"status": "unverified", "correct_value": None,
                "conclusion": "未提供搜索结果内容", "reasoning": "", "confidence": "低"}

    # Build case context string
    context_parts = []
    for k, v in case_context.items():
        if v:
            context_parts.append(f"{k}: {v}")
    context_str = "\n".join(context_parts) if context_parts else "（无其他字段信息）"

    # Truncate pasted content if too long
    max_paste_len = 8000
    if len(pasted_content) > max_paste_len:
        pasted_content = pasted_content[:max_paste_len] + "\n\n... (内容过长，已截断)"

    prompt = FIELD_VERIFY_FROM_PASTE_NAME_PROMPT.format(
        current_value=current_value if current_value else "（空）",
        case_context=context_str,
        pasted_content=pasted_content,
    )

    result = _call_yunwu(prompt, api_key, model=model, temperature=0.3, max_tokens=2000)
    if isinstance(result, dict):
        return result
    return {"status": "unverified", "correct_value": None,
            "conclusion": "AI分析失败", "reasoning": str(result), "confidence": "低"}


# ---- Direct DeepSeek knowledge-based verification (no browser search) ----

FIELD_VERIFY_DEEPSEEK_PROMPT = """我是一名研究人员，正在核验建筑信息。
【建筑基本信息】
{case_info}

【当前需要核验的字段】
字段名：{field_label}
当前值：{current_value}

请按以下要求核验：

1. 判断这个字段的当前值是否准确。
2. 如果不准确，建议应该改为什么，依据是什么，依据应该包括描述和完整链接。
3. 如果准确，依据是什么，依据应该包括描述和完整链接。

返回结果必须包含两部分，按以下顺序输出：

【文字描述】
首先输出一段自然语言文字，内容包括：
- 结论（该字段值是否准确）
- 判断依据（详细描述推理过程，并附上完整的参考链接）

【JSON数据】
然后输出一个JSON代码块（用```json```包裹），格式如下：
{{
  "status": "confirmed/conflict/found/unverified",
  "current_value": "当前值",
  "correct_value": "正确或合理的值",
  "conclusion": "一句话总结判断",
  "reasoning": "详细推理过程",
  "references": ["相关参考来源（描述和完整链接）"],
  "confidence": "高/中/低"
}}

注意：文字描述和JSON数据缺一不可。"""

FIELD_VERIFY_DEEPSEEK_NAME_PROMPT = """我是一名研究人员，正在核验建筑信息。
建筑所在国家是{country}，建筑名称是{current_value}。

1. 判断这个建筑名称是否常用。
2. 如果不常用，应该改为什么。
3. 常用的名称，官方语言，英文，中文，其他名称，分别是什么。

返回结果必须包含两部分，按以下顺序输出：

【文字描述】
首先输出一段自然语言文字，内容包括：
- 结论（该名称是否常用、是否准确）
- 判断依据（详细描述推理过程，并附上完整的参考链接）

【JSON数据】
然后输出一个JSON代码块（用```json```包裹），格式如下：
{{
  "status": "confirmed或conflict",
  "current_value": "当前名称",
  "correct_value": "正确或建议的名称",
  "chinese_name": "中文名称",
  "english_name": "英文名称（没有则填null）",
  "official_language_name": "当地官方语言名称（没有则填null）",
  "common_names": ["其他常用名称"],
  "other_names": ["其他别名"],
  "conclusion": "一句话总结",
  "reasoning": "详细推理过程和判断依据",
  "references": ["参考来源（描述和完整链接）"],
  "confidence": "高/中/低"
}}

注意：文字描述和JSON数据缺一不可。"""

# Gemini-specific name verification prompt — leverages Gemini's multilingual strength
# and factual knowledge, no rigid DeepSeek-style 【】 structure needed.
FIELD_VERIFY_GEMINI_NAME_PROMPT = """You are an architectural historian specializing in overseas Chinese architecture. Verify the building name below.

Country: {country}
Building name to verify: {current_value}
Other case details: {case_info}

Step 1 — Cross-lingual lookup:
Search your knowledge for this building. Check the name across:
- Chinese (Simplified & Traditional)
- English
- Local official language of the country it's in
- Any other regional languages where it might be referenced

Step 2 — Determine the authoritative name:
Identify which name is most commonly used in academic/architectural references, and which is the official local name.

Step 3 — Output your findings:
First, write a natural-language analysis paragraph (in Chinese) covering:
- Whether the current name is accurate and commonly used
- What the correct/preferred name should be
- Reasoning with specific references and URLs

Then, output a JSON object wrapped in ```json```:
{{
  "status": "confirmed or conflict",
  "current_value": "the name being verified",
  "correct_value": "correct or recommended name",
  "chinese_name": "中文名称",
  "english_name": "English name (null if not found)",
  "official_language_name": "name in the local official language (null if not found)",
  "common_names": ["other commonly used names"],
  "other_names": ["other known aliases or historical names"],
  "conclusion": "one-sentence summary (in Chinese)",
  "reasoning": "detailed reasoning (in Chinese) with cross-lingual evidence",
  "references": ["source description with full URL"],
  "confidence": "高/中/低"
}}

Important: the text analysis and JSON are both required."""


def verify_field_deepseek(field_name, field_label, current_value, case_data, api_key, model=DEFAULT_MODEL):
    """Verify a single field using DeepSeek's knowledge base (no web search needed).

    field_name: internal field name
    field_label: display label
    current_value: current value of the field
    case_data: dict of all case fields for context
    api_key: DeepSeek API key

    Returns: dict with status, correct_value, conclusion, reasoning, references, confidence.
    """
    # Build case info summary
    info_lines = []
    key_fields = ["name", "country", "city", "function", "construction_year",
                  "designer", "latitude", "longitude"]
    if isinstance(case_data, dict):
        for f in key_fields:
            val = case_data.get(f)
            if val:
                info_lines.append(f"{f}: {val}")
    case_info = "\n".join(info_lines) if info_lines else "（无其他信息）"

    if not current_value:
        current_value = "（空，待补充）"

    if field_name == "name":
        country = case_data.get("country") if isinstance(case_data, dict) else ""
        prompt = FIELD_VERIFY_DEEPSEEK_NAME_PROMPT.format(
            country=country,
            current_value=str(current_value),
        )
    else:
        prompt = FIELD_VERIFY_DEEPSEEK_PROMPT.format(
            case_info=case_info,
            field_label=field_label,
            current_value=str(current_value),
        )

    result = _call_yunwu(prompt, api_key, model=model, temperature=0.3, max_tokens=2000)
    if isinstance(result, dict):
        return result
    return {"status": "unverified", "correct_value": None,
            "conclusion": "AI分析失败", "reasoning": str(result),
            "references": [], "confidence": "低"}


def verify_field_gemini(field_name, field_label, current_value, case_data, api_key):
    """Verify a single field using Gemini via Yunwu API.

    Uses Gemini-optimized prompts that leverage Gemini's multilingual
    capabilities and factual knowledge.
    """
    # Build case info summary
    info_lines = []
    key_fields = ["name", "country", "city", "function", "construction_year",
                  "designer", "latitude", "longitude"]
    if isinstance(case_data, dict):
        for f in key_fields:
            if f != field_name:  # exclude the field being verified
                val = case_data.get(f)
                if val:
                    info_lines.append(f"{f}: {val}")
    case_info = "\n".join(info_lines) if info_lines else "（无其他信息）"

    if not current_value:
        current_value = "（空，待补充）"

    # Name field: use Gemini-specific prompt (leverages multilingual strength)
    if field_name == "name":
        country = case_data.get("country") if isinstance(case_data, dict) else "未知"
        prompt = FIELD_VERIFY_GEMINI_NAME_PROMPT.format(
            country=country,
            current_value=str(current_value),
            case_info=case_info,
        )
    else:
        prompt = FIELD_VERIFY_DEEPSEEK_PROMPT.format(
            case_info=case_info,
            field_label=field_label,
            current_value=str(current_value),
        )

    # Gemini responses tend to be more verbose than DeepSeek;
    # name field needs extra tokens for multilingual name variants
    max_tokens = 8000 if field_name == "name" else 4000

    client, model = _channel_client("field_completion", api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=max_tokens,
    )
    content = response.choices[0].message.content.strip()

    # For name field: extract both text description and JSON
    text_description = ""
    if field_name == "name":
        result, text_description = _parse_response_with_text(content)
    else:
        result = _parse_response(content)

    if isinstance(result, dict):
        if text_description:
            result["text_description"] = text_description
        return result
    return {"status": "unverified", "correct_value": None,
            "conclusion": "AI分析失败", "reasoning": str(result),
            "references": [], "confidence": "低"}


# ---- One-click auto-complete ----

AUTO_COMPLETE_PROMPT = """你是一个海外中式建筑研究的专家。请根据建筑名称，利用你的知识库，补全以下案例的缺失字段。

建筑名称：{building_name}

当前已有信息：
{known_fields}

缺失的字段：{missing_fields}

请对每个缺失字段：
1. 根据建筑名称和你对海外中式建筑的了解，查找补全
2. 给出结论（一句话总结判断）
3. 给出数据依据（详细说明信息来源和推理过程，如果提到具体网站或来源名称，必须在references中附带该来源的完整URL）
4. 给出参考完整链接（可以访问的URL地址，必须是真实可访问的完整网址，如 https://www.example.com/page，不要只写网站名称）
5. 如果确实查不到，将 value 标注为 null

**重要：必须为待补全字段列表中的【每一个字段】返回结果。即使查不到，也要包含该字段并将value设为null。不要省略任何字段。**

字段值约束：
- quantity：单体建筑或一组建筑群填1，整村/整片区/大范围建筑群填10
- building_status 只能填：状态良好 / 破损待修复 / 损毁严重 / 拆除 / 未知（无法判断时填未知）
- spread_method 填一个简洁的词组（如：华人移民、传教士、政治决策、商业贸易、文化交流、殖民建设等，可自定义）
- coordinates 格式为 "纬度, 经度"（如 "37.79, -122.41"）
- 年份字段只填4位数字年份
- designer 不限于个人姓名，可以是建筑事务所、设计团体、施工组织、或身份描述（如"当地华人工匠"、"XX建筑事务所"、"寺院僧众"等），只要有相关信息就填入，不要因为查不到具体人名就返回null

请以 JSON 格式返回（必须包含待补全字段列表中的每一个字段）：
{{"completions": {{"city": {{"value": "San Francisco", "conclusion": "...", "reasoning": "...", "references": ["..."]}}, "designer": {{"value": null, "conclusion": "未查到设计师信息", "reasoning": "...", "references": []}} }}, "notes": "补充说明"}}
只返回 JSON，不要其他内容。"""

# Gemini-optimized auto-complete prompt — leverages Gemini's multilingual
# knowledge of overseas Chinese architecture for richer completions.
AUTO_COMPLETE_GEMINI_PROMPT = """You are an architectural historian specializing in overseas Chinese architecture. Complete missing fields for the building below using your knowledge.

{name_block}
Country: {country}

Already known:
{known_fields}

Fields to complete:
{missing_fields}

For each missing field:
1. Use your knowledge of Chinese diaspora architecture to find the correct value
2. Provide a conclusion (one-sentence summary of your judgment)
3. Provide reasoning (detailed explanation of information sources and inference process)
4. Provide reference links (accessible URLs)
5. If you cannot determine the value, set value to null
6. For quantity, use 1 for a single building or building complex, use 10 for a whole village/district/large area with many buildings
7. For coordinates, output as "latitude, longitude"
8. For building_status, ONLY use one of: "状态良好", "破损待修复", "损毁严重", "拆除", "未知" (use "未知" when unsure)
9. For spread_method, use a concise phrase (e.g. "华人移民", "传教士", "商业贸易", "文化交流", "殖民建设", etc. — can be customized)
10. For designer, it is NOT limited to individual names — it can be an architecture firm, design collective, construction organization, or identity description (e.g. "local Chinese craftsmen", "XX Architects", "temple monks", etc.). Fill in any relevant information; do NOT return null just because no specific person's name is found

**CRITICAL: You MUST return EVERY SINGLE field from the "Fields to complete" list in your JSON response. Do NOT skip any field — if you cannot determine the value, include the field with value set to null.**

Return ONLY a JSON object (no markdown, no extra text):
{{"completions": {{"city": {{"value": "San Francisco", "conclusion": "...", "reasoning": "...", "references": ["..."]}}, "designer": {{"value": null, "conclusion": "No designer information found", "reasoning": "...", "references": []}} }}, "notes": "brief notes"}}"""


def _build_name_block(case_data):
    """Build a name block from all available name variants."""
    parts = []
    def _s(key):
        return (case_data.get(key) or "").strip() if isinstance(case_data, dict) else ""
    name_cn = _s("name")
    name_en = _s("name_en")
    name_official = _s("name_official")
    if name_cn:
        parts.append(f"Chinese name: {name_cn}")
    if name_en:
        parts.append(f"English name: {name_en}")
    if name_official:
        parts.append(f"Official language name: {name_official}")
    for field, label in [("common_names", "Common names"), ("other_names", "Other names")]:
        raw = case_data.get(field, "") if isinstance(case_data, dict) else ""
        if raw:
            if isinstance(raw, list):
                items = raw
            elif isinstance(raw, str) and raw.strip():
                try:
                    items = json.loads(raw)
                except (json.JSONDecodeError, TypeError):
                    items = [raw]
            else:
                continue
            if items:
                parts.append(f"{label}: {', '.join(str(i) for i in items)}")
    return "\n".join(parts) if parts else f"Building name: {case_data.get('name', '')}"


def auto_complete_with_gemini(case_data, api_key):
    """Direct Gemini API auto-complete — no browser needed.

    Uses Gemini's multilingual knowledge of overseas Chinese architecture
    to fill missing fields with confidence scores. Searches by all known
    name variants (Chinese, English, official, common, other names).
    """
    building_name = case_data.get("name", "")
    if not building_name:
        return {"error": "建筑名称为空，无法补全"}

    completable = ["city", "construction_year", "demolition_year",
                   "restoration_year", "building_status", "quantity",
                   "function", "designer", "coordinates",
                   "latitude", "longitude", "spread_method"]

    country = str(case_data.get("country", "")).strip() or "未知"
    name_block = _build_name_block(case_data)

    prompt = AUTO_COMPLETE_GEMINI_PROMPT.format(
        name_block=name_block,
        country=country,
        known_fields="（无）",
        missing_fields=", ".join(completable),
    )

    client, model = _channel_client("field_completion", api_key, timeout=120)
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=16000,
    )
    content = response.choices[0].message.content.strip()
    result = _parse_response(content)
    if isinstance(result, dict):
        return result
    return {"completions": {}, "notes": f"AI返回解析失败: {str(result)[:200]}"}


# ---- Gemini image search ----

GEMINI_IMAGE_SEARCH_PROMPT = """Search for images of this building. Use ALL name variants below.

{name_block}
Country: {country}

For each image found:
1. Search and retrieve the actual image
2. Return it as base64-encoded JPEG data
3. Include a brief Chinese description

Return a JSON array (no markdown, no extra text):
[{{"description": "中文描述", "base64": "BASE64_DATA"}}, ...]
Return exactly 5 images. Only output the JSON array."""


def search_images_with_gemini(name_block, country, api_key, count=10):
    """Use Gemini to search for building images, returning them as base64 data.

    Makes multiple API calls (each returns ~5 images) to get the requested count.
    Returns: list of (temp_file_path, temp_file_path, description) tuples.
    """
    import base64
    import tempfile

    if not name_block or not name_block.strip():
        return []

    all_images = []
    rounds = (count + 4) // 5  # ceil(count/5)
    client, model = _channel_client("field_completion", api_key, timeout=180)

    for rnd in range(rounds):
        remaining = count - len(all_images)
        if remaining <= 0:
            break
        ask = min(5, remaining)

        prompt = GEMINI_IMAGE_SEARCH_PROMPT.format(
            name_block=name_block,
            country=country or "未知",
        ).replace("Return exactly 5 images", f"Return exactly {ask} images")

        try:
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=16000,
            )
        except Exception:
            break

        content = response.choices[0].message.content.strip()
        result = _parse_response(content)
        if not isinstance(result, list):
            continue

        for item in result[:ask]:
            if not isinstance(item, dict):
                continue
            b64 = item.get("base64", "")
            desc = item.get("description", "").strip()
            if not b64:
                continue
            try:
                clean = b64.replace("\n", "").replace("\r", "").replace(" ", "")
                img_bytes = base64.b64decode(clean)
                if len(img_bytes) < 500:
                    continue
                tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
                tmp.write(img_bytes)
                tmp.close()
                all_images.append((tmp.name, tmp.name, desc))
            except Exception:
                continue

        if len(content) < 8000:
            break  # response was short, model can't find more

    return all_images[:count]


def auto_complete_fields(case_data, ds_key=None, gemini_key=None):
    """One-click complete: DS first, then Gemini validates/refines.

    Returns dict with completions keyed by field_name.
    """
    building_name = case_data.get("name", "")
    if not building_name:
        return {"error": "建筑名称为空，无法补全"}

    # Fields we can auto-complete
    completable = ["country", "city", "construction_year", "demolition_year",
                   "restoration_year", "building_status", "quantity",
                   "function", "designer", "coordinates",
                   "latitude", "longitude"]
    # Identify missing
    missing = []
    for f in completable:
        val = case_data.get(f)
        if val is None or str(val).strip() == "":
            missing.append(f)
    if not missing:
        return {"completions": {}, "notes": "所有字段已填写完整"}

    # Known fields for context
    known_lines = []
    for f in completable:
        val = case_data.get(f)
        if val is not None and str(val).strip() != "":
            known_lines.append(f"  {f}: {val}")
    known_block = "\n".join(known_lines) if known_lines else "（无）"

    prompt = AUTO_COMPLETE_PROMPT.format(
        building_name=building_name,
        known_fields=known_block,
        missing_fields=", ".join(missing),
    )

    # Step 1: DS first
    ds_base, ds_model, ds_cfg_key = config.get_llm("search_parse")
    ds_result = _call_yunwu(prompt, ds_key or ds_cfg_key, model=ds_model, base_url=ds_base,
                            temperature=0.3, max_tokens=4000, timeout=90)
    if not isinstance(ds_result, dict):
        ds_result = {"completions": {}, "notes": str(ds_result)}

    # Step 2: Gemini validates/refines
    refine_prompt = f"""请审核并完善以下建筑案例的补全结果。

建筑名称：{building_name}
缺失字段：{', '.join(missing)}

DS 补全结果：
{json.dumps(ds_result, ensure_ascii=False, indent=2)}

请检查并修正：
1. 核实每个字段的值是否准确
2. 修正错误的值
3. 补充 DS 遗漏的字段
4. 给出最终置信度

返回 JSON：{{"completions": {{"field_name": {{"value": "...", "confidence": 0.8, "conclusion": "...", "reasoning": "...", "references": ["..."]}}}}, "notes": "审核备注"}}
只返回 JSON，不要其他内容。"""

    gem_base, gem_model, gem_cfg_key = config.get_llm("field_completion")
    gemini_result = _call_yunwu(refine_prompt, gemini_key or gem_cfg_key, model=gem_model, base_url=gem_base,
                                temperature=0.3, max_tokens=4000, timeout=90)
    if not isinstance(gemini_result, dict):
        gemini_result = ds_result  # fallback to DS result

    # Merge: prefer Gemini, fallback to DS
    final = ds_result.get("completions", {})
    if isinstance(gemini_result.get("completions"), dict):
        final = gemini_result["completions"]
    notes = gemini_result.get("notes", "") if isinstance(gemini_result, dict) else ""

    # Also handle latitude/longitude from coordinates field
    if "coordinates" in final and final["coordinates"].get("value"):
        coord_val = final["coordinates"]["value"]
        parts = [p.strip() for p in str(coord_val).split(",") if p.strip()]
        if len(parts) >= 2:
            final["latitude"] = {"value": parts[0], "confidence": final["coordinates"].get("confidence", 0.5)}
            final["longitude"] = {"value": parts[1], "confidence": final["coordinates"].get("confidence", 0.5)}
            final.pop("coordinates", None)

    return {"completions": final, "notes": notes, "missing": missing,
            "ds_result": ds_result, "gemini_result": gemini_result}


def generate_summary(text, api_key, model=DEFAULT_MODEL):
    """Generate a 1-2 sentence summary using DeepSeek."""
    if not text or not text.strip():
        return None
    truncated = text[:4000]
    prompt = (
        "请用1-2句话总结以下学术文献的主要内容。"
        "用中文回答，不超过150字，直接给出总结，不要任何前缀或说明。\n\n"
        f"{truncated}"
    )
    try:
        result = _call_yunwu(prompt, api_key, model=model,
                             temperature=0.3, max_tokens=200)
        if isinstance(result, str):
            return result.strip()
        if isinstance(result, list) and result:
            return str(result[0]).strip()
        return str(result).strip()
    except Exception as e:
        return f"[摘要生成失败: {e}]"


def generate_summary_from_meta_description(description):
    """Use meta description as summary if it's substantive."""
    if description and len(description.strip()) >= 20:
        return description.strip()
    return None

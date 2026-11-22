"""
Workflow engine — user-definable multi-step operations triggered by voice/AI.

A workflow is a named function that takes params and executes a sequence of
actions across the app. Workflows are registered in WORKFLOW_REGISTRY and
can be called from ChatAssistant or the AI dispatcher.
"""

import json
import threading
import urllib.parse
import webbrowser

import config

from ui.keyword_gen import (
    _all_countries, _get_country_langs, COUNTRY_GROUPS,
    CORE_KEYWORDS, LANGUAGE_OPTIONS,
)


def _find_country(name):
    """Find a country record by Chinese name, English name, or partial match."""
    for c in _all_countries():
        if name in (c[0], c[1]) or name.lower() in c[1].lower():
            return c
    # Fuzzy match on Chinese name
    for c in _all_countries():
        if name in c[0] or c[0] in name:
            return c
    return None


def _resolve_language(lang_name, country_record):
    """Map user-spoken language to one of three options: 中文, English, 官方语言.
    Returns (language_code, display_name) or (None, None)."""
    if isinstance(country_record, str):
        country_record = _find_country(country_record)
        if not country_record:
            return None, None

    lang_name = lang_name.strip()
    lang_lower = lang_name.lower()

    # Chinese
    if lang_lower in ("中文", "chinese", "zh", "cn", "中", "汉语", "华语"):
        return "zh", "中文"

    # English
    if lang_lower in ("英文", "english", "en", "英", "英语"):
        return "en", "English"

    # Official language — use the country's pre-defined official language(s)
    # country_record[2] is e.g. "Japanese" or "Hindi, English"
    # Use the first official language for translation
    if lang_lower in ("官方语言", "官方", "official", "当地语言", "本地语言", "母语"):
        official = _get_country_langs(country_record)[0]
        return "official", official

    return None, None


def _translate_keyword(kw_zh, lang_code, lang_display):
    """Translate a Chinese keyword to the target language.
    Uses pre-built translations from KEYWORD_TRANSLATIONS."""
    if lang_code == "zh" or lang_display == "中文":
        return kw_zh

    # English: use CORE_KEYWORDS mapping
    if lang_code == "en" or lang_display == "English":
        for i, k in enumerate(CORE_KEYWORDS["zh"]):
            if k == kw_zh and i < len(CORE_KEYWORDS.get("en", [])):
                return CORE_KEYWORDS["en"][i]
        return kw_zh

    # Other languages: use pre-built KEYWORD_TRANSLATIONS
    try:
        from utils.keyword_translations import KEYWORD_TRANSLATIONS
        mapping = KEYWORD_TRANSLATIONS.get(lang_display, {})
        return mapping.get(kw_zh, kw_zh)
    except ImportError:
        return kw_zh


def _translate_country(country_record, lang_code, lang_display):
    """Translate a country name to the target language."""
    if lang_code == "zh" or lang_display == "中文":
        return country_record[0]
    if lang_code == "en" or lang_display == "English":
        return country_record[1]

    try:
        from utils.keyword_translations import KEYWORD_TRANSLATIONS
        mapping = KEYWORD_TRANSLATIONS.get(lang_display, {})
        return mapping.get(country_record[0], country_record[1])
    except ImportError:
        return country_record[1]


def _get_case_keywords(country_name):
    """Get additional keywords from raw_cases for this country."""
    try:
        import database as db
        conn = db.get_connection()
        rows = conn.execute(
            "SELECT search_keywords FROM raw_cases WHERE country=? AND search_keywords IS NOT NULL AND search_keywords != ''",
            (country_name,)
        ).fetchall()
        conn.close()

        keywords = set()
        for r in rows:
            for kw in r[0].split(","):
                kw = kw.strip()
                if kw and len(kw) > 1:
                    keywords.add(kw)
        return list(keywords)
    except Exception:
        return []


def execute_one_click_search(country, language, engines=None, chat=None):
    """Execute the '一键搜索' workflow.

    Args:
        country: Country name (Chinese or English)
        language: Language name (中文, English, or official language name)
        engines: List of engine names to use. None = all engines.
        chat: ChatAssistant instance for status updates (optional).

    Returns:
        dict with 'count', 'queries', 'engines' keys.
    """
    # 1. Find country
    country_record = _find_country(country)
    if not country_record:
        if chat:
            chat._add_message("assistant", f"未找到国家「{country}」，请检查国家名称。")
        return None

    # 2. Resolve language
    lang_code, lang_display = _resolve_language(language, country_record)
    if not lang_code:
        if chat:
            chat._add_message("assistant",
                f"未识别语言「{language}」。支持：中文、English，或该国家的官方语言。")
        return None

    if chat:
        chat.set_status(f"一键搜索: {country_record[0]} ({lang_display})")

    # 3. Translate country name
    country_translated = _translate_country(country_record, lang_code, lang_display)

    # 4. Get core keywords and translate them
    core_kw_zh = list(CORE_KEYWORDS["zh"])
    search_queries = []
    for kw_zh in core_kw_zh:
        kw_translated = _translate_keyword(kw_zh, lang_code, lang_display)
        query = f"{country_translated} {kw_translated}"
        search_queries.append(query)

    # 5. Add case-specific keywords from database
    case_kws = _get_case_keywords(country_record[0])
    for ck in case_kws:
        if ck not in search_queries:
            search_queries.append(ck)

    # 6. Determine search engines based on language
    from ui.search_tab import ALL_ENGINES
    if engines:
        selected_engines = [e for e in engines if e in ALL_ENGINES]
    elif lang_code == "zh":
        selected_engines = ["百度"]
    else:
        selected_engines = ["Google"]

    if not selected_engines:
        selected_engines = ["Google"]

    # 7. Open Chrome tabs
    try:
        chrome = webbrowser.get('chrome')
    except webbrowser.Error:
        chrome = None

    total_tabs = 0
    open_errors = []
    for query in search_queries:
        encoded = urllib.parse.quote(query)
        for name in selected_engines:
            url = ALL_ENGINES[name].format(encoded)
            try:
                if chrome:
                    chrome.open(url)
                else:
                    webbrowser.open(url)
                total_tabs += 1
            except Exception as e:
                open_errors.append(f"{name}: {e}")

    result = {
        "count": total_tabs,
        "queries": search_queries,
        "engines": selected_engines,
        "country": country_translated,
        "language": lang_display,
    }

    if chat:
        msg = (
            f"一键搜索完成！\n"
            f"国家：{result['country']}\n"
            f"语言：{result['language']}\n"
            f"检索词：{len(search_queries)} 个\n"
            f"搜索引擎：{len(selected_engines)} 个\n"
            f"共打开 {total_tabs} 个 Chrome 标签页"
        )
        if open_errors:
            msg += f"\n\n打开失败的标签页({len(open_errors)}个):\n" + "\n".join(open_errors[:5])
        chat._add_message("assistant", msg)
        chat.set_status("")

    return result


# ==================================================================
#  Interactive Workflow Session
# ==================================================================

class WorkflowSession:
    """Holds the state of an in-progress interactive workflow.

    The session intercepts user input until the workflow completes
    or is cancelled. Each step returns a question for the user or
    None to auto-advance.
    """
    def __init__(self, name, chat, mw):
        self.name = name
        self.chat = chat
        self.mw = mw
        self.step = 0
        self.context = {}
        self.active = True
        self._error_state = False

    def next(self, user_input=None):
        """Advance to next step. user_input is the user's response
        to the previous question (or None for initial step)."""
        if not self.active:
            return

        # If in error recovery mode, handle the user's choice
        if self._error_state:
            self._handle_error_recovery(user_input)
            return

        try:
            fn = getattr(self, f"_step_{self.step}", None)
            if fn is None:
                self._finish()
                return
            fn(user_input)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.chat._add_message("error", f"步骤{self.step}出错: {e}")
            self._ask_error_recovery()

    def _ask_error_recovery(self):
        """Ask user how to handle an error."""
        self._error_state = True
        self._ask(f"步骤{self.step}出错，请选择：\n— 重试 / 跳过 / 结束")

    def _handle_error_recovery(self, user_input):
        """Process user's error recovery choice."""
        self._error_state = False
        if not user_input:
            user_input = ""
        if any(w in user_input for w in ["重试", "重新", "再来", "retry", "再试"]):
            self.chat._add_message("assistant", f"正在重试步骤{self.step}...")
            self.next()
        elif any(w in user_input for w in ["跳过", "skip", "继续"]):
            self.chat._add_message("assistant", f"已跳过步骤{self.step}，继续下一步。")
            self._advance()
        else:
            self.chat._add_message("assistant", f"工作流「{self.name}」已结束。")
            self.chat.set_status("")
            self.active = False

    def _ask(self, question):
        """Pause and ask the user a question."""
        self.chat._add_message("assistant", question)
        self.chat.set_status(f"[{self.name}] 等待你的回复...")

    def _advance(self):
        """Auto-advance to next step."""
        self.step += 1
        self.mw.root.after(300, lambda: self.next())

    def _run_async(self, fn, on_done):
        """Run fn in a background thread, then call on_done(result) on main thread.
        If fn raises, the error is passed to on_done as the result and the step
        must handle it."""
        def worker():
            try:
                result = fn()
            except Exception as e:
                result = e
            self.mw.root.after(0, lambda: on_done(result))
        threading.Thread(target=worker, daemon=True).start()

    def _finish(self):
        self.chat._add_message("assistant", f"工作流「{self.name}」已完成。")
        self.chat.set_status("")
        self.active = False


# ==================================================================
# API Keys & Shared Utilities
# ==================================================================

DS_BASE_URL, DS_MODEL, DS_API_KEY = config.get_llm("verify_assist")
G_BASE_URL, G_MODEL, G_API_KEY = config.get_llm("field_completion")

from utils.ai_extractor import (
    verify_field_deepseek, verify_field_gemini,
    auto_complete_with_gemini, _call_yunwu,
    AUTO_COMPLETE_PROMPT, _build_name_block,
)


def _call_ds(prompt, chat=None):
    """Call DeepSeek API and return the response text."""
    import openai
    if chat:
        chat.set_status("DS API 调用中...")
    client = openai.OpenAI(api_key=DS_API_KEY, base_url=DS_BASE_URL)
    resp = client.chat.completions.create(
        model=DS_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=4000,
    )
    return resp.choices[0].message.content


def _call_gemini(prompt, chat=None):
    """Call Gemini API via Yunwu and return the response text."""
    import openai
    if chat:
        chat.set_status("Gemini API 调用中...")
    client = openai.OpenAI(
        base_url=G_BASE_URL,
        api_key=G_API_KEY,
    )
    resp = client.chat.completions.create(
        model=G_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=16000,
    )
    return resp.choices[0].message.content


# ==================================================================
#  Workflow: 一键DS核验
# ==================================================================

class DsVerifySession(WorkflowSession):
    """Interactive session for 一键DS核验."""

    def _step_0(self, user_input):
        """List cases on first call; process user selection on subsequent calls."""
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        tree = vp.queue_tree
        children = tree.get_children()

        if not children:
            self.chat._add_message("assistant", "核验队列为空，请先在原始数据页提交案例到核验队列。")
            self.chat.set_status("")
            self.active = False
            return

        cases = []
        for item in children:
            vals = tree.item(item, "values")
            if vals:
                cases.append((vals[0], vals[1] if len(vals) > 1 else "?"))

        self.context["cases"] = cases

        if user_input and user_input.strip():
            target = None
            try:
                idx = int(user_input.strip()) - 1
                if 0 <= idx < len(cases):
                    target = cases[idx]
            except ValueError:
                for cid, name in cases:
                    if user_input.strip().lower() in name.lower():
                        target = (cid, name)
                        break

            if target:
                for item in children:
                    vals = tree.item(item, "values")
                    if vals and str(vals[0]) == str(target[0]):
                        tree.selection_set(item)
                        tree.see(item)
                        vp._on_select(None)
                        break

                case = vp._current_case
                if not case:
                    self.chat._add_message("error", "无法加载案例数据。")
                    self.chat.set_status("")
                    self.active = False
                    return

                self.context["case"] = case
                name = case.get("name", "")
                country = case.get("country", "未知")
                self.chat._add_message("assistant", f"已选中：{name} ({country})")
                self._advance()
                return

            case_list = "\n".join([f"  {i+1}. {name} (ID:{cid})" for i, (cid, name) in enumerate(cases)])
            self._ask(f"未找到案例「{user_input}」，请重新选择：\n{case_list}")
            return

        case_list = "\n".join([f"  {i+1}. {name} (ID:{cid})" for i, (cid, name) in enumerate(cases)])
        self._ask(f"请选择要核验的案例（说序号或名称）：\n{case_list}")

    def _step_1(self, _user_input):
        """DS name verification — same logic as UI DS button."""
        case = self.context.get("case", {})
        name = case.get("name", "")

        self.chat.set_status("DS核验名称中...")
        self.chat._add_message("assistant", "正在调用DeepSeek核验建筑名称...")
        self.mw.root.update()

        def do_api():
            return verify_field_deepseek(
                "name", "建筑名称", str(name), case, DS_API_KEY
            )

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"DS API调用失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["name_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== DS核验 名称核验结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            vp._last_paste_label = "DS核验"
            vp._paste_sources.clear()
            vp._paste_sources.add("DS核验")
            vp._save_paste_content("DS核验")

            try:
                import database as db
                db.insert_verification_log(
                    case.get("id"), "field_verify_deepseek",
                    "DeepSeek核验: 建筑名称",
                    json.dumps(result, ensure_ascii=False))
            except Exception:
                pass

            self.chat._add_message("assistant", "DS已返回名称核验结果，正在分析...")
            self._advance()

        self._run_async(do_api, on_done)

    def _step_2(self, _user_input):
        """Run name field analysis and save results."""
        vp = self.mw.organize_tab.verify_panel
        try:
            vp._run_name_field_analysis()
            self.chat._add_message("assistant", "名称字段分析完成，结果已保存。\n正在调用DS一键补全其他字段...")
        except Exception as e:
            self.chat._add_message("error", f"名称分析失败: {e}")
            raise
        case_id = self.context.get("case", {}).get("id")
        if case_id:
            import database as db
            self.context["case"] = db.get_raw_case_by_id(case_id) or vp._current_case
        self._advance()

    def _step_3(self, _user_input):
        """DS auto-complete — same logic as UI 一键DS补全 button."""
        case = self.context.get("case", {})

        self.chat.set_status("DS一键补全中...")
        self.chat._add_message("assistant", "正在调用DeepSeek一键补全所有字段...")
        self.mw.root.update()

        name_block = _build_name_block(case)
        completable = ["city", "construction_year", "demolition_year",
                       "restoration_year", "building_status", "quantity",
                       "function", "designer", "coordinates",
                       "latitude", "longitude", "spread_method"]
        country = str(case.get("country", "")).strip() or "未知"
        known_block = f"country: {country}"

        prompt = AUTO_COMPLETE_PROMPT.format(
            building_name=name_block,
            known_fields=known_block,
            missing_fields=", ".join(completable),
        )

        def do_api():
            return _call_yunwu(prompt, DS_API_KEY, model="deepseek-chat",
                               temperature=0.3, max_tokens=8000, timeout=90)

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"DS API调用失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["complete_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== DeepSeek一键补全结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            self.chat._add_message("assistant", "DS已返回补全结果，正在填充字段...")
            self._advance()

        self._run_async(do_api, on_done)

    def _step_4(self, _user_input):
        """Run 其他字段信息分析 to fill fields."""
        vp = self.mw.organize_tab.verify_panel
        try:
            vp._parse_and_fill_from_text()
            self.chat._add_message("assistant", "其他字段信息分析完成。\n正在保存结果...")
        except Exception as e:
            self.chat._add_message("error", f"其他字段信息分析失败: {e}")
            raise

        self._advance()

    def _step_5(self, _user_input):
        """Proceed to image search."""
        self.chat._add_message("assistant", "结果已保存。\n正在打开浏览器搜索图片...")
        self._advance()

    def _build_image_query(self):
        """Build image search query from case name + country + city."""
        case = self.context.get("case", {})
        parts = []
        for key in ["name", "name_en", "name_official"]:
            v = (case.get(key) or "").strip()
            if v:
                parts.append(v)
        for key in ["country", "city"]:
            v = (case.get(key) or "").strip()
            if v:
                parts.append(v)
        return " ".join(parts)

    def _open_image_search(self, engine="baidu"):
        """Open image search in Chrome browser."""
        query = self._build_image_query()
        if engine == "google":
            url = f"https://www.google.com/search?tbm=isch&q={urllib.parse.quote(query)}"
        else:
            query_encoded = urllib.parse.quote(f"{query} 建筑")
            url = f"https://image.baidu.com/search?word={query_encoded}"
        try:
            webbrowser.get('chrome').open(url)
        except webbrowser.Error:
            webbrowser.open(url)

    def _step_6(self, _user_input):
        """Open Baidu image search."""
        self.chat._add_message("assistant",
            "正在打开百度图片搜索...\n"
            "请在浏览器中查看并下载合适的图片，\n"
            "然后回到核验面板通过「上传图片」按钮添加到案例中。")
        self._open_image_search("baidu")
        self._ask("是否找到满意的图片？\n— 满意 / 不满意（换谷歌搜图）")
        self.step = 7

    def _step_7(self, user_input):
        """Handle Baidu image satisfaction check."""
        if user_input and any(w in user_input for w in ["不满意", "没有", "谷歌", "google", "否", "不行", "不好"]):
            self.chat._add_message("assistant", "正在打开谷歌图片搜索...")
            self._open_image_search("google")
            self._ask("是否找到满意的图片？\n— 满意 / 不满意（返回百度搜）")
            self.step = 8
            return
        if user_input and any(w in user_input for w in ["再加", "继续加", "再加一张", "百度"]):
            self._open_image_search("baidu")
            self._ask("是否找到满意的图片？\n— 满意 / 不满意（换谷歌搜图）")
            return
        self.chat._add_message("assistant", "一键DS核验完成！")
        self.chat.set_status("")
        self.active = False

    def _step_8(self, user_input):
        """Handle Google image satisfaction check."""
        if user_input and any(w in user_input for w in ["不满意", "没有", "百度", "不行", "不好"]):
            self.chat._add_message("assistant", "正在重新使用百度搜索图片...")
            self._open_image_search("baidu")
            self._ask("是否找到满意的图片？\n— 满意 / 不满意（换谷歌搜图）")
            self.step = 7
            return
        if user_input and any(w in user_input for w in ["再加", "继续加", "再加一张", "谷歌"]):
            self._open_image_search("google")
            self._ask("是否找到满意的图片？\n— 满意 / 不满意（返回百度搜）")
            return
        self.chat._add_message("assistant", "一键DS核验完成！")
        self.chat.set_status("")
        self.active = False


def execute_ds_verify(chat=None, mw=None):
    """Launch the interactive DS核验 workflow."""
    session = DsVerifySession("一键DS核验", chat, mw)
    session.next()
    return session


# ==================================================================
#  Workflow: 一键G核验
# ==================================================================

class GVerifySession(DsVerifySession):
    """Interactive session for 一键G核验 — same flow as DS but uses Gemini API."""

    def _step_1(self, _user_input):
        """Gemini name verification — same logic as UI G button."""
        case = self.context.get("case", {})
        name = case.get("name", "")

        self.chat.set_status("Gemini核验名称中...")
        self.chat._add_message("assistant", "正在调用Gemini核验建筑名称...")
        self.mw.root.update()

        def do_api():
            return verify_field_gemini(
                "name", "建筑名称", str(name), case, G_API_KEY
            )

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"Gemini API调用失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["name_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== Gemini核验 名称核验结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            vp._last_paste_label = "G核验"
            vp._paste_sources.clear()
            vp._paste_sources.add("G核验")
            vp._save_paste_content("G核验")

            try:
                import database as db
                db.insert_verification_log(
                    case.get("id"), "field_verify_gemini",
                    "Gemini核验: 建筑名称",
                    json.dumps(result, ensure_ascii=False))
            except Exception:
                pass

            self.chat._add_message("assistant", "Gemini已返回名称核验结果，正在分析...")
            self._advance()

        self._run_async(do_api, on_done)

    def _step_3(self, _user_input):
        """Gemini auto-complete — same logic as UI 一键G补全 button."""
        case = self.context.get("case", {})

        self.chat.set_status("Gemini一键补全中...")
        self.chat._add_message("assistant", "正在调用Gemini一键补全所有字段...")
        self.mw.root.update()

        def do_api():
            return auto_complete_with_gemini(case, G_API_KEY)

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"Gemini API调用失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["complete_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== Gemini一键补全结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            self.chat._add_message("assistant", "Gemini已返回补全结果，正在填充字段...")
            self._advance()

        self._run_async(do_api, on_done)


def execute_g_verify(chat=None, mw=None):
    """Launch the interactive G核验 workflow."""
    session = GVerifySession("一键G核验", chat, mw)
    session.next()
    return session


# ==================================================================
#  Workflow: 一键核验 (DS first, G fallback if not satisfied)
# ==================================================================

class VerifySession(DsVerifySession):
    """一键核验工作流。

    第一步: 选中案例
    第二步: DS核验名称 → 名称字段分析 → DS一键补全 → 其他字段分析 → 问满意
           满意 → 百度搜图
           不满意 → G核验名称 → 名称字段分析 → G一键补全 → 其他字段分析 → 百度搜图
    第三步: 百度搜图 → 问满意 → 不满意 → 谷歌搜图 → 结束
    """

    # ==================================================================
    # 第一步: 选中案例 (继承 DsVerifySession._step_0)
    # ==================================================================

    # ==================================================================
    # 第二步a: DS核验名称 → 名称字段分析 → 进入 Step 2
    # ==================================================================

    def _step_1(self, _user_input):
        case = self.context.get("case", {})
        name = case.get("name", "")

        self.chat.set_status("DS核验名称中...")
        self.chat._add_message("assistant", "正在调用DeepSeek核验建筑名称...")
        self.mw.root.update()

        def do_api():
            return verify_field_deepseek(
                "name", "建筑名称", str(name), case, DS_API_KEY
            )

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"DS API调用失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["name_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== DS核验 名称核验结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            vp._last_paste_label = "DS核验"
            vp._paste_sources.clear()
            vp._paste_sources.add("DS核验")
            vp._save_paste_content("DS核验")

            try:
                import database as db
                db.insert_verification_log(
                    case.get("id"), "field_verify_deepseek",
                    "DeepSeek核验: 建筑名称",
                    json.dumps(result, ensure_ascii=False))
            except Exception:
                pass

            # 名称字段信息分析
            try:
                vp._run_name_field_analysis()
                self.chat._add_message("assistant", "名称字段分析完成，结果已保存。\n正在调用DS一键补全...")
            except Exception as e:
                self.chat._add_message("error", f"名称分析失败: {e}")
                self._ask_error_recovery()
                return

            case_id = case.get("id")
            if case_id:
                import database as db
                self.context["case"] = db.get_raw_case_by_id(case_id) or vp._current_case

            self.step = 2
            self.mw.root.after(300, lambda: self.next())

        self._run_async(do_api, on_done)

    # ==================================================================
    # 第二步b: DS一键补全 → 其他字段分析 → 问满意 → Step 3
    # ==================================================================

    def _step_2(self, _user_input):
        case = self.context.get("case", {})

        self.chat.set_status("DS一键补全中...")
        self.chat._add_message("assistant", "正在调用DeepSeek一键补全所有字段...")
        self.mw.root.update()

        name_block = _build_name_block(case)
        completable = ["city", "construction_year", "demolition_year",
                       "restoration_year", "building_status", "quantity",
                       "function", "designer", "coordinates",
                       "latitude", "longitude", "spread_method"]
        country = str(case.get("country", "")).strip() or "未知"
        known_block = f"country: {country}"

        prompt = AUTO_COMPLETE_PROMPT.format(
            building_name=name_block,
            known_fields=known_block,
            missing_fields=", ".join(completable),
        )

        def do_api():
            return _call_yunwu(prompt, DS_API_KEY, model="deepseek-chat",
                               temperature=0.3, max_tokens=8000, timeout=90)

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"DS API调用失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["complete_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== DeepSeek一键补全结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            # 其他字段信息分析
            try:
                vp._parse_and_fill_from_text()
                self.chat._add_message("assistant", "其他字段信息分析完成，结果已保存。")
            except Exception as e:
                self.chat._add_message("error", f"其他字段信息分析失败: {e}")
                self._ask_error_recovery()
                return

            self.step = 3
            self._ask("DS补全是否满意？\n— 满意（百度搜图） / 不满意（换G核验）")

        self._run_async(do_api, on_done)

    # ==================================================================
    # 第二步c: 处理满意度 → 满意→Step4(百度搜图) / 不满意→Step5(G核验)
    # ==================================================================

    def _step_3(self, user_input):
        if user_input and any(w in user_input for w in ["不满意", "没有", "否", "不行", "不好", "g核验", "G核验", "换"]):
            self.chat._add_message("assistant", "正在切换到Gemini核验...")
            self.step = 5
            self.mw.root.after(300, lambda: self.next())
            return

        # 满意 → 百度搜图
        self.chat._add_message("assistant", "正在打开百度图片搜索...")
        self.step = 4
        self.mw.root.after(300, lambda: self.next())

    # ==================================================================
    # 第二步d(不满意分支): G核验名称 → 名称字段分析 → Step 6
    # ==================================================================

    def _step_5(self, _user_input):
        case = self.context.get("case", {})
        name = case.get("name", "")

        self.chat.set_status("G核验名称中...")
        self.chat._add_message("assistant", "正在调用Gemini核验建筑名称...")
        self.mw.root.update()

        def do_api():
            return verify_field_gemini(
                "name", "建筑名称", str(name), case, G_API_KEY
            )

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"Gemini名称核验失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["name_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== Gemini核验 名称核验结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            vp._last_paste_label = "G核验"
            vp._paste_sources.clear()
            vp._paste_sources.add("G核验")
            vp._save_paste_content("G核验")

            try:
                import database as db
                db.insert_verification_log(
                    case.get("id"), "field_verify_gemini",
                    "Gemini核验: 建筑名称",
                    json.dumps(result, ensure_ascii=False))
            except Exception:
                pass

            # 名称字段信息分析
            try:
                vp._run_name_field_analysis()
                self.chat._add_message("assistant", "G名称核验完成，结果已保存。\n正在调用Gemini一键补全...")
            except Exception as e:
                self.chat._add_message("error", f"名称分析失败: {e}")
                self._ask_error_recovery()
                return

            case_id = case.get("id")
            if case_id:
                import database as db
                self.context["case"] = db.get_raw_case_by_id(case_id) or vp._current_case

            self.step = 6
            self.mw.root.after(300, lambda: self.next())

        self._run_async(do_api, on_done)

    # ==================================================================
    # 第二步e(不满意分支): G一键补全 → 其他字段分析 → 百度搜图(Step 4)
    # ==================================================================

    def _step_6(self, _user_input):
        case = self.context.get("case", {})

        self.chat.set_status("Gemini一键补全中...")
        self.chat._add_message("assistant", "正在调用Gemini一键补全所有字段...")
        self.mw.root.update()

        def do_api():
            return auto_complete_with_gemini(case, G_API_KEY)

        def on_done(result):
            if isinstance(result, Exception):
                self.chat._add_message("error", f"Gemini补全失败: {result}")
                self._ask_error_recovery()
                return

            vp = self.mw.organize_tab.verify_panel
            self.context["complete_result"] = result

            vp.paste_text.configure(state="normal")
            vp.paste_text.delete("1.0", "end")
            vp.paste_text.insert("1.0", "=== Gemini一键补全结果 ===\n")
            vp.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
            vp.paste_text.configure(state="disabled")

            # 其他字段信息分析
            try:
                vp._parse_and_fill_from_text()
                self.chat._add_message("assistant", "G一键补全完成，其他字段分析已保存。")
            except Exception as e:
                self.chat._add_message("error", f"其他字段信息分析失败: {e}")
                self._ask_error_recovery()
                return

            self.chat._add_message("assistant", "正在打开百度图片搜索...")
            self.step = 4
            self.mw.root.after(300, lambda: self.next())

        self._run_async(do_api, on_done)

    # ==================================================================
    # 第三步a: 百度搜图 → 问满意 → Step 7
    # ==================================================================

    def _step_4(self, _user_input):
        self.chat._add_message("assistant",
            "正在打开百度图片搜索...\n"
            "请在浏览器中查看并下载合适的图片，\n"
            "然后回到核验面板通过「上传图片」按钮添加到案例中。")
        self._open_image_search("baidu")
        self.step = 7
        self._ask("百度图片是否满意？\n— 满意 / 不满意（换谷歌搜图）")

    # ==================================================================
    # 第三步b: 处理图片满意度 → 满意→结束 / 不满意→Step 8(谷歌搜图)
    # ==================================================================

    def _step_7(self, user_input):
        if user_input and any(w in user_input for w in ["不满意", "没有", "谷歌", "google", "否", "不行", "不好"]):
            self.chat._add_message("assistant", "正在打开谷歌图片搜索...")
            self.step = 8
            self.mw.root.after(300, lambda: self.next())
            return
        self.chat._add_message("assistant", "一键核验完成！")
        self.chat.set_status("")
        self.active = False

    # ==================================================================
    # 第三步c: 谷歌搜图 → 结束
    # ==================================================================

    def _step_8(self, _user_input):
        self.chat._add_message("assistant",
            "正在打开谷歌图片搜索...\n"
            "请在浏览器中查看并下载合适的图片。")
        self._open_image_search("google")
        self.chat._add_message("assistant", "一键核验完成！")
        self.chat.set_status("")
        self.active = False


def execute_verify(chat=None, mw=None):
    """Launch the interactive 一键核验 workflow."""
    session = VerifySession("一键核验", chat, mw)
    session.next()
    return session


# Registry of all available workflows
WORKFLOW_REGISTRY = {
    "一键搜索": {
        "description": "对国家+语言的所有关键词进行一键搜索",
        "params": ["country", "language"],
        "function": execute_one_click_search,
        "is_interactive": False,
        "examples": [
            "一键搜索 日本 中文",
            "一键搜索 印度 英文",
            "一键搜索 缅甸 English",
        ],
    },
    "一键DS核验": {
        "description": "用DeepSeek对核验队列中的案例进行名称核验+字段补全+图片保存",
        "params": [],
        "function": execute_ds_verify,
        "is_interactive": True,
        "examples": [
            "一键DS核验",
            "ds核验",
        ],
    },
    "一键G核验": {
        "description": "用Gemini对核验队列中的案例进行名称核验+字段补全+图片保存",
        "params": [],
        "function": execute_g_verify,
        "is_interactive": True,
        "examples": [
            "一键G核验",
            "g核验",
        ],
    },
    "一键核验": {
        "description": "DS优先，不满意换G，对核验队列案例进行名称核验+字段补全+图片保存",
        "params": [],
        "function": execute_verify,
        "is_interactive": True,
        "examples": [
            "一键核验",
            "核验",
        ],
    },
}


def execute_workflow(name, params, chat=None, mw=None):
    """Execute a named workflow with params. Returns the workflow's result or None."""
    wf = WORKFLOW_REGISTRY.get(name)
    if not wf:
        if chat:
            chat._add_message("error", f"未找到工作流「{name}」。可用：{list(WORKFLOW_REGISTRY.keys())}")
        return None

    fn = wf["function"]
    try:
        if wf.get("is_interactive"):
            return fn(chat=chat, mw=mw)
        return fn(**params, chat=chat)
    except Exception as e:
        if chat:
            chat._add_message("error", f"工作流 {name} 执行失败: {e}")
        return None

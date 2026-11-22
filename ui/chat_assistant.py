"""
Chat-based assistant panel that automates manual UI operations.
Supports regex shortcuts (instant), AI-powered natural language commands, and voice input.
"""

import json
import os
import re
import tempfile
import threading
import tkinter as tk
from tkinter import ttk
import wave

import openai
import sounddevice as sd
import config

# Vosk model path
_VOSK_MODEL_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "data", "vosk-model", "vosk-model-small-cn"
)
_vosk_model = None

def _get_vosk_model():
    """Lazy-load Vosk model if available."""
    global _vosk_model
    if _vosk_model is not None:
        return _vosk_model
    if os.path.isdir(_VOSK_MODEL_DIR):
        try:
            from vosk import Model
            _vosk_model = Model(_VOSK_MODEL_DIR)
            return _vosk_model
        except Exception:
            pass
    return None

API_BASE_URL, DEFAULT_MODEL, _CHAT_API_KEY = config.get_llm("verify_assist")

SYSTEM_PROMPT = """你是海外中式建筑数据库的全功能操作助手。根据用户的自然语言指令，返回对应的操作JSON。

返回格式：{"action": "<动作名>", "params": {"key": "value"}}

=== 支持的全部动作 ===

【导航】
- switch_tab: 切换主标签页。params: {"tab": "search|organize|analysis"}
- switch_subtab: 切换子标签页。params: {"subtab": "raw|verify|database"}
- next_case: 选择下一个案例
- prev_case: 选择上一个案例

【案例查询页 - 网页搜索】
- web_search: 用关键词在搜索引擎中搜索。params: {"keyword": "搜索词"}
- generate_keywords: 生成搜索关键词

【案例查询页 - 库内搜索】
- local_search: 库内搜索。params: {"keyword": "词", "type": "建筑类型", "country": "国家"}
- clear_filter: 清除筛选条件
- show_detail: 查看选中案例详情
- delete_selected: 删除选中案例
- select_all: 全选
- clear_selection: 取消全选

【案例查询页 - AI提取】
- ai_extract: 开始AI提取案例
- ai_load_file: 从文件加载内容
- ai_export: 导出AI提取结果到Excel
- ai_add_to_db: 将AI提取结果添加到数据库
- ai_clear: 清空AI提取结果

【数据整理-原始数据】
- submit_verify: 提交选中项到核验队列
- delete_raw: 删除选中的原始数据
- refresh_raw: 刷新原始数据列表

【数据整理-真实性核验】
- refresh_queue: 刷新核验队列
- open_verify_detail: 打开核验详情
- enable_field_edit: 进入字段编辑模式
- save_field_changes: 保存字段修改
- auto_complete_all: 一键补全所有字段
- approve: 核验通过
- reject: 核验拒绝
- apply_field_result: 应用当前字段的核验结果
- save_manual_result: 保存手动核验结果
- clear_result: 清空核验结果
- clear_paste: 清空粘贴内容
- name_field_analysis: 分析名称字段信息
- other_field_analysis: 分析其他字段信息
- download_image: 下载当前案例图片
- undo_delete_image: 撤销删除图片
- refresh_image: 刷新图片预览
- search_field: 搜索/核验某字段。params: {"field": "字段名", "engine": "baidu|google|ds|gemini"}
- complete_field: 补全某字段。params: {"field": "字段名"}
- verify_field: 核验某字段。params: {"field": "字段名"}
- search_active_field: 用搜索引擎查当前活跃字段。params: {"engine": "baidu|google|ds|gemini"}
- select_local_image: 选择本地图片
- delete_verify: 从核验队列删除选中项
- set_engines: 设置搜索引擎选择。params: {"group": "grey|research", "select_all": true/false}

【数据整理-正式案例】
- import_data: 导入数据
- export_data: 导出数据到Excel
- find_duplicates: 查找重复案例
- clean_data: 数据清洗
- check_missing: 缺失字段检查
- add_case: 新增案例
- edit_selected: 编辑选中案例
- delete_db: 删除选中案例（正式库）
- refresh_db: 刷新正式案例列表

【数据分析】
- chart_by: 按维度显示图表。params: {"dimension": "type|country|year|protection"}
- export_all: 导出全部数据到Excel
- export_stats: 按当前维度导出统计数据

【通用】
- refresh: 刷新当前页面
- help: 显示帮助
- workflow: 执行工作流。params: {"name": "一键搜索", "country": "国家名", "language": "语言"}
- unknown: 无法识别

用户说"百度搜索名称"→ action=search_field, params={"field":"name","engine":"baidu"}
用户说"一键搜索日本中文"→ action=workflow, params={"name":"一键搜索","country":"日本","language":"中文"}
用户说"删除这个"→ 根据当前所在页面选择delete_selected/delete_raw/delete_db
用户说"删除这个"→ 根据当前所在页面选择delete_selected/delete_raw/delete_db
只返回JSON，不要其他内容。"""

# ---------------------------------------------------------------------------
#  Shortcut patterns — matched locally, no API call
# ---------------------------------------------------------------------------

SHORTCUTS = [
    # (regex pattern, handler method name)
    (r"^(帮助|help|命令|你能做什么|你会什么|怎么用|使用说明)$", "_cmd_help"),
    (r"^(搜索|查找|搜|帮我搜|帮我查|查一下|搜一下)\s*(.*)", "_cmd_search"),
    (r"^打开核验", "_cmd_switch_verify"),
    (r"^(切换|跳转|去|打开)(搜索|案例查询)页?", "_cmd_switch_search"),
    (r"^(切换|跳转|去|打开)(数据整理|整理)页?", "_cmd_switch_organize"),
    (r"^(切换|跳转|去|打开)(数据分析|分析)页?", "_cmd_switch_analysis"),
    (r"^(切换|跳转|去|打开)原始数据", "_cmd_switch_raw"),
    (r"^(切换|跳转|去|打开)正式案例", "_cmd_switch_database"),
    (r"^(刷新|更新列表|更新一下|刷新一下)$", "_cmd_refresh"),
    (r"^(导入|导入数据|帮我导入)$", "_cmd_import"),
    (r"^(导出|导出数据|帮我导出)$", "_cmd_export"),
    (r"^(自动补全|一键补全|自动填充|帮我补全|帮我填)$", "_cmd_auto_complete"),
    (r"^(删除选中|删除|删掉|帮我删)$", "_cmd_delete"),
    (r"^(添加|新建|新增)(一个)?(案例)?$", "_cmd_add_case"),
    (r"^(查看详情|详情|打开详情)$", "_cmd_view_detail"),
    (r"^(下一个|下一条|下个|next|往后翻)$", "_cmd_next"),
    (r"^(上一个|上一条|上个|prev|往前翻)$", "_cmd_prev"),
    (r"^(提交核验|提交审核|送审)$", "_cmd_submit_verify"),
    (r"^(批准|通过|approve|同意)$", "_cmd_approve"),
    (r"^(拒绝|reject|驳回|不同意)$", "_cmd_reject"),
    (r"^(全选|全部选择|选全部)$", "_cmd_select_all"),
    (r"^(清空|取消全选)$", "_cmd_clear_selection"),
    (r"^(取消|停止|暂停|cancel|stop)$", "_cmd_cancel_workflow"),
    # Chitchat / greetings
    (r"^(你好|您好|嗨|hi|hello|谢谢|谢了|多谢|再见|拜拜|bye)$", "_cmd_chitchat"),
    # Workflow: 一键搜索
    (r"^一键搜索\s*(.*?)\s+(.*)", "_cmd_workflow_search"),
    # Workflow: 一键DS核验
    (r"^(一键DS核验|ds核验|DS核验|ds验证|DS验证)$", "_cmd_ds_verify"),
    # Workflow: 一键G核验
    (r"^(一键G核验|g核验|G核验|g验证|G验证)$", "_cmd_g_verify"),
    # Workflow: 一键核验 (DS→G loop)
    (r"^(一键核验|核验)$", "_cmd_verify"),
    # Catch-all natural speech patterns for search
    (r".*(帮我|我想|我要|想)(搜|搜索|查|找)(一下|一查)?\s*(.+)", "_cmd_search_natural"),
    (r".*(帮我|我想|我要|让我)(看看|看一下|瞧瞧).*", "_cmd_view_detail"),
]


# ---------------------------------------------------------------------------
#  VoiceInput — microphone recording + speech-to-text
# ---------------------------------------------------------------------------

class VoiceInput:
    """Record audio from the default microphone and transcribe via speech API."""

    def __init__(self, chat_ui):
        self.chat = chat_ui
        self._recording = False

    def start(self):
        """Start recording in a background thread."""
        if self._recording:
            return
        self._recording = True
        self.chat.set_status("🔴 录音中... 请说话 (3秒)")
        self.chat._add_message("system", "[语音识别中...]")
        threading.Thread(target=self._record_and_transcribe, daemon=True).start()

    def _record_and_transcribe(self):
        fs = 16000  # sample rate
        duration = 3  # seconds — shorter for less lag
        tmp_path = None
        try:
            # Record
            audio_data = sd.rec(int(duration * fs), samplerate=fs, channels=1, dtype="int16")
            sd.wait()

            # Save to temp WAV
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".wav")
            os.close(tmp_fd)
            with wave.open(tmp_path, "wb") as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(fs)
                wf.writeframes(audio_data.flatten().tobytes())

            # Offline recognition via Vosk
            self.chat.main_window.root.after(0,
                lambda: self.chat.set_status("识别中..."))
            text = self._transcribe_offline(tmp_path)

            if text and text.strip():
                self.chat.main_window.root.after(0,
                    lambda t=text.strip(): self._on_result(t))
            else:
                self.chat.main_window.root.after(0, self._on_no_result)
        except Exception as e:
            self.chat.main_window.root.after(0,
                lambda: self._on_error(str(e)[:80]))
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
            self._recording = False

    def _transcribe_offline(self, wav_path):
        """Try Vosk offline recognition. Returns text or None."""
        model = _get_vosk_model()
        if model is None:
            return None
        try:
            from vosk import KaldiRecognizer
            rec = KaldiRecognizer(model, 16000)
            with wave.open(wav_path, "rb") as wf:
                while True:
                    data = wf.readframes(4000)
                    if len(data) == 0:
                        break
                    rec.AcceptWaveform(data)
            result = json.loads(rec.FinalResult())
            return result.get("text", "").strip()
        except Exception:
            return None

    def _on_result(self, text):
        self.chat._add_message("system", f"[语音识别] {text}")
        self.chat.input_var.set(text)
        self.chat.set_status("")
        self.chat._send()

    def _on_no_result(self):
        if not _get_vosk_model():
            self.chat._add_message("system", "[离线语音模型未安装，无法使用语音识别]")
        else:
            self.chat._add_message("system", "[未能识别语音，请重试]")
        self.chat.set_status("")

    def _on_error(self, err):
        self.chat._add_message("error", f"语音识别失败: {err}")
        self.chat.set_status("")


# ---------------------------------------------------------------------------
#  ChatAssistant — chat UI panel
# ---------------------------------------------------------------------------

class ChatAssistant(ttk.Frame):
    """Collapsible right-side chat panel."""

    def __init__(self, parent, main_window):
        super().__init__(parent)
        self.main_window = main_window
        self.dispatcher = CommandDispatcher(main_window, self)
        self.voice = VoiceInput(self)
        self._build_ui()

    def _build_ui(self):
        # Header
        header = ttk.Frame(self)
        header.pack(fill="x", padx=5, pady=(5, 0))
        ttk.Label(header, text="助手", font=("", 11, "bold")).pack(side="left")
        ttk.Button(header, text="清空", width=4, command=self.clear).pack(side="right")

        # Chat history
        hist_frame = ttk.Frame(self)
        hist_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.history = tk.Text(hist_frame, wrap="word", state="disabled",
                               font=("", 10), relief="flat", borderwidth=0,
                               padx=6, pady=6, bg="#f5f5f5", fg="#1a1a1a")
        scroll = ttk.Scrollbar(hist_frame, command=self.history.yview)
        self.history.configure(yscrollcommand=scroll.set)
        self.history.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Configure text tags
        self.history.tag_configure("user", foreground="#2563eb", font=("", 10, "bold"))
        self.history.tag_configure("assistant", foreground="#16a34a")
        self.history.tag_configure("error", foreground="#dc2626")
        self.history.tag_configure("system", foreground="#1a1a1a", font=("", 9))

        # Input area
        input_frame = ttk.Frame(self)
        input_frame.pack(fill="x", padx=5, pady=(0, 5))
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_var,
                                     font=("", 10))
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(0, 2))
        self.input_entry.bind("<Return>", lambda e: self._send())
        self.mic_btn = ttk.Button(input_frame, text="🎤", width=3,
                                  command=self.voice.start)
        self.mic_btn.pack(side="left", padx=(0, 2))
        self.stop_btn = ttk.Button(input_frame, text="停止", command=self._stop,
                                   width=4, style="Stop.TButton")
        self.send_btn = ttk.Button(input_frame, text="发送", command=self._send, width=6)
        self.send_btn.pack(side="right")

        # Style for stop button
        style = ttk.Style()
        style.configure("Stop.TButton", foreground="#dc2626")

        self._running = False

        # Status
        self.status_label = ttk.Label(self, text="", foreground="gray", font=("", 8))
        self.status_label.pack(padx=5, pady=(0, 5), anchor="w")

        # Welcome
        self._add_message("assistant",
                          "你好！我是软件助手。\n输入命令即可帮你操作，输入「帮助」查看可用命令。")

        self.input_entry.focus_set()

    def _send(self):
        text = self.input_var.get().strip()
        if not text:
            return
        self.input_var.set("")
        self._add_message("user", text)
        self._show_running()
        self.status_label.config(text="处理中...")
        self.dispatcher._cancelled = False
        threading.Thread(target=self.dispatcher.dispatch, args=(text,), daemon=True).start()

    def _show_running(self):
        self._running = True
        self.send_btn.pack_forget()
        self.stop_btn.pack(side="right")

    def _hide_stop(self):
        self._running = False
        self.stop_btn.pack_forget()
        self.send_btn.pack(side="right")

    def _stop(self):
        self.dispatcher._cancelled = True
        self._add_message("system", "[已请求停止]")
        self._hide_stop()
        self.set_status("已停止")

    def _add_message(self, role, text):
        def _update():
            self.history.configure(state="normal")
            if role == "user":
                self.history.insert("end", "你: ", "user")
                self.history.insert("end", text + "\n\n", "user")
            elif role == "assistant":
                self.history.insert("end", "助手: ", "assistant")
                self.history.insert("end", text + "\n\n", "assistant")
            elif role == "error":
                self.history.insert("end", text + "\n\n", "error")
            elif role == "system":
                self.history.insert("end", text + "\n", "system")
            self.history.configure(state="disabled")
            self.history.see("end")
        self.main_window.root.after(0, _update)

    def clear(self):
        def _update():
            self.history.configure(state="normal")
            self.history.delete("1.0", "end")
            self.history.configure(state="disabled")
        self.main_window.root.after(0, _update)

    def set_status(self, text):
        self.main_window.root.after(0, lambda t=text: self.status_label.config(text=t))


# ---------------------------------------------------------------------------
#  CommandDispatcher — shortcut matching + AI fallback
# ---------------------------------------------------------------------------

class CommandDispatcher:
    """Parses user input and executes the corresponding UI action."""

    def __init__(self, main_window, chat_ui):
        self.mw = main_window
        self.chat = chat_ui
        self.active_session = None
        self._cancelled = False

    def dispatch(self, text):
        # Allow cancel/stop to interrupt any active session
        if text.strip() in ("取消", "停止", "暂停", "cancel", "stop"):
            if self.active_session and self.active_session.active:
                self.active_session.active = False
                self.chat.set_status("")
                self.chat._add_message("assistant", "工作流已取消。")
                self.active_session = None
                return

        # If an interactive session is active, forward input to it
        if self.active_session and self.active_session.active:
            self.mw.root.after(0, lambda t=text: self.active_session.next(t))
            return

        # Try regex shortcuts first
        for pattern, handler_name in SHORTCUTS:
            m = re.match(pattern, text, re.IGNORECASE)
            if m:
                handler = getattr(self, handler_name, None)
                if handler:
                    self.mw.root.after(0, lambda h=handler, g=m: self._safe_run(h, g))
                return

        # Fall back to AI
        self.mw.root.after(0, lambda: self.chat.status_label.config(text="未匹配快捷命令，AI思考中..."))
        self._ai_dispatch(text)

    def _restore_ui(self):
        if self.chat._running:
            self.mw.root.after(0, self.chat._hide_stop)

    def _safe_run(self, handler, match):
        if self._cancelled:
            self._restore_ui()
            return
        try:
            handler(match)
        except Exception as e:
            self.chat._add_message("error", f"执行失败: {e}")
            self.chat.set_status("")
        finally:
            self._restore_ui()

    # ---- Shortcut handlers ------------------------------------------------

    def _cmd_help(self, m):
        help_text = """可用命令：
• 搜索 <关键词> — 在搜索引擎中搜索
• 核验 — 切换到真实性核验页
• 切换到搜索页/整理页/分析页 — 切换标签
• 刷新 — 刷新当前列表
• 导入/导出 — 导入或导出数据
• 自动补全 — AI一键补全当前案例
• 添加案例 — 新建案例
• 删除选中 — 删除选中的记录
• 查看详情 — 查看选中案例详情
• 下一个/上一个 — 导航案例
• 提交核验 — 提交选中项到核验队列
• 批准/拒绝 — 核验通过/不通过
• 帮助 — 显示此帮助"""
        self.chat._add_message("assistant", help_text)
        self.chat.set_status("")

    def _cmd_workflow_search(self, m):
        country = m.group(1).strip()
        language = m.group(2).strip()
        if not country or not language:
            self.chat._add_message("assistant", "请说明国家和语言，如：一键搜索 日本 中文")
            self.chat.set_status("")
            return
        threading.Thread(target=lambda: self._run_workflow("一键搜索",
            country=country, language=language), daemon=True).start()

    def _cmd_ds_verify(self, m):
        self._run_workflow("一键DS核验")

    def _cmd_g_verify(self, m):
        self._run_workflow("一键G核验")

    def _cmd_verify(self, m):
        self._run_workflow("一键核验")

    def _run_workflow(self, name, **params):
        from ui.workflows import execute_workflow
        result = execute_workflow(name, params, chat=self.chat, mw=self.mw)
        # If the workflow is interactive, store the session
        from ui.workflows import WORKFLOW_REGISTRY
        wf = WORKFLOW_REGISTRY.get(name, {})
        if wf.get("is_interactive") and result:
            self.active_session = result

    def _cmd_chitchat(self, m):
        text = m.group(0).lower()
        if any(w in text for w in ["你好", "您好", "嗨", "hi", "hello"]):
            self.chat._add_message("assistant", "你好！有什么可以帮你的？输入「帮助」查看我能做什么。")
        elif any(w in text for w in ["谢谢", "谢了", "多谢"]):
            self.chat._add_message("assistant", "不客气！")
        elif any(w in text for w in ["再见", "拜拜", "bye"]):
            self.chat._add_message("assistant", "再见！")
        self.chat.set_status("")

    def _cmd_search(self, m):
        keyword = m.group(2).strip()
        if not keyword:
            self.chat._add_message("assistant", "请输入搜索关键词，如：搜索 唐人街")
            self.chat.set_status("")
            return
        # Switch to search tab (index 0)
        self.mw.notebook.select(0)
        # Set keyword in search entry
        self.mw.search_tab.web_search_var.set(keyword)
        self.chat._add_message("assistant", f"已切换到搜索页，关键词：{keyword}。请选择搜索引擎后点击搜索。")
        self.chat.set_status("")

    def _cmd_search_natural(self, m):
        # m.group(4) captures the keyword from "帮我搜 X"
        keyword = m.group(4).strip() if m.lastindex and m.lastindex >= 4 else m.group(2).strip()
        if not keyword:
            self.chat._add_message("assistant", "请说出要搜索的关键词，如：搜索唐人街")
            self.chat.set_status("")
            return
        self.mw.notebook.select(0)
        self.mw.search_tab.web_search_var.set(keyword)
        self.chat._add_message("assistant", f"已切换到搜索页，关键词：{keyword}。请选择搜索引擎后点击搜索。")
        self.chat.set_status("")

    def _cmd_switch_verify(self, m):
        self.mw.notebook.select(1)  # organize tab
        self.mw.organize_tab.notebook.select(1)  # verify subtab
        self.chat._add_message("assistant", "已切换到真实性核验页。")
        self.chat.set_status("")

    def _cmd_switch_search(self, m):
        self.mw.notebook.select(0)
        self.chat._add_message("assistant", "已切换到案例查询页。")
        self.chat.set_status("")

    def _cmd_switch_organize(self, m):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(0)  # raw data
        self.chat._add_message("assistant", "已切换到数据整理页。")
        self.chat.set_status("")

    def _cmd_switch_analysis(self, m):
        self.mw.notebook.select(2)
        self.chat._add_message("assistant", "已切换到数据分析页。")
        self.chat.set_status("")

    def _cmd_switch_raw(self, m):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(0)
        self.chat._add_message("assistant", "已切换到原始数据页。")
        self.chat.set_status("")

    def _cmd_switch_database(self, m):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.chat._add_message("assistant", "已切换到正式案例页。")
        self.chat.set_status("")

    def _cmd_refresh(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 0:
            self.mw.search_tab.refresh()
        elif tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                self.mw.organize_tab.raw_panel.refresh()
            elif sub_idx == 1:
                self.mw.organize_tab.verify_panel.refresh()
            elif sub_idx == 2:
                self.mw.organize_tab.db_panel.refresh()
        self.chat._add_message("assistant", "已刷新当前列表。")
        self.chat.set_status("")

    def _cmd_import(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                self.mw.organize_tab.raw_panel._import_data()
            elif sub_idx == 2:
                self.mw.organize_tab.db_panel._import_data()
        self.chat._add_message("assistant", "已打开导入对话框。")
        self.chat.set_status("")

    def _cmd_export(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                self.mw.organize_tab.raw_panel._export_data()
            elif sub_idx == 2:
                self.mw.organize_tab.db_panel._export_data_cn()
        self.chat._add_message("assistant", "已打开导出对话框。")
        self.chat.set_status("")

    def _cmd_auto_complete(self, m):
        self.mw.notebook.select(1)  # switch to organize
        self.mw.organize_tab.notebook.select(1)  # switch to verify
        # The verify panel's _auto_complete_all requires a current case
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_current_case") and vp._current_case:
            vp._auto_complete_all()
            self.chat._add_message("assistant", "已触发一键补全。")
        else:
            self.chat._add_message("assistant", "请先在核验页面选择一个案例。")
        self.chat.set_status("")

    def _cmd_delete(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 0:
            self.mw.search_tab._delete_selected()
        elif tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                self.mw.organize_tab.raw_panel._delete_selected()
            elif sub_idx == 1:
                self.mw.organize_tab.verify_panel._delete_selected()
            elif sub_idx == 2:
                self.mw.organize_tab.db_panel._delete_selected()
        self.chat._add_message("assistant", "已执行删除操作。")
        self.chat.set_status("")

    def _cmd_add_case(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 2:
                self.mw.organize_tab.db_panel._add_case()
                self.chat._add_message("assistant", "已打开新建案例对话框。")
        else:
            self.mw.notebook.select(1)
            self.mw.organize_tab.notebook.select(2)
            self.mw.organize_tab.db_panel._add_case()
            self.chat._add_message("assistant", "已切换到正式案例页并打开新建对话框。")
        self.chat.set_status("")

    def _cmd_view_detail(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 0:
            self.mw.search_tab._show_detail()
        elif tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                self.mw.organize_tab.raw_panel._view_detail()
            elif sub_idx == 1:
                self.mw.organize_tab.verify_panel._open_verify_detail()
            elif sub_idx == 2:
                self.mw.organize_tab.db_panel._on_db_tree_click(None)
        self.chat._add_message("assistant", "已尝试打开详情。")
        self.chat.set_status("")

    def _cmd_next(self, m):
        self._navigate_tree(1)
        self.chat.set_status("")

    def _cmd_prev(self, m):
        self._navigate_tree(-1)
        self.chat.set_status("")

    def _navigate_tree(self, direction):
        """Move tree selection up or down by one row."""
        tab_idx = self.mw.notebook.index("current")
        tree = None
        if tab_idx == 0:
            tree = self.mw.search_tab.tree
        elif tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                tree = self.mw.organize_tab.raw_panel.tree
            elif sub_idx == 1:
                tree = self.mw.organize_tab.verify_panel.queue_tree
            elif sub_idx == 2:
                tree = self.mw.organize_tab.db_panel.tree

        if tree is None:
            return

        children = tree.get_children()
        if not children:
            return

        sel = tree.selection()
        if not sel:
            target = children[0]
        else:
            idx = children.index(sel[0])
            new_idx = idx + direction
            if new_idx < 0 or new_idx >= len(children):
                return
            target = children[new_idx]

        tree.selection_set(target)
        tree.see(target)
        # Trigger the appropriate select handler
        if tab_idx == 0:
            self.mw.search_tab._show_detail()
        elif tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 1:
                self.mw.organize_tab.verify_panel._on_select(None)

    def _cmd_submit_verify(self, m):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(0)
        self.mw.organize_tab.raw_panel._submit_for_verification()
        self.chat._add_message("assistant", "已提交核验。")
        self.chat.set_status("")

    def _cmd_approve(self, m):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_current_case") and vp._current_case:
            vp._approve()
            self.chat._add_message("assistant", "已批准当前核验。")
        else:
            self.chat._add_message("assistant", "请先在核验页面选择一个案例。")
        self.chat.set_status("")

    def _cmd_reject(self, m):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_current_case") and vp._current_case:
            vp._reject()
            self.chat._add_message("assistant", "已拒绝当前核验。")
        else:
            self.chat._add_message("assistant", "请先在核验页面选择一个案例。")
        self.chat.set_status("")

    def _cmd_select_all(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 0:
            tree = self.mw.search_tab.tree
        elif tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                tree = self.mw.organize_tab.raw_panel.tree
            elif sub_idx == 1:
                tree = self.mw.organize_tab.verify_panel.queue_tree
            elif sub_idx == 2:
                tree = self.mw.organize_tab.db_panel.tree
            else:
                tree = None
        else:
            tree = None
        if tree:
            tree.selection_set(tree.get_children())
            self.chat._add_message("assistant", "已全选。")
        self.chat.set_status("")

    def _cmd_clear_selection(self, m):
        tab_idx = self.mw.notebook.index("current")
        if tab_idx == 0:
            tree = self.mw.search_tab.tree
        elif tab_idx == 1:
            sub_idx = self.mw.organize_tab.notebook.index("current")
            if sub_idx == 0:
                tree = self.mw.organize_tab.raw_panel.tree
            elif sub_idx == 1:
                tree = self.mw.organize_tab.verify_panel.queue_tree
            elif sub_idx == 2:
                tree = self.mw.organize_tab.db_panel.tree
            else:
                tree = None
        else:
            tree = None
        if tree:
            tree.selection_set()
            self.chat._add_message("assistant", "已取消选择。")
        self.chat.set_status("")

    def _cmd_cancel_workflow(self, m):
        if self.active_session and self.active_session.active:
            self.active_session.active = False
            self.chat._add_message("assistant", "工作流已取消。")
            self.chat.set_status("")
            self.active_session = None
        else:
            self.chat._add_message("assistant", "当前没有运行中的工作流。")

    # ---- AI fallback -------------------------------------------------------

    def _ai_dispatch(self, text):
        """Send to DeepSeek for natural language understanding."""
        try:
            client = openai.OpenAI(api_key=_CHAT_API_KEY,
                                   base_url=API_BASE_URL)
            response = client.chat.completions.create(
                model=DEFAULT_MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": text},
                ],
                temperature=0.1,
                max_tokens=256,
            )
            raw = response.choices[0].message.content
            result = self._parse_ai_response(raw)
        except Exception as e:
            self.mw.root.after(0, lambda: self._ai_fail(f"AI调用失败: {e}"))
            self._restore_ui()
            return

        if self._cancelled:
            self._restore_ui()
            return
        self.mw.root.after(0, lambda r=result: self._execute_ai_action(r))
        self._restore_ui()

    def _parse_ai_response(self, raw):
        try:
            # Try direct JSON parse
            return json.loads(raw)
        except json.JSONDecodeError:
            # Try to extract JSON from markdown code blocks
            m = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', raw, re.DOTALL)
            if m:
                return json.loads(m.group(1))
            # Try to find bare JSON object
            m = re.search(r'\{.*\}', raw, re.DOTALL)
            if m:
                return json.loads(m.group(0))
        return {"action": "unknown", "params": {}}

    def _execute_ai_action(self, result):
        if self._cancelled:
            self._restore_ui()
            return
        action = result.get("action", "unknown")
        params = result.get("params", {})

        handler = self.AI_HANDLERS.get(action)
        if handler:
            try:
                handler(self, params)
                self.chat._add_message("assistant", f"已执行：{action}")
            except Exception as e:
                self.chat._add_message("error", f"执行 {action} 失败: {e}")
        else:
            self.chat._add_message("assistant",
                                   "抱歉，我没有理解你的意思。请换一种说法，或输入「帮助」查看可用命令。")
        self.chat.set_status("")
        self._restore_ui()

    # ==================================================================
    #  AI Action Handlers — one for every supported action
    # ==================================================================

    def _ai_switch_tab(self, p):
        idx = {"search": 0, "organize": 1, "analysis": 2}.get(p.get("tab", ""))
        if idx is not None:
            self.mw.notebook.select(idx)
            names = {0: "案例查询", 1: "数据整理", 2: "数据分析"}
            self.chat._add_message("assistant", f"已切换到{names[idx]}页。")

    def _ai_switch_subtab(self, p):
        idx = {"raw": 0, "verify": 1, "database": 2}.get(p.get("subtab", ""))
        if idx is not None:
            self.mw.notebook.select(1)
            self.mw.organize_tab.notebook.select(idx)
            names = {0: "原始数据", 1: "真实性核验", 2: "正式案例"}
            self.chat._add_message("assistant", f"已切换到{names[idx]}页。")

    def _ai_web_search(self, p):
        kw = p.get("keyword", "").strip()
        if kw:
            self.mw.notebook.select(0)
            self.mw.search_tab.web_search_var.set(kw)
            self.chat._add_message("assistant", f"已填入搜索关键词：{kw}。请选择搜索引擎后点击搜索。")

    def _ai_generate_keywords(self, p):
        self.mw.notebook.select(0)
        self.mw.search_tab._open_keyword_gen()
        self.chat._add_message("assistant", "已打开关键词生成对话框。")

    def _ai_local_search(self, p):
        self.mw.notebook.select(0)
        if p.get("keyword"):
            self.mw.search_tab.local_search_var.set(p["keyword"])
        if p.get("type"):
            self.mw.search_tab.type_var.set(p["type"])
        if p.get("country"):
            self.mw.search_tab.country_var.set(p["country"])
        self.mw.search_tab.refresh()
        self.chat._add_message("assistant", "已执行库内搜索并刷新。")

    def _ai_clear_filter(self, p):
        self.mw.notebook.select(0)
        self.mw.search_tab._clear_filter()
        self.chat._add_message("assistant", "已清除筛选条件。")

    def _ai_show_detail(self, p):
        self.mw.search_tab._show_detail()
        self.chat._add_message("assistant", "已打开详情。")

    def _ai_delete_selected(self, p):
        self._cmd_delete(None)

    def _ai_select_all(self, p):
        self._cmd_select_all(None)

    def _ai_clear_selection(self, p):
        self._cmd_clear_selection(None)

    def _ai_ai_extract(self, p):
        self.mw.notebook.select(0)
        self.mw.search_tab._ai_extract()
        self.chat._add_message("assistant", "已开始AI提取。")

    def _ai_ai_load_file(self, p):
        self.mw.notebook.select(0)
        self.mw.search_tab._ai_load_file()
        self.chat._add_message("assistant", "已打开文件选择对话框。")

    def _ai_ai_export(self, p):
        self.mw.notebook.select(0)
        self.mw.search_tab._ai_export()
        self.chat._add_message("assistant", "已导出AI提取结果。")

    def _ai_ai_add_to_db(self, p):
        self.mw.notebook.select(0)
        self.mw.search_tab._ai_add_to_db()
        self.chat._add_message("assistant", "已将提取结果添加到数据库。")

    def _ai_ai_clear(self, p):
        self.mw.notebook.select(0)
        self.mw.search_tab._ai_clear()
        self.chat._add_message("assistant", "已清空AI提取结果。")

    def _ai_submit_verify(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(0)
        self.mw.organize_tab.raw_panel._submit_for_verification()
        self.chat._add_message("assistant", "已提交核验。")

    def _ai_delete_raw(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(0)
        self.mw.organize_tab.raw_panel._delete_selected()
        self.chat._add_message("assistant", "已删除选中的原始数据。")

    def _ai_refresh_raw(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(0)
        self.mw.organize_tab.raw_panel.refresh()
        self.chat._add_message("assistant", "已刷新原始数据列表。")

    def _ai_refresh_queue(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel.refresh()
        self.chat._add_message("assistant", "已刷新核验队列。")

    def _ai_open_verify_detail(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._open_verify_detail()
        self.chat._add_message("assistant", "已打开核验详情。")

    def _ai_enable_field_edit(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._enable_field_edit()
        self.chat._add_message("assistant", "已进入字段编辑模式。修改后请说「保存修改」。")

    def _ai_save_field_changes(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._save_field_changes()
        self.chat._add_message("assistant", "已保存字段修改。")

    def _ai_auto_complete_all(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_current_case") and vp._current_case:
            vp._auto_complete_all()
            self.chat._add_message("assistant", "已执行一键补全。")
        else:
            self.chat._add_message("assistant", "请先在核验页面选择一个案例。")

    def _ai_approve(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_current_case") and vp._current_case:
            vp._approve()
            self.chat._add_message("assistant", "已通过核验。")
        else:
            self.chat._add_message("assistant", "请先在核验页面选择一个案例。")

    def _ai_reject(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_current_case") and vp._current_case:
            vp._reject()
            self.chat._add_message("assistant", "已拒绝核验。")
        else:
            self.chat._add_message("assistant", "请先在核验页面选择一个案例。")

    def _ai_apply_field_result(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._apply_field_result()
        self.chat._add_message("assistant", "已应用字段核验结果。")

    def _ai_save_manual_result(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._save_manual_result()
        self.chat._add_message("assistant", "已保存核验结果。")

    def _ai_clear_result(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._clear_result()
        self.chat._add_message("assistant", "已清空核验结果。")

    def _ai_clear_paste(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._clear_paste_area()
        self.chat._add_message("assistant", "已清空粘贴内容。")

    def _ai_name_field_analysis(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._run_name_field_analysis()
        self.chat._add_message("assistant", "已分析名称字段。")

    def _ai_other_field_analysis(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._parse_and_fill_from_text()
        self.chat._add_message("assistant", "已分析其他字段。")

    def _ai_download_image(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._download_image()
        self.chat._add_message("assistant", "已下载图片。")

    def _ai_undo_delete_image(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._undo_delete_image()
        self.chat._add_message("assistant", "已撤销图片删除。")

    def _ai_refresh_image(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._refresh_image_preview(None, None)
        self.chat._add_message("assistant", "已刷新图片。")

    def _ai_search_field(self, p):
        field = p.get("field", "")
        engine = p.get("engine", "baidu")
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if not hasattr(vp, "_current_case") or not vp._current_case:
            self.chat._add_message("assistant", "请先在核验页面选择一个案例。")
            return
        # Map engine to the correct method
        if engine == "ds":
            vp._open_field_deepseek(field, None)
        elif engine == "gemini":
            vp._verify_field_gemini(field, None)
        else:
            vp._open_field_search(field, None, engine)
        self.chat._add_message("assistant", f"已用{engine}搜索{field}字段。")

    def _ai_complete_field(self, p):
        field = p.get("field", "")
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_set_active_field"):
            vp._set_active_field(field, None)
        self.chat._add_message("assistant", f"已补全{field}字段。")

    def _ai_import_data(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._import_data()
        self.chat._add_message("assistant", "已打开导入对话框。")

    def _ai_export_data(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._export_data_cn()
        self.chat._add_message("assistant", "已打开导出对话框。")

    def _ai_find_duplicates(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._show_duplicates()
        self.chat._add_message("assistant", "已查找重复案例。")

    def _ai_clean_data(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._clean_data()
        self.chat._add_message("assistant", "已执行数据清洗。")

    def _ai_check_missing(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._check_missing()
        self.chat._add_message("assistant", "已执行缺失检查。")

    def _ai_add_case(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._add_case()
        self.chat._add_message("assistant", "已打开新增案例对话框。")

    def _ai_edit_selected(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._edit_selected()
        self.chat._add_message("assistant", "已打开编辑对话框。")

    def _ai_delete_db(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel._delete_selected()
        self.chat._add_message("assistant", "已删除选中的正式案例。")

    def _ai_refresh_db(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(2)
        self.mw.organize_tab.db_panel.refresh()
        self.chat._add_message("assistant", "已刷新正式案例列表。")

    def _ai_chart_by(self, p):
        dim = p.get("dimension", "type")
        dim_map = {"type": "building_type", "country": "country",
                   "year": "year_built", "protection": "protection_level"}
        var = dim_map.get(dim, dim)
        self.mw.notebook.select(2)
        self.mw.analysis_tab.dim_var.set(var)
        self.mw.analysis_tab._refresh_chart()
        self.chat._add_message("assistant", f"已切换到{var}维度图表。")

    def _ai_export_all(self, p):
        self.mw.notebook.select(2)
        self.mw.analysis_tab._export_all()
        self.chat._add_message("assistant", "已导出全部数据。")

    def _ai_export_stats(self, p):
        self.mw.notebook.select(2)
        self.mw.analysis_tab._export_stats()
        self.chat._add_message("assistant", "已导出统计数据。")

    def _ai_refresh(self, p):
        self._cmd_refresh(None)

    def _ai_next_case(self, p):
        self._cmd_next(None)

    def _ai_prev_case(self, p):
        self._cmd_prev(None)

    def _ai_help(self, p):
        self._cmd_help(None)

    def _ai_workflow(self, p):
        name = p.get("name", "")
        country = p.get("country", "")
        language = p.get("language", "")
        if not name:
            self.chat._add_message("assistant", "请说明工作流名称。如：一键DS核验")
            return
        # Check if it's an interactive workflow that takes no params
        from ui.workflows import WORKFLOW_REGISTRY
        wf = WORKFLOW_REGISTRY.get(name, {})
        if wf.get("is_interactive"):
            self._run_workflow(name)
            return
        if not country or not language:
            self.chat._add_message("assistant", "请说明工作流名称、国家和语言。如：一键搜索 日本 中文")
            return
        threading.Thread(target=lambda: self._run_workflow(name,
            country=country, language=language), daemon=True).start()

    def _ai_ds_verify(self, p):
        self._run_workflow("一键DS核验")

    def _ai_g_verify(self, p):
        self._run_workflow("一键G核验")

    def _ai_verify(self, p):
        self._run_workflow("一键核验")

    def _ai_delete_verify(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._delete_selected()
        self.chat._add_message("assistant", "已从核验队列删除选中项。")

    def _ai_verify_field(self, p):
        field = p.get("field", "")
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if hasattr(vp, "_set_active_field"):
            vp._set_active_field(field, None)
        self.chat._add_message("assistant", f"已触发{field}字段核验。")

    def _ai_search_active_field(self, p):
        engine = p.get("engine", "baidu")
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        vp = self.mw.organize_tab.verify_panel
        if engine == "ds":
            vp._verify_field_deepseek(vp._active_field, None)
        elif engine == "gemini":
            vp._verify_active_field_gemini()
        else:
            vp._search_active_field(engine)
        self.chat._add_message("assistant", f"已用{engine}搜索当前活跃字段。")

    def _ai_select_local_image(self, p):
        self.mw.notebook.select(1)
        self.mw.organize_tab.notebook.select(1)
        self.mw.organize_tab.verify_panel._select_local_image()
        self.chat._add_message("assistant", "已打开本地图片选择。")

    def _ai_set_engines(self, p):
        group = p.get("group", "grey")  # "grey" or "research"
        select_all = p.get("select_all", True)
        self.mw.notebook.select(0)
        from ui.search_tab import GREY_LITERATURE, RESEARCH_LITERATURE
        engines = GREY_LITERATURE if group == "grey" else RESEARCH_LITERATURE
        self.mw.search_tab._set_engines(engines, select_all)
        self.chat._add_message("assistant",
            f"已{'全选' if select_all else '全不选'}{'灰色文献' if group == 'grey' else '研究文献'}搜索引擎。")

    # Master dispatch table
    AI_HANDLERS = {
        "switch_tab": _ai_switch_tab,
        "switch_subtab": _ai_switch_subtab,
        "web_search": _ai_web_search,
        "generate_keywords": _ai_generate_keywords,
        "local_search": _ai_local_search,
        "clear_filter": _ai_clear_filter,
        "show_detail": _ai_show_detail,
        "delete_selected": _ai_delete_selected,
        "select_all": _ai_select_all,
        "clear_selection": _ai_clear_selection,
        "ai_extract": _ai_ai_extract,
        "ai_load_file": _ai_ai_load_file,
        "ai_export": _ai_ai_export,
        "ai_add_to_db": _ai_ai_add_to_db,
        "ai_clear": _ai_ai_clear,
        "submit_verify": _ai_submit_verify,
        "delete_raw": _ai_delete_raw,
        "refresh_raw": _ai_refresh_raw,
        "refresh_queue": _ai_refresh_queue,
        "open_verify_detail": _ai_open_verify_detail,
        "enable_field_edit": _ai_enable_field_edit,
        "save_field_changes": _ai_save_field_changes,
        "auto_complete_all": _ai_auto_complete_all,
        "auto_complete": _ai_auto_complete_all,
        "approve": _ai_approve,
        "reject": _ai_reject,
        "apply_field_result": _ai_apply_field_result,
        "save_manual_result": _ai_save_manual_result,
        "clear_result": _ai_clear_result,
        "clear_paste": _ai_clear_paste,
        "name_field_analysis": _ai_name_field_analysis,
        "other_field_analysis": _ai_other_field_analysis,
        "download_image": _ai_download_image,
        "undo_delete_image": _ai_undo_delete_image,
        "refresh_image": _ai_refresh_image,
        "search_field": _ai_search_field,
        "complete_field": _ai_complete_field,
        "verify_field": _ai_verify_field,
        "search_active_field": _ai_search_active_field,
        "select_local_image": _ai_select_local_image,
        "delete_verify": _ai_delete_verify,
        "set_engines": _ai_set_engines,
        "import_data": _ai_import_data,
        "export_data": _ai_export_data,
        "find_duplicates": _ai_find_duplicates,
        "clean_data": _ai_clean_data,
        "check_missing": _ai_check_missing,
        "add_case": _ai_add_case,
        "edit_selected": _ai_edit_selected,
        "delete_db": _ai_delete_db,
        "refresh_db": _ai_refresh_db,
        "chart_by": _ai_chart_by,
        "export_all": _ai_export_all,
        "export_stats": _ai_export_stats,
        "refresh": _ai_refresh,
        "next_case": _ai_next_case,
        "prev_case": _ai_prev_case,
        "help": _ai_help,
        "workflow": _ai_workflow,
        "ds_verify": _ai_ds_verify,
        "g_verify": _ai_g_verify,
        "verify": _ai_verify,
    }

    def _ai_fail(self, msg):
        self.chat._add_message("error", f"AI调用失败: {msg}\n请换一种说法再试，或输入「帮助」查看可用命令。")
        self.chat.set_status("")
        self._restore_ui()

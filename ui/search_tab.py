import json
import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from http.server import HTTPServer, BaseHTTPRequestHandler
import webbrowser
import urllib.parse
import urllib.request
import database as db
import config
from ui.keyword_gen import KeywordGenDialog
from ui.image_components import ImageThumbnailStrip, FullImageViewer, parse_image_urls, load_tree_thumbnail


def _make_tree_draggable(tree):
    """Enable drag-and-drop reordering on a grouped Treeview.
    Supports multi-select: drag multiple selected items together.
    Group headers (G_) move among root-level G_ nodes.
    Leaf items from the same parent move within their group."""
    drag_data = {"items": [], "parent": None, "start_y": 0, "dragging": False}

    def _on_press(event):
        region = tree.identify_region(event.x, event.y)
        if region in ("heading", "separator"):
            return
        row = tree.identify_row(event.y)
        if not row:
            drag_data["items"] = []
            return
        sel = list(tree.selection())
        if row in sel:
            # Drag all selected items together
            drag_data["items"] = sel
        else:
            tree.selection_set(row)
            drag_data["items"] = [row]
        drag_data["parent"] = tree.parent(row)
        drag_data["start_y"] = event.y
        drag_data["dragging"] = False

    def _on_motion(event):
        items = drag_data.get("items")
        if not items:
            return
        if not drag_data["dragging"]:
            if abs(event.y - drag_data["start_y"]) < 5:
                return
            drag_data["dragging"] = True
            tree.configure(cursor="fleur")

        target = tree.identify_row(event.y)
        if not target or target in items:
            return

        # Determine the type of drag operation
        all_groups = all(i.startswith("G_") for i in items)

        # For leaf items, check they're all from the same parent
        all_leaves_same_parent = False
        move_parent = ""
        if not all_groups:
            parents = set()
            mixed = False
            for i in items:
                if i.startswith("G_"):
                    mixed = True
                    break
                parents.add(tree.parent(i))
            if not mixed and len(parents) == 1:
                all_leaves_same_parent = True
                move_parent = list(parents)[0] or ""

        if all_groups and target.startswith("G_"):
            # Move all selected G_ nodes before target
            children = list(tree.get_children(""))
            for item in items:
                if item in children:
                    children.remove(item)
            try:
                idx = children.index(target)
                for i, item in enumerate(items):
                    children.insert(idx + i, item)
                for i, cid in enumerate(children):
                    tree.move(cid, "", i)
            except ValueError:
                pass
        elif all_leaves_same_parent and not target.startswith("G_"):
            # Move all selected leaves within their parent group
            children = list(tree.get_children(move_parent))
            for item in items:
                if item in children:
                    children.remove(item)
            try:
                idx = children.index(target)
                for i, item in enumerate(items):
                    children.insert(idx + i, item)
                for i, cid in enumerate(children):
                    tree.move(cid, move_parent, i)
            except ValueError:
                pass

    def _on_release(event):
        if not drag_data.get("dragging"):
            drag_data["items"] = []
            return
        drag_data["dragging"] = False
        drag_data["items"] = []
        drag_data["parent"] = None
        tree.configure(cursor="")

    tree.bind("<ButtonPress-1>", _on_press)
    tree.bind("<B1-Motion>", _on_motion)
    tree.bind("<ButtonRelease-1>", _on_release)
from utils.ai_extractor import extract_case_names
from utils.web_search import PROXY
from utils.import_export import export_to_excel
import pandas as pd


def _append_to_paste(vp, text):
    """Append text to verify panel's paste_text widget."""
    try:
        vp.paste_text.configure(state="normal")
        vp.paste_text.delete("1.0", "end")
        vp.paste_text.insert("end", text)
        vp.paste_text.see("end")
        vp.paste_text.configure(state="disabled")
    except Exception:
        pass


class _ExtractServerHandler(BaseHTTPRequestHandler):
    verify_receiver = None  # 由 VerificationPanel 设置，接收 (text, url) 用于字段核验

    def do_POST(self):
        if self.path == '/extract':
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))
            text = body.get('text', '')
            url = body.get('url', '')
            title = body.get('title', '')
            description = body.get('description', '')
            author = body.get('author', '')
            pub_date = body.get('pub_date', '')
            # Store metadata for literature library
            self.server.tab._browser_last_title = title
            self.server.tab._browser_last_url = url
            self.server.tab._browser_last_description = description
            self.server.tab._browser_last_author = author
            self.server.tab._browser_last_pub_date = pub_date
            # Route based on which tab is active
            mw = self.server.tab.frame.winfo_toplevel()
            tab_idx = mw.notebook.index("current") if hasattr(mw, 'notebook') else 0
            if tab_idx == 1:
                # 真实性核验页 — verify_receiver (字段核验) or paste_text (搜索结果内容)
                receiver = _ExtractServerHandler.verify_receiver
                if receiver:
                    mw.after(0, receiver, text, url)
                else:
                    vp = mw.organize_tab.verify_panel
                    mw.after(0, lambda t=text: _append_to_paste(vp, t))
            else:
                # 案例查询页 — 文章内容
                mw.after(0, self.server.tab._receive_browser_text, text, url)
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            resp = json.dumps({'ok': True, 'count': len(text)})
            self.wfile.write(resp.encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, format, *args):
        pass


GREY_LITERATURE = {
    "Google": "https://www.google.com/search?q={}",
    "百度": "https://www.baidu.com/s?wd={}",
}

RESEARCH_LITERATURE = {
    "Web of Science": "https://www.webofscience.com/wos/woscc/basic-search",
    "Google Scholar": "https://scholar.google.com/scholar?q={}",
    "CNKI (知网)": "https://kns.cnki.net/kns8/defaultresult/index?kwd={}",
    "中国国家图书馆": "http://find.nlc.cn/search/showDocDetails?query={}",
    "罗马国家图书馆": "https://opac.bncre.ie-roma.it/ricerca.asp?query={}",
}

ALL_ENGINES = {**GREY_LITERATURE, **RESEARCH_LITERATURE}


class SearchTab:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.ai_results = []
        self._photo_cache = {}
        self._browser_last_title = ""
        self._browser_last_url = ""
        self._browser_last_description = ""
        self._browser_last_author = ""
        self._browser_last_pub_date = ""
        self._start_extract_server()
        self._build_web_search_section()
        self._build_ai_extract_section()
        self._build_local_search_bar()
        self._build_tree()
        self._build_context_menu()
        self.refresh()

    def _build_web_search_section(self):
        section = ttk.LabelFrame(self.frame, text=" 搜索引擎搜索 ", padding=5)
        section.pack(fill="x", padx=5, pady=(5, 0))

        # Row 0: keyword entry + search button + keyword listbox on the right
        kw_frame = ttk.Frame(section)
        kw_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(kw_frame, text="关键词：").pack(side="left")
        self.web_search_var = tk.StringVar()
        web_entry = ttk.Entry(kw_frame, textvariable=self.web_search_var, width=30)
        web_entry.pack(side="left", padx=5)
        web_entry.bind("<Return>", lambda e: self._web_search())
        ttk.Button(kw_frame, text="生成关键词", command=self._open_keyword_gen).pack(side="left", padx=2)
        ttk.Button(kw_frame, text="搜索", command=self._web_search).pack(side="left", padx=2)
        ttk.Label(kw_frame, text="（在 Chrome 中打开搜索结果）", foreground="gray").pack(side="left", padx=5)

        # Keyword listbox on the same row, to the right
        kw_list_frame = ttk.LabelFrame(kw_frame, text=" 选中关键词 ", padding=3)
        kw_list_frame.pack(side="right", fill="x", expand=True)

        self.kw_listbox = tk.Listbox(kw_list_frame, selectmode="extended", height=3, width=20,
                                     exportselection=False)
        kw_scroll = ttk.Scrollbar(kw_list_frame, orient="vertical", command=self.kw_listbox.yview)
        self.kw_listbox.configure(yscrollcommand=kw_scroll.set)
        self.kw_listbox.pack(side="left", fill="both", expand=True)
        kw_scroll.pack(side="right", fill="y")

        kw_btn_bar = ttk.Frame(kw_list_frame)
        kw_btn_bar.pack(fill="x", pady=(1, 0))
        ttk.Button(kw_btn_bar, text="全选", width=5,
                   command=lambda: self.kw_listbox.select_set(0, tk.END)).pack(side="left", padx=1)
        ttk.Button(kw_btn_bar, text="清空", width=5,
                   command=lambda: self.kw_listbox.selection_clear(0, tk.END)).pack(side="left", padx=1)

        # Row 1: grey literature + research literature (2 columns)
        panels_frame = ttk.Frame(section)
        panels_frame.pack(fill="x")

        self.engine_vars = {}
        for name in ALL_ENGINES:
            var = tk.BooleanVar(value=True)
            self.engine_vars[name] = var

        # -- Grey literature panel --
        grey_frame = ttk.LabelFrame(panels_frame, text=" 灰色文献 ", padding=5)
        grey_frame.pack(side="left", fill="both", padx=(0, 5))

        btn_bar = ttk.Frame(grey_frame)
        btn_bar.pack(fill="x", pady=(0, 3))
        ttk.Button(btn_bar, text="全选", command=lambda: self._set_engines(GREY_LITERATURE, True)).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="全不选", command=lambda: self._set_engines(GREY_LITERATURE, False)).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="api调用筛选", command=self._serper_search).pack(side="right", padx=2)

        for name in GREY_LITERATURE:
            tk.Checkbutton(grey_frame, text=name, variable=self.engine_vars[name],
                           onvalue=True, offvalue=False, anchor="w",
                           fg="black", selectcolor="#ffffff",
                           bg="#ffffff", activebackground="#f0f0f0").pack(anchor="w", fill="x")

        # -- Research literature panel --
        research_frame = ttk.LabelFrame(panels_frame, text=" 研究文献 ", padding=5)
        research_frame.pack(side="left", fill="both", padx=(5, 0))

        btn_bar2 = ttk.Frame(research_frame)
        btn_bar2.pack(fill="x", pady=(0, 3))
        ttk.Button(btn_bar2, text="全选", command=lambda: self._set_engines(RESEARCH_LITERATURE, True)).pack(side="left", padx=2)
        ttk.Button(btn_bar2, text="全不选", command=lambda: self._set_engines(RESEARCH_LITERATURE, False)).pack(side="left", padx=2)
        ttk.Button(btn_bar2, text="文献筛选", command=self._open_literature_search).pack(side="right", padx=2)
        ttk.Button(btn_bar2, text="文献库", command=self._open_literature_library).pack(side="right", padx=2)

        for name in RESEARCH_LITERATURE:
            tk.Checkbutton(research_frame, text=name, variable=self.engine_vars[name],
                           onvalue=True, offvalue=False, anchor="w",
                           fg="black", selectcolor="#ffffff",
                           bg="#ffffff", activebackground="#f0f0f0").pack(anchor="w", fill="x")

    def _start_extract_server(self):
        import socket
        import subprocess
        # Kill any existing process on port 9876 (leftover from previous session)
        try:
            out = subprocess.check_output(
                ["/usr/sbin/lsof", "-ti", ":9876"], stderr=subprocess.DEVNULL, timeout=5
            ).decode().strip()
            if out:
                for pid in out.split("\n"):
                    try:
                        subprocess.run(["/bin/kill", "-9", pid.strip()], timeout=3)
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            handler = _ExtractServerHandler
            self._extract_server = HTTPServer(('127.0.0.1', 9876), handler)
            self._extract_server.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self._extract_server.tab = self
            t = threading.Thread(target=self._extract_server.serve_forever, daemon=True)
            t.start()
            import sys
            print("[extract-server] started on port 9876", file=sys.stderr, flush=True)
        except Exception as e:
            import sys
            import traceback
            print(f"[extract-server] failed to start: {e}", file=sys.stderr, flush=True)
            traceback.print_exc()
            self._extract_server = None

    def _receive_browser_text(self, text, url):
        current = self.ai_text_input.get("1.0", "end-1c").strip()
        if current:
            self.ai_text_input.insert("end", f"\n\n--- [来源] {url} ---\n\n{text}")
        else:
            self.ai_text_input.insert("end", f"--- [来源] {url} ---\n\n{text}")
        count = len(text)
        total = len(self.ai_text_input.get("1.0", "end-1c"))
        self.ai_status_var.set(f"已添加 {count} 字符（累计 {total} 字符），可继续浏览添加更多页面")

    def _set_engines(self, engine_dict, value):
        for name in engine_dict:
            self.engine_vars[name].set(value)

    def _build_ai_extract_section(self):
        section = ttk.LabelFrame(self.frame, text=" AI 提取案例 ", padding=5)
        section.pack(fill="x", padx=5, pady=(5, 0))

        # -- Row 0: API key --
        key_bar = ttk.Frame(section)
        key_bar.pack(fill="x", pady=(0, 3))
        ttk.Label(key_bar, text="API 密钥：").pack(side="left")
        ttk.Label(key_bar, text="（搜索筛选 / 参考文献解析 · 默认 deepseek-chat）", foreground="#888", font=("", 8)).pack(side="left")
        self.ai_api_key_var = tk.StringVar(value=config.get_llm("search_parse")[2])
        api_entry = ttk.Entry(key_bar, textvariable=self.ai_api_key_var, width=42, show="*")
        api_entry.pack(side="left", padx=5)
        ttk.Button(key_bar, text="显示", command=lambda: api_entry.config(show="" if api_entry.cget("show") else "*")).pack(side="left")
        ttk.Label(key_bar, text="所有 API 配置见「设置 → API 配置」", foreground="#888", font=("", 8)).pack(side="right")

        # Extension hint (compact)
        self.ai_show_hint = tk.BooleanVar(value=True)
        ttk.Label(key_bar, text="需要安装浏览器扩展？", foreground="#888", font=("", 8), cursor="hand2").pack(side="left", padx=20)
        self.ai_hint_label = ttk.Label(section, foreground="#888", font=("", 8))
        self._update_ext_hint()

        # -- Row 1: toolbar --
        toolbar = ttk.Frame(section)
        toolbar.pack(fill="x", pady=(0, 5))
        self.ai_status_var = tk.StringVar(value="就绪 — 浏览器点扩展图标添加页面，可多次添加后点「开始提取」一键分析")
        ttk.Label(toolbar, textvariable=self.ai_status_var, foreground="gray").pack(side="left")
        ttk.Button(toolbar, text="保存到文献库", command=self._save_to_library).pack(side="right", padx=3)
        ttk.Button(toolbar, text="添加到数据库", command=self._ai_add_to_db).pack(side="right", padx=3)
        ttk.Button(toolbar, text="导出 Excel", command=self._ai_export).pack(side="right", padx=3)
        ttk.Button(toolbar, text="开始提取", command=self._ai_extract).pack(side="right", padx=3)

        # -- Row 2: text input (left) + results (right) --
        paned = ttk.PanedWindow(section, orient="horizontal", height=170)
        paned.pack(fill="x")

        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        header_frame = ttk.Frame(left_frame)
        header_frame.pack(fill="x")
        ttk.Label(header_frame, text="文章内容：", font=("", 9)).pack(side="left")
        ttk.Button(header_frame, text="清空", width=4, command=self._ai_clear_text).pack(side="right")
        self.ai_text_input = tk.Text(left_frame, wrap="word", height=8, font=("", 10))
        text_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.ai_text_input.yview)
        self.ai_text_input.configure(yscrollcommand=text_scroll.set)
        self.ai_text_input.pack(side="left", fill="both", expand=True)
        text_scroll.pack(side="right", fill="y")

        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=2)
        result_header = ttk.Frame(right_frame)
        result_header.pack(fill="x")
        ttk.Label(result_header, text="提取结果（选中后可导出或入库）：", font=("", 9)).pack(side="left")
        ttk.Button(result_header, text="清空", width=4, command=self._ai_clear_results).pack(side="right")
        columns = ("name", "country", "city", "source")
        self.ai_tree = ttk.Treeview(right_frame, columns=columns, show="headings", selectmode="extended", height=8)
        self.ai_tree.heading("name", text="建筑名称")
        self.ai_tree.heading("country", text="国家")
        self.ai_tree.heading("city", text="城市")
        self.ai_tree.heading("source", text="来源")
        self.ai_tree.column("name", width=180)
        self.ai_tree.column("country", width=80)
        self.ai_tree.column("city", width=80)
        self.ai_tree.column("source", width=150)
        ai_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.ai_tree.yview)
        self.ai_tree.configure(yscrollcommand=ai_scroll.set)
        self.ai_tree.pack(side="left", fill="both", expand=True)
        ai_scroll.pack(side="right", fill="y")

    def _update_ext_hint(self):
        self.ai_hint_label.config(
            text="Chrome → 扩展程序 → 开发者模式 → 加载已解压的扩展程序 → 选择程序目录下的 ext 文件夹 → 安装后点图标自动提取。")
        self.ai_hint_label.pack(fill="x", pady=(0, 3))
        self.ai_hint_label.bind("<Button-1>", lambda e: self.ai_hint_label.pack_forget())

    def _ai_clear(self):
        self._ai_clear_text()
        self._ai_clear_results()
        self.ai_status_var.set("已清空")

    def _ai_clear_text(self):
        self.ai_text_input.delete("1.0", "end")

    def _ai_clear_results(self):
        for item in self.ai_tree.get_children():
            self.ai_tree.delete(item)
        self.ai_results = []

    def _save_to_library(self):
        text = self.ai_text_input.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showinfo("提示", "文章内容为空，请先从浏览器扩展回传或手动输入内容。")
            return
        title = self._browser_last_title or "未命名文献"
        url = self._browser_last_url or ""
        # Check deduplication
        existing = db.search_literature_by_title(title)
        if existing and url:
            for e in existing:
                if e.get("source_url") == url:
                    if not messagebox.askyesno("重复", f"文献库中已存在「{title}」，是否更新内容？"):
                        return
                    db.update_literature(e["id"], {"content": text})
                    self.ai_status_var.set(f"已更新文献：{title}")
                    return
        # Generate summary
        api_key = self.ai_api_key_var.get().strip()
        summary = None
        if self._browser_last_description:
            from utils.ai_extractor import generate_summary_from_meta_description
            summary = generate_summary_from_meta_description(self._browser_last_description)
        if not summary and api_key:
            from utils.ai_extractor import generate_summary
            summary = generate_summary(text, api_key)
        lid = db.insert_literature({
            "title": title,
            "authors": self._browser_last_author or "",
            "pub_date": self._browser_last_pub_date or "",
            "source_url": url,
            "source_engine": "浏览器扩展",
            "content": text,
            "summary": summary,
        })
        if lid:
            self.ai_status_var.set(f"已保存到文献库：{title}")
            messagebox.showinfo("保存完成", f"文献「{title}」已保存到文献库。")

    def _open_literature_search(self):
        LiteratureSearchDialog(self.frame, search_tab=self)

    def _open_literature_library(self):
        api_key = self.ai_api_key_var.get().strip()
        LiteratureLibraryDialog(self.frame, api_key=api_key, search_tab=self)

    def _ai_extract(self):
        from utils.perf_monitor import Tracker
        text = self.ai_text_input.get("1.0", "end-1c").strip()
        if not text:
            return
        self._ai_clear_results()
        with Tracker(f"AI提取 ({len(text)}字符)"):
            self._do_ai_extract(text)

    def _do_ai_extract(self, text):
        api_key = self.ai_api_key_var.get().strip()
        if not api_key:
            messagebox.showinfo("提示", "请输入 API 密钥。")
            return

        import re
        # Split by article boundaries
        parts = re.split(r'---\s*\[来源\].*?---', text)
        articles = [p.strip() for p in parts if p.strip()]

        # Group articles into batches of ~8000 chars each
        BATCH_SIZE = 20000
        batches = []
        current_batch = ""
        for article in articles:
            if current_batch and len(current_batch) + len(article) > BATCH_SIZE:
                batches.append(current_batch.strip())
                current_batch = article
            else:
                if current_batch:
                    current_batch += "\n\n" + article
                else:
                    current_batch = article
        if current_batch.strip():
            batches.append(current_batch.strip())

        if len(batches) <= 1:
            # Single batch — process directly
            self.ai_status_var.set("正在分析中，请稍候...")
            self.frame.update()
            try:
                results = extract_case_names(text, api_key)
                if not isinstance(results, list):
                    results = []
                self.ai_results = results
            except Exception as e:
                messagebox.showerror("提取失败", f"API 调用出错：{e}")
                self.ai_status_var.set("提取失败")
                return
            self._display_ai_results()
            self.ai_status_var.set(f"提取完成，共 {len(self.ai_results)} 条案例名称")
        else:
            # Multiple batches — accumulate results
            self._batch_extract(batches, api_key)

    def _batch_extract(self, batches, api_key):
        """Process batches in background, accumulating results into self.ai_results."""
        import threading
        total = len(batches)
        done_flag = [False]
        error_holder = [None]
        progress = [0]
        accumulated = []  # thread-local accumulation, merged in _poll

        def _run():
            try:
                for i, batch in enumerate(batches):
                    results = extract_case_names(batch, api_key)
                    if isinstance(results, list):
                        accumulated.extend(results)
                    progress[0] = i + 1
            except Exception as e:
                error_holder[0] = str(e)
            done_flag[0] = True

        threading.Thread(target=_run, daemon=True).start()

        def _poll():
            if done_flag[0]:
                if error_holder[0]:
                    messagebox.showerror("提取失败", f"API 调用出错：{error_holder[0]}")
                    self.ai_status_var.set("提取失败")
                else:
                    self.ai_results = accumulated
                    self._display_ai_results()
                    self.ai_status_var.set(f"提取完成，{total}批次共 {len(accumulated)} 条案例")
            else:
                self.ai_status_var.set(f"正在分析 {progress[0]+1}/{total} 批次...")
                self.frame.after(300, _poll)

        self.ai_status_var.set(f"正在分析 1/{total} 批次...")
        self.frame.after(300, _poll)

    def _display_ai_results(self):
        """Refresh the AI results treeview with current self.ai_results."""
        for item in self.ai_tree.get_children():
            self.ai_tree.delete(item)
        for r in self.ai_results:
            if isinstance(r, dict) and r.get("name"):
                self.ai_tree.insert("", "end", values=(
                    r.get("name", ""),
                    r.get("country", "") or "",
                    r.get("city", "") or "",
                    r.get("source_url", "") or "",
                ))

    def _ai_export(self):
        if not self.ai_results:
            messagebox.showinfo("提示", "暂无提取结果。")
            return
        data = []
        for i, r in enumerate(self.ai_results, 1):
            if isinstance(r, dict):
                data.append({
                    "序号": i,
                    "建筑名称": r.get("name", ""),
                    "国家": r.get("country", "") or "",
                    "城市": r.get("city", "") or "",
                    "来源链接": r.get("source_url", "") or "",
                })
        df = pd.DataFrame(data)
        if export_to_excel(df, "导出案例名称"):
            messagebox.showinfo("导出完成", f"已导出 {len(data)} 条案例名称。")

    def _ai_add_to_db(self):
        if not self.ai_results:
            messagebox.showinfo("提示", "暂无提取结果。")
            return
        from utils.perf_monitor import Tracker
        added = 0
        lookup_count = 0
        with Tracker(f"添加到数据库 ({len(self.ai_results)}条)"):
            for r in self.ai_results:
                if isinstance(r, dict) and r.get("name"):
                    # Look up search keywords from main search results by URL only
                    search_kw = None
                    source_url = r.get("source_url")
                    source_urls = r.get("source_urls") or ([source_url] if source_url else [])
                    for u in source_urls:
                        if u:
                            looked_up = db.lookup_search_keywords_for_url(u)
                            if looked_up:
                                search_kw = looked_up
                                lookup_count += 1
                                break

                    data = {
                        "name": r.get("name", ""),
                        "name_en": r.get("name_en"),
                        "name_official": r.get("name_official"),
                        "country": r.get("country"),
                        "city": r.get("city"),
                        "construction_year": r.get("construction_year"),
                        "demolition_year": r.get("demolition_year"),
                        "restoration_year": r.get("restoration_year"),
                        "quantity": r.get("quantity", 1),
                        "function": r.get("function"),
                        "designer": r.get("designer"),
                        "latitude": r.get("latitude"),
                        "longitude": r.get("longitude"),
                        "search_keywords": json.dumps(search_kw, ensure_ascii=False) if search_kw else None,
                        "image_urls": json.dumps(r.get("image_urls"), ensure_ascii=False) if r.get("image_urls") else None,
                        "source_urls": json.dumps(source_urls, ensure_ascii=False) if source_urls else None,
                    }
                    try:
                        raw_id = db.insert_raw_case(data)
                        if source_url:
                            db.insert_case_source(raw_id, source_url)
                        elif isinstance(r.get("source_urls"), list):
                            for u in r["source_urls"]:
                                db.insert_case_source(raw_id, u)
                        added += 1
                    except Exception:
                        pass
        messagebox.showinfo("完成", f"已添加 {added} 条原始案例到数据库。请在「数据整理→原始数据」中查看。")
        if added > 0:
            self.refresh()

    def _build_local_search_bar(self):
        bar = ttk.LabelFrame(self.frame, text=" 库内搜索 ", padding=5)
        bar.pack(fill="x", padx=5, pady=(5, 0))

        ttk.Label(bar, text="关键词：").pack(side="left")
        self.local_search_var = tk.StringVar()
        local_entry = ttk.Entry(bar, textvariable=self.local_search_var, width=35)
        local_entry.pack(side="left", padx=5)
        local_entry.bind("<Return>", lambda e: self.refresh())

        ttk.Button(bar, text="搜索", command=self.refresh).pack(side="left", padx=2)

        ttk.Label(bar, text="  类型：").pack(side="left", padx=(15, 0))
        self.type_var = tk.StringVar()
        self.type_combo = ttk.Combobox(bar, textvariable=self.type_var, width=12, state="readonly")
        self.type_combo["values"] = [""] + db.BUILDING_TYPES
        self.type_combo.pack(side="left", padx=5)

        ttk.Label(bar, text="  国家：").pack(side="left", padx=(10, 0))
        self.country_var = tk.StringVar()
        self.country_combo = ttk.Combobox(bar, textvariable=self.country_var, width=12)
        self.country_combo.pack(side="left", padx=5)
        self._load_countries()

        ttk.Button(bar, text="筛选", command=self.refresh).pack(side="left", padx=5)
        ttk.Button(bar, text="清除筛选", command=self._clear_filter).pack(side="left")

        ttk.Button(bar, text="查看详情", command=self._show_detail).pack(side="right", padx=5)
        ttk.Button(bar, text="刷新", command=self.refresh).pack(side="right", padx=2)
        ttk.Button(bar, text="删除选中", command=self._delete_selected).pack(side="right", padx=2)

    def _open_keyword_gen(self):
        KeywordGenDialog(self.frame, callback=self._on_keywords_selected)

    def _on_keywords_selected(self, keywords):
        self.kw_listbox.delete(0, tk.END)
        for kw in keywords:
            self.kw_listbox.insert(tk.END, kw)
        self.kw_listbox.select_set(0, tk.END)
        if keywords:
            self.web_search_var.set(keywords[0])

    def _web_search(self):
        selected_indices = self.kw_listbox.curselection()
        if selected_indices:
            queries = [self.kw_listbox.get(i) for i in selected_indices]
        else:
            query = self.web_search_var.get().strip()
            if not query:
                messagebox.showinfo("提示", "请输入搜索关键词或从下方列表选择。")
                return
            queries = [query]

        selected_engines = [name for name, var in self.engine_vars.items() if var.get()]
        if not selected_engines:
            messagebox.showinfo("提示", "请至少选择一个搜索引擎。")
            return

        try:
            chrome = webbrowser.get('chrome')
        except webbrowser.Error:
            chrome = None

        for query in queries:
            encoded = urllib.parse.quote(query)
            for name in selected_engines:
                url = ALL_ENGINES[name].format(encoded)
                if chrome:
                    chrome.open(url)
                else:
                    webbrowser.open(url)

    def _serper_search(self):
        """Open the Serper search dialog."""
        keyword = self.web_search_var.get().strip()
        SerperDialog(self.frame, keyword=keyword, search_tab=self)

    def _build_tree(self):
        tree_frame = ttk.Frame(self.frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=2)

        columns = ("id", "name", "country", "city", "year_built", "building_type", "architect",
                   "condition")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", selectmode="extended")

        self.tree.heading("#0", text="图片")
        self.tree.column("#0", width=100, minwidth=80)

        col_headers = {"id": "ID", "name": "建筑名称", "country": "国家", "city": "城市",
                       "year_built": "建成年代", "building_type": "建筑类型", "architect": "建筑师",
                       "condition": "建筑状态"}
        col_widths = {"id": 50, "name": 180, "country": 90, "city": 90,
                      "year_built": 80, "building_type": 90, "architect": 120,
                      "condition": 70}

        for col in columns:
            self.tree.heading(col, text=col_headers[col], command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=col_widths.get(col, 100))

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda e: self._show_detail())

    def _build_context_menu(self):
        self.ctx_menu = tk.Menu(self.frame, tearoff=0)
        self.ctx_menu.add_command(label="查看详情", command=self._show_detail)
        self.ctx_menu.add_command(label="删除", command=self._delete_selected)
        self.tree.bind("<Button-2>", self._popup_context)
        self.tree.bind("<Button-3>", self._popup_context)
        self.tree.bind("<Control-2>", self._popup_context)

    def _popup_context(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.ctx_menu.post(event.x_root, event.y_root)

    def _sort_by(self, col):
        rows = [(self.tree.set(item, col), item) for item in self.tree.get_children("")]
        reverse = getattr(self, "_sort_reverse", False)
        try:
            rows.sort(key=lambda x: float(x[0]) if x[0].replace(".", "").isdigit() else x[0], reverse=reverse)
        except (ValueError, AttributeError, TypeError):
            rows.sort(key=lambda x: x[0], reverse=reverse)
        for i, (_, item) in enumerate(rows):
            self.tree.move(item, "", i)
        self._sort_reverse = not reverse

    def _load_countries(self):
        cases = db.get_all_cases()
        countries = sorted(set(c["country"] for c in cases if c.get("country")))
        self.country_combo["values"] = [""] + countries

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._photo_cache.clear()

        filters = {}
        if self.type_var.get():
            filters["building_type"] = self.type_var.get()
        if self.country_var.get():
            filters["country"] = self.country_var.get()

        cases = db.get_all_cases(search=self.local_search_var.get() or None, filters=filters if filters else None)
        for c in cases:
            image_urls_raw = c.get("image_urls") or ""
            image_path = c.get("image_path")
            urls = parse_image_urls(image_urls_raw, image_path)
            count_text = f"[{len(urls)}张]" if urls else "—"

            item_id = self.tree.insert("", "end", text=count_text, values=(
                c["id"], c["name"] or "", c["country"] or "", c["city"] or "",
                c["year_built"] or "", c["building_type"] or "", c["architect"] or "",
                c.get("condition") or "",
            ))

            if urls:
                threading.Thread(
                    target=load_tree_thumbnail,
                    args=(self.tree, item_id, urls[0], c["id"], self._photo_cache),
                    daemon=True
                ).start()

    def _clear_filter(self):
        self.local_search_var.set("")
        self.type_var.set("")
        self.country_var.set("")
        self.refresh()

    def _show_detail(self):
        case_id = self._get_selected_id()
        if not case_id:
            return
        case = db.get_case_by_id(case_id)
        if not case:
            return
        DetailDialog(self.frame, case)

    def _delete_selected(self):
        items = self.tree.selection()
        if not items:
            messagebox.showinfo("提示", "请先选择要删除的记录。")
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(items)} 条记录吗？"):
            for item in items:
                case_id = self.tree.item(item, "values")[0]
                db.delete_case(int(case_id))
            self.refresh()

    def _get_selected_id(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return None
        return int(self.tree.item(sel[0], "values")[0])


class DetailDialog:
    def __init__(self, parent, case):
        self.case = case
        self.win = tk.Toplevel(parent)
        self.win.title(f"案例详情 - {case.get('name', '')}")
        self.win.geometry("750x850")
        self.win.transient(parent)

        canvas = tk.Canvas(self.win)
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=720)
        canvas.configure(yscrollcommand=scrollbar.set)

        # ---- Image section ----
        img_section = ttk.LabelFrame(scroll_frame, text=" 建筑图片（点击放大） ", padding=5)
        img_section.pack(fill="x", padx=10, pady=(10, 0))
        ImageThumbnailStrip(img_section,
                             case.get("image_urls", ""),
                             case.get("image_path"))

        # ---- Info fields ----
        info_frame = ttk.Frame(scroll_frame)
        info_frame.pack(fill="x", padx=10, pady=(8, 0))

        fields_display = [(f, l) for f, l, *_ in db.FIELDS] + [
            ("created_at", "创建时间"), ("updated_at", "更新时间")
        ]
        for i, (field, label) in enumerate(fields_display):
            ttk.Label(info_frame, text=label + "：", font=("", 10, "bold")).grid(
                row=i, column=0, sticky="ne", padx=(5, 5), pady=2)
            value = case.get(field, "")
            if value is None:
                value = ""
            val_label = ttk.Label(info_frame, text=str(value), wraplength=480, justify="left")
            val_label.grid(row=i, column=1, sticky="w", padx=(0, 5), pady=2)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.win)
        btn_frame.pack(pady=8)
        ttk.Button(btn_frame, text="编辑", command=self._edit).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="关闭", command=self.win.destroy).pack(side="left", padx=5)

    def _edit(self):
        dlg = EditDialog(self.win, self.case)
        self.win.wait_window(dlg.win)
        # Refresh detail after edit
        updated = db.get_case_by_id(self.case.get("id"))
        if updated:
            self.win.destroy()
            DetailDialog(self.win.master.master, updated)


class EditDialog:
    def __init__(self, parent, case):
        self.case = case
        self.win = tk.Toplevel(parent)
        self.win.title(f"编辑案例 - {case.get('name', '')}")
        self.win.geometry("700x850")
        self.win.transient(parent)
        self.win.grab_set()

        canvas = tk.Canvas(self.win)
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=680)
        canvas.configure(yscrollcommand=scrollbar.set)

        # ---- Image section ----
        img_section = ttk.LabelFrame(scroll_frame, text=" 建筑图片（点击放大） ", padding=5)
        img_section.pack(fill="x", padx=10, pady=(10, 0))
        ImageThumbnailStrip(img_section,
                             case.get("image_urls", ""),
                             case.get("image_path"))

        # ---- Form fields ----
        form_frame = ttk.Frame(scroll_frame)
        form_frame.pack(fill="x", padx=10, pady=(8, 0))

        self.entries = {}
        for i, (field, label, *_t) in enumerate(db.FIELDS):
            ttk.Label(form_frame, text=label + "：").grid(row=i, column=0, sticky="ne", padx=(5, 5), pady=2)
            var = tk.StringVar(value=case.get(field, "") or "")
            if field == "building_type":
                w = ttk.Combobox(form_frame, textvariable=var, values=db.BUILDING_TYPES, state="readonly", width=38)
            elif field == "condition":
                w = ttk.Combobox(form_frame, textvariable=var, values=db.CONDITION_OPTIONS, state="readonly", width=38)
            elif field == "description" or field == "notes":
                w = tk.Text(form_frame, width=40, height=4)
                w.insert("1.0", var.get())
                self.entries[field] = ("text", w)
                w.grid(row=i, column=1, sticky="w", padx=(0, 15), pady=3)
                continue
            elif field == "area_sqm":
                w = ttk.Entry(form_frame, textvariable=var, width=40)
            else:
                w = ttk.Entry(form_frame, textvariable=var, width=40)
            self.entries[field] = ("entry", var)
            w.grid(row=i, column=1, sticky="w", padx=(0, 15), pady=3)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(self.win)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="保存", command=self._save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=self.win.destroy).pack(side="left", padx=5)

    def _save(self):
        data = {}
        for field, (wtype, widget) in self.entries.items():
            if wtype == "text":
                val = widget.get("1.0", "end-1c").strip()
            else:
                val = widget.get().strip()
            data[field] = val if val else None

        for num_field in ["area_sqm", "latitude", "longitude"]:
            try:
                if data.get(num_field):
                    data[num_field] = float(data[num_field])
                else:
                    data[num_field] = None
            except ValueError:
                messagebox.showerror("错误", f"{num_field} 请输入数字。")
                return

        try:
            if data.get("quantity"):
                data["quantity"] = int(data["quantity"])
            else:
                data["quantity"] = 1
        except ValueError:
            messagebox.showerror("错误", "数量请输入整数。")
            return

        if self.case.get("id"):
            db.update_case(self.case["id"], data)
        else:
            db.insert_case(data)
        messagebox.showinfo("完成", "案例已保存。")
        self.win.destroy()


class LiteratureEditDialog:
    def __init__(self, parent, literature, on_save=None):
        self.lit = literature
        self.on_save = on_save
        self._file_var = tk.StringVar(value=literature.get("file_path", "") or "")
        self.win = tk.Toplevel(parent)
        self.win.title(f"编辑文献 - {literature.get('title', '')}")
        self.win.geometry("600x580")
        self.win.transient(parent)
        self.win.grab_set()

        frame = ttk.Frame(self.win, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="标题：").grid(row=0, column=0, sticky="ne", pady=2)
        self.title_var = tk.StringVar(value=literature.get("title", ""))
        ttk.Entry(frame, textvariable=self.title_var, width=50).grid(row=0, column=1, sticky="w", padx=5)

        ttk.Label(frame, text="作者：").grid(row=1, column=0, sticky="ne", pady=2)
        self.authors_var = tk.StringVar(value=literature.get("authors", "") or "")
        ttk.Entry(frame, textvariable=self.authors_var, width=50).grid(row=1, column=1, sticky="w", padx=5)

        ttk.Label(frame, text="发表日期：").grid(row=2, column=0, sticky="ne", pady=2)
        self.pub_date_var = tk.StringVar(value=literature.get("pub_date", "") or "")
        ttk.Entry(frame, textvariable=self.pub_date_var, width=50).grid(row=2, column=1, sticky="w", padx=5)

        ttk.Label(frame, text="来源URL：").grid(row=3, column=0, sticky="ne", pady=2)
        ttk.Label(frame, text=literature.get("source_url", "") or "—", foreground="gray").grid(
            row=3, column=1, sticky="w", padx=5)

        ttk.Label(frame, text="来源引擎：").grid(row=4, column=0, sticky="ne", pady=2)
        ttk.Label(frame, text=literature.get("source_engine", "") or "—", foreground="gray").grid(
            row=4, column=1, sticky="w", padx=5)

        ttk.Label(frame, text="附件文件：").grid(row=5, column=0, sticky="ne", pady=2)
        file_row = ttk.Frame(frame)
        file_row.grid(row=5, column=1, sticky="w", padx=5)
        file_entry = ttk.Entry(file_row, textvariable=self._file_var, font=("", 8),
                               state="readonly", width=35)
        file_entry.pack(side="left")
        ttk.Button(file_row, text="上传文件", command=self._attach_file).pack(side="left", padx=3)
        if self._file_var.get():
            ttk.Button(file_row, text="打开", command=self._open_attached).pack(side="left")
        ttk.Button(file_row, text="清除", command=lambda: self._file_var.set("")).pack(side="left", padx=1)

        ttk.Label(frame, text="摘要：").grid(row=6, column=0, sticky="ne", pady=(10, 2))
        self.summary_text = tk.Text(frame, wrap="word", font=("", 10), height=5, width=45)
        self.summary_text.grid(row=6, column=1, sticky="w", padx=5, pady=(10, 2))
        self.summary_text.insert("1.0", literature.get("summary", "") or "")

        ttk.Label(frame, text="备注：").grid(row=7, column=0, sticky="ne", pady=2)
        self.notes_text = tk.Text(frame, wrap="word", font=("", 10), height=3, width=45)
        self.notes_text.grid(row=7, column=1, sticky="w", padx=5)
        self.notes_text.insert("1.0", literature.get("notes", "") or "")

        btn_bar = ttk.Frame(frame)
        btn_bar.grid(row=8, column=1, sticky="e", pady=(10, 0))
        ttk.Button(btn_bar, text="保存", command=self._save).pack(side="left", padx=5)
        ttk.Button(btn_bar, text="取消", command=self.win.destroy).pack(side="left", padx=5)

    def _attach_file(self):
        path = filedialog.askopenfilename(filetypes=[
            ("All supported", "*.pdf *.docx *.doc *.txt *.md *.csv *.xlsx *.pptx"),
            ("PDF files", "*.pdf"),
            ("Word documents", "*.docx *.doc"),
            ("Text files", "*.txt *.md *.csv"),
            ("All files", "*.*"),
        ])
        if not path:
            return
        import shutil
        import uuid
        ext = os.path.splitext(path)[1].lower()
        lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
        os.makedirs(lit_dir, exist_ok=True)
        dest_name = f"{uuid.uuid4().hex}{ext}"
        shutil.copy2(path, os.path.join(lit_dir, dest_name))
        self._file_var.set(dest_name)
        messagebox.showinfo("完成", "文件已附加。保存后生效。", parent=self.win)

    def _open_attached(self):
        lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
        os.makedirs(lit_dir, exist_ok=True)
        file_path = self._file_var.get().strip()
        full_path = os.path.join(lit_dir, file_path) if file_path else ""
        import subprocess
        import platform
        try:
            if platform.system() == "Darwin":
                if full_path and os.path.exists(full_path):
                    subprocess.run(["open", "-R", full_path])
                else:
                    subprocess.run(["open", lit_dir])
            elif platform.system() == "Windows":
                if full_path and os.path.exists(full_path):
                    subprocess.run(["explorer", "/select,", full_path])
                else:
                    os.startfile(lit_dir)
            else:
                if full_path and os.path.exists(full_path):
                    subprocess.run(["xdg-open", os.path.dirname(full_path)])
                else:
                    subprocess.run(["xdg-open", lit_dir])
        except Exception as e:
            messagebox.showerror("打开失败", str(e), parent=self.win)

    def _save(self):
        data = {
            "title": self.title_var.get().strip(),
            "authors": self.authors_var.get().strip(),
            "pub_date": self.pub_date_var.get().strip(),
            "summary": self.summary_text.get("1.0", "end-1c").strip(),
            "notes": self.notes_text.get("1.0", "end-1c").strip(),
            "file_path": self._file_var.get().strip(),
        }
        db.update_literature(self.lit["id"], data)
        if self.on_save:
            self.on_save()
        self.win.destroy()


class SerperDialog:
    """Independent dialog for Serper.dev Google search + AI filtering.
    Supports single search and batch queue search — results accumulate."""

    def __init__(self, parent, keyword="", search_tab=None):
        self._search_tab = search_tab
        self._api_key = config.get_llm("search_parse")[2]
        self._results = []
        self._query_groups = {}  # query_label -> [results]
        self._idx_counter = 0
        self._queue_running = False
        self._queue_paused = False
        self._queue_stop = False
        self._removed_info = []  # (removed_item, kept_item, reason)
        self._undo_stack = []    # [{type, items, ...}, ...] max 10
        self._status_dlg = None
        self._use_serper = tk.BooleanVar(value=False)
        self._use_baidu = tk.BooleanVar(value=False)

        self.win = tk.Toplevel(parent)
        self.win.title("Serper 搜索")
        self.win.geometry("1050x680")
        self.win.transient(parent)
        self._build_ui()

        self._load_from_db()

        # Check for saved queue checkpoint (crash recovery)
        saved = self._load_checkpoint()
        if saved and saved.get("keywords"):
            remaining = saved["keywords"][saved.get("current_idx", 0):]
            if remaining:
                if messagebox.askyesno("恢复队列",
                    f"检测到未完成的队列任务。\n"
                    f"已完成: {saved.get('current_idx', 0)}/{len(saved['keywords'])}个词\n"
                    f"剩余: {len(remaining)}个词\n\n是否恢复到队列？"):
                    for kw in remaining:
                        self._queue_listbox.insert("end", kw)
                    self._status_var.set(f"已恢复 {len(remaining)} 个待处理关键词")

        if keyword:
            self._keyword_var.set(keyword)
            self.win.after(100, self._do_search)

    def _build_ui(self):
        toolbar = ttk.Frame(self.win, padding=5)
        toolbar.pack(fill="x")

        ttk.Label(toolbar, text="关键词：").pack(side="left")
        self._keyword_var = tk.StringVar()
        kw_entry = ttk.Entry(toolbar, textvariable=self._keyword_var, width=28)
        kw_entry.pack(side="left", padx=5)
        kw_entry.bind("<Return>", lambda e: self._do_search())
        ttk.Button(toolbar, text="生成关键词",
                   command=self._open_keyword_gen).pack(side="left", padx=2)
        ttk.Button(toolbar, text="serper搜索", command=self._do_search).pack(side="left", padx=2)
        ttk.Button(toolbar, text="百度搜索api", command=self._do_baidu_search).pack(side="left", padx=2)
        ttk.Button(toolbar, text="队列日志", command=self._open_log_viewer).pack(side="left", padx=2)

        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(toolbar, textvariable=self._status_var).pack(side="right")

        # Stats + action bar
        self._stats_frame = ttk.Frame(self.win, padding=(5, 0))
        self._stats_frame.pack(fill="x")
        self._stats_label = ttk.Label(self._stats_frame, text="")
        self._stats_label.pack(side="left")
        ttk.Button(self._stats_frame, text="发送到文献库",
                   command=self._send_to_library).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="原始结果",
                   command=self._show_raw_results).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="刷新",
                   command=self._refresh_from_db).pack(side="right", padx=2)
        self._expand_btn = ttk.Button(self._stats_frame, text="展开/合上",
                   command=self._expand_selected)
        self._expand_btn.pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="在浏览器打开选中",
                   command=self._open_selected).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="数据清洗",
                   command=self._delete_selected).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="简单去重",
                   command=self._deduplicate).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="模糊去重",
                   command=self._fuzzy_dedup_selected).pack(side="right", padx=2)
        self._undo_btn = ttk.Button(self._stats_frame, text="撤销",
                                     command=self._undo_last, state="disabled")
        self._undo_btn.pack(side="right", padx=2)
        self._undo_log_var = tk.StringVar(value="")
        self._undo_log_label = ttk.Label(self._stats_frame, textvariable=self._undo_log_var,
                                          font=("", 7), foreground="gray")
        self._undo_log_label.pack(side="right", padx=(0, 8))

        # Filter bar
        filter_frame = ttk.Frame(self.win, padding=(5, 2))
        filter_frame.pack(fill="x")
        ttk.Label(filter_frame, text="按检索词筛选：").pack(side="left")
        self._filter_query_var = tk.StringVar(value="全部")
        self._filter_query_combo = ttk.Combobox(filter_frame, textvariable=self._filter_query_var,
                                                 state="readonly", width=18)
        self._filter_query_combo.pack(side="left", padx=2)
        self._filter_query_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        ttk.Button(filter_frame, text="清除筛选", command=self._clear_filter, width=6).pack(side="left", padx=5)
        ttk.Label(filter_frame, text="搜索：").pack(side="left")
        self._filter_search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_frame, textvariable=self._filter_search_var, width=18)
        search_entry.pack(side="left", padx=2)
        self._filter_search_var.trace_add("write", lambda *_: self._apply_filter())

        # Main area: queue panel (left) + treeview (right)
        main_paned = ttk.PanedWindow(self.win, orient="horizontal")
        main_paned.pack(fill="both", expand=True, padx=5, pady=(5, 0))

        # -- Queue panel --
        queue_frame = ttk.LabelFrame(main_paned, text=" 关键词队列 ", padding=5)
        main_paned.add(queue_frame, weight=0)

        listbox_frame = ttk.Frame(queue_frame)
        listbox_frame.pack(fill="both", expand=True)
        self._queue_listbox = tk.Listbox(listbox_frame, width=20, height=12,
                                          selectmode="extended", exportselection=False)
        self._queue_listbox.pack(side="left", fill="both", expand=True)
        q_scroll = ttk.Scrollbar(listbox_frame, orient="vertical", command=self._queue_listbox.yview)
        self._queue_listbox.configure(yscrollcommand=q_scroll.set)
        q_scroll.pack(side="right", fill="y")

        qbtn_frame = ttk.Frame(queue_frame)
        qbtn_frame.pack(fill="x", pady=(3, 0))
        ttk.Button(qbtn_frame, text="添加", command=self._queue_add, width=5).pack(side="left", padx=1)
        ttk.Button(qbtn_frame, text="删除", command=self._queue_del, width=5).pack(side="left", padx=1)
        ttk.Button(qbtn_frame, text="清空", command=self._queue_clear, width=5).pack(side="left", padx=1)

        # Engine selection
        eng_frame = ttk.Frame(queue_frame)
        eng_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(eng_frame, text="发送至：").pack(side="left")
        ttk.Checkbutton(eng_frame, text="Serper", variable=self._use_serper).pack(side="left", padx=1)
        ttk.Checkbutton(eng_frame, text="百度", variable=self._use_baidu).pack(side="left", padx=1)

        qctrl_frame = ttk.Frame(queue_frame)
        qctrl_frame.pack(fill="x", pady=(5, 0))
        self._queue_start_btn = ttk.Button(qctrl_frame, text="排队搜索",
                                            command=self._queue_start)
        self._queue_start_btn.pack(fill="x")
        self._queue_pause_btn = ttk.Button(qctrl_frame, text="暂停",
                                            command=self._queue_pause, state="disabled")
        self._queue_pause_btn.pack(fill="x", pady=(2, 0))
        ttk.Button(qctrl_frame, text="停止", command=self._queue_stop_search).pack(fill="x", pady=(2, 0))

        self._queue_progress = ttk.Progressbar(queue_frame, mode="determinate")
        self._queue_progress.pack(fill="x", pady=(5, 0))
        self._queue_progress_var = tk.StringVar(value="")
        ttk.Label(queue_frame, textvariable=self._queue_progress_var,
                  font=("", 8)).pack()

        # -- Treeview area --
        tree_panel = ttk.Frame(main_paned)
        main_paned.add(tree_panel, weight=1)

        # Treeview
        tree_frame = ttk.Frame(tree_panel)
        tree_frame.pack(fill="both", expand=True)

        self._sort_col = None
        self._sort_desc = False

        columns = ("query", "title", "reason")
        self._tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings",
                                   selectmode="extended")
        self._tree.heading("#0", text="搜索词组 / #")
        self._tree.column("#0", width=200, minwidth=150)
        self._tree.heading("query", text="搜索词",
                           command=lambda: self._sort_by("query"))
        self._tree.column("query", width=0, minwidth=0, stretch=False)
        self._tree.heading("title", text="标题",
                           command=lambda: self._sort_by("title"))
        self._tree.column("title", width=350, minwidth=200)
        self._tree.heading("reason", text="AI 筛选理由",
                           command=lambda: self._sort_by("reason"))
        self._tree.column("reason", width=350, minwidth=200)

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        _make_tree_draggable(self._tree)

        # Detail area
        detail_frame = ttk.LabelFrame(self.win, text=" 选中详情 ", padding=5)
        detail_frame.pack(fill="x", padx=5, pady=(5, 0))

        self._detail_text = tk.Text(detail_frame, height=4, wrap="word", state="disabled")
        self._detail_text.pack(fill="both")

    def _add_batch_to_tree(self, query_label, results):
        """Add results under a parent query node. Creates parent if new."""
        if query_label not in self._query_groups:
            self._query_groups[query_label] = []
            parent_iid = f"G_{query_label}"
            self._tree.insert("", "end", iid=parent_iid, open=False,
                              text=f"  {query_label}", values=(query_label, "", ""))
        else:
            parent_iid = f"G_{query_label}"
        for r in results:
            self._idx_counter += 1
            r["_query"] = query_label
            r["_idx"] = self._idx_counter
            self._results.append(r)
            self._query_groups[query_label].append(r)
            self._tree.insert(parent_iid, "end", iid=str(self._idx_counter),
                              text=str(self._idx_counter),
                              values=("", r["title"],
                                      r.get("relevance_reason", "")))
        # Update parent text with count
        count = len(self._query_groups[query_label])
        self._tree.item(parent_iid, text=f"  {query_label} ({count}条)")

    def _do_search(self):
        query = self._keyword_var.get().strip()
        if not query:
            messagebox.showinfo("提示", "请输入搜索关键词")
            return

        self._status_var.set(f"搜索中: {query} ...")
        self.win.update()

        from utils.smart_search import smart_search
        try:
            new_results, stats = smart_search(query, api_key=self._api_key)
        except Exception as e:
            err_msg = str(e)
            if "timeout" in err_msg.lower() or "timed out" in err_msg.lower():
                messagebox.showerror("Serper搜索超时",
                    f"Serper API 3分钟未响应，已超时。\n"
                    f"可能原因：代理未启动、网络不稳定、或API额度已用完。\n\n"
                    f"建议：检查代理({PROXY})是否运行，或稍后重试。")
            elif "Connection refused" in err_msg or "Errno 61" in err_msg:
                messagebox.showerror("Serper连接失败",
                    f"Serper代理({PROXY})连接被拒绝。\n请确认代理是否已启动。")
            else:
                messagebox.showerror("Serper搜索失败", f"{err_msg}\n\n可能原因：网络问题或API额度已用完。")
            self._status_var.set("搜索失败")
            return

        # Accumulate (grouped under parent node)
        self._add_batch_to_tree(query, new_results)

        # Save to database
        try:
            import database as db
            db.insert_filtered_results(query, new_results)
        except Exception as e:
            print(f"[SerperDialog] save db error: {e}")

        # Stats for this search + total
        total_relevant = len(self._results)
        self._stats_label.config(text=(
            f"本次: {query} → 翻{stats['pages_fetched']}页/{stats['total_raw']}条 → "
            f"相关{len(new_results)}条 | "
            f"累计: {total_relevant} 条"
        ))
        self._status_var.set(f"「{query}」完成，+{len(new_results)}条（累计 {total_relevant} 条）")
        self._update_filter_queries()

    def _do_baidu_search(self):
        query = self._keyword_var.get().strip()
        if not query:
            messagebox.showinfo("提示", "请输入搜索关键词")
            return

        self._status_var.set(f"百度搜索中: {query} ...")
        self.win.update()

        from utils.smart_search import baidu_smart_search
        try:
            new_results, stats = baidu_smart_search(query, api_key=self._api_key)
        except Exception as e:
            err_msg = str(e)
            if "timeout" in err_msg.lower() or "timed out" in err_msg.lower():
                messagebox.showerror("百度搜索超时",
                    f"百度API 3分钟未响应，已超时。\n可能网络不稳定或API不可用，请稍后重试。")
            else:
                messagebox.showerror("百度搜索失败", f"{err_msg}\n\n可能原因：网络问题或API不可用。")
            self._status_var.set("百度搜索失败")
            return

        self._add_batch_to_tree(f"[百度]{query}", new_results)

        try:
            import database as db
            db.insert_filtered_results(f"[百度]{query}", new_results)
        except Exception as e:
            print(f"[SerperDialog] save baidu db error: {e}")

        total_relevant = len(self._results)
        self._stats_label.config(text=(
            f"本次: [百度]{query} → {stats['total_raw']}条 → "
            f"相关{len(new_results)}条 | "
            f"累计: {total_relevant} 条"
        ))
        self._status_var.set(f"「[百度]{query}」完成，+{len(new_results)}条（累计 {total_relevant} 条）")
        self._update_filter_queries()

    def _open_log_viewer(self):
        """Open a log viewer dialog showing historical search logs."""
        import os
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
        log_files = sorted(
            [f for f in os.listdir(log_dir) if f.endswith(".log")],
            reverse=True
        ) if os.path.isdir(log_dir) else []

        dlg = tk.Toplevel(self.win)
        dlg.title("搜索日志")
        dlg.geometry("650x500")
        dlg.transient(self.win)

        # File selector
        sel_frame = ttk.Frame(dlg, padding=5)
        sel_frame.pack(fill="x")
        ttk.Label(sel_frame, text="日志文件：").pack(side="left")
        file_var = tk.StringVar()
        file_combo = ttk.Combobox(sel_frame, textvariable=file_var,
                                  values=log_files, state="readonly", width=50)
        file_combo.pack(side="left", padx=5)
        if log_files:
            file_combo.current(0)

        def _refresh():
            nonlocal log_files
            log_files = sorted(
                [f for f in os.listdir(log_dir) if f.endswith(".log")],
                reverse=True
            ) if os.path.isdir(log_dir) else []
            file_combo["values"] = log_files
            if log_files:
                file_combo.current(0)
                _load()

        def _load():
            fname = file_var.get()
            if not fname:
                return
            fpath = os.path.join(log_dir, fname)
            log_text.configure(state="normal")
            log_text.delete("1.0", "end")
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    content = f.read()
                log_text.insert("1.0", content)
            except Exception as e:
                log_text.insert("1.0", f"读取日志失败: {e}")
            log_text.configure(state="disabled")
            log_text.see("end")

        file_combo.bind("<<ComboboxSelected>>", lambda e: _load())
        ttk.Button(sel_frame, text="刷新", command=_refresh).pack(side="left", padx=2)
        ttk.Button(sel_frame, text="补跑失败词", command=lambda: self._retry_from_log(file_var.get())).pack(side="left", padx=2)

        # Log content
        log_frame = ttk.Frame(dlg, padding=5)
        log_frame.pack(fill="both", expand=True)
        log_text = tk.Text(log_frame, wrap="word", font=("", 9))
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
        log_text.configure(yscrollcommand=log_scroll.set)
        log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        ttk.Button(dlg, text="关闭", command=dlg.destroy).pack(pady=5)

        if log_files:
            _load()
        else:
            log_text.configure(state="normal")
            log_text.insert("1.0", "暂无日志文件。\n\n队列搜索的日志会自动保存到 data/logs/ 目录。")
            log_text.configure(state="disabled")

    def _retry_from_log(self, log_filename):
        """Parse failed keywords from log file and add them back to queue."""
        if not log_filename:
            return
        import os, re as _re
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
        fpath = os.path.join(log_dir, log_filename)
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception:
            messagebox.showerror("错误", f"无法读取日志文件: {log_filename}")
            return
        failed = []
        for line in content.split("\n"):
            if "[错误]" in line or "[跳过]" in line:
                m = _re.search(r'\] \[(?:Scholar|Serper|百度)\] (.+?):', line)
                if m:
                    kw = m.group(1).strip()
                    if kw not in failed:
                        failed.append(kw)
        if not failed:
            messagebox.showinfo("提示", "该日志中没有找到失败或跳过的关键词。")
            return
        existing = set(self._queue_listbox.get(0, "end"))
        added = 0
        for kw in failed:
            if kw not in existing:
                self._queue_listbox.insert("end", kw)
                existing.add(kw)
                added += 1
        messagebox.showinfo("补跑完成", f"从日志中提取 {len(failed)} 个失败词，成功添加 {added} 个到队列。")

    def _open_keyword_gen(self):
        """Open the keyword generation dialog and add keywords to queue."""
        from ui.keyword_gen import KeywordGenDialog
        def on_keywords(keywords):
            if keywords:
                self._keyword_var.set(keywords[0])
                for kw in keywords:
                    if kw not in self._queue_listbox.get(0, "end"):
                        self._queue_listbox.insert("end", kw)
        KeywordGenDialog(self.win, callback=on_keywords)

    def _load_from_db(self):
        """Load all filtered results from database, grouped by query."""
        try:
            import database as db
            saved = db.get_all_filtered_results()
            # Group by query
            groups = {}
            for r in saved:
                q = r["query"]
                if q not in groups:
                    groups[q] = []
                groups[q].append(r)
            for query_label, items in groups.items():
                batch = []
                for r in items:
                    self._idx_counter = max(self._idx_counter, r["id"])
                    item = {
                        "title": r["title"],
                        "snippet": r["snippet"] or "",
                        "link": r["link"] or "",
                        "relevance_reason": r.get("relevance_reason", ""),
                        "_query": query_label,
                        "_idx": r["id"],
                    }
                    self._results.append(item)
                    batch.append(item)
                # Build parent node and children
                parent_iid = f"G_{query_label}"
                self._query_groups[query_label] = batch
                self._tree.insert("", "end", iid=parent_iid, open=False,
                                  text=f"  {query_label} ({len(batch)}条)",
                                  values=(query_label, "", ""))
                for item in batch:
                    self._tree.insert(parent_iid, "end", iid=str(item["_idx"]),
                                      text=str(item["_idx"]),
                                      values=("", item["title"],
                                              item.get("relevance_reason", "")))
            self._update_filter_queries()
            total = len(self._results)
            if total > 0:
                self._status_var.set(f"就绪（已加载 {total} 条结果）")
                self._stats_label.config(text=f"数据库共 {total} 条结果")
        except Exception as e:
            print(f"[SerperDialog] load db error: {e}")

    def _refresh_from_db(self):
        """Clear tree and reload from filtered_results table."""
        self._tree.delete(*self._tree.get_children())
        self._results.clear()
        self._query_groups.clear()
        self._idx_counter = 0
        self._load_from_db()
        self._update_filter_queries()
        self._stats_label.config(text=f"累计 {len(self._results)} 条结果")
        self._status_var.set(f"已刷新，共 {len(self._results)} 条结果")

    def _update_filter_queries(self):
        queries = sorted(set(r["_query"] for r in self._results))
        self._filter_query_combo["values"] = ["全部"] + queries

    def _apply_filter(self):
        query_filter = self._filter_query_var.get()
        keyword = self._filter_search_var.get().strip().lower()
        self._tree.delete(*self._tree.get_children())

        visible = 0
        if query_filter == "全部":
            for q_label, items in self._query_groups.items():
                filtered = [it for it in items
                            if not keyword or keyword in it["title"].lower()]
                if not filtered:
                    continue
                visible += len(filtered)
                parent_iid = f"G_{q_label}"
                self._tree.insert("", "end", iid=parent_iid, open=False,
                                  text=f"  {q_label} ({len(filtered)}条)",
                                  values=("", q_label, "", ""))
                for item in filtered:
                    self._tree.insert(parent_iid, "end", iid=str(item["_idx"]),
                                      text=str(item["_idx"]),
                                      values=("", item["title"],
                                              item.get("relevance_reason", "")))
        else:
            for q_label, items in self._query_groups.items():
                if q_label == query_filter:
                    filtered = [it for it in items
                                if not keyword or keyword in it["title"].lower()]
                    visible = len(filtered)
                    parent_iid = f"G_{q_label}"
                    self._tree.insert("", "end", iid=parent_iid, open=True,
                                      text=f"  {q_label} ({len(filtered)}条)",
                                      values=("", q_label, "", ""))
                    for item in filtered:
                        self._tree.insert(parent_iid, "end", iid=str(item["_idx"]),
                                          text=str(item["_idx"]),
                                          values=("", item["title"],
                                                  item.get("relevance_reason", "")))
                    break

        self._status_var.set(f"筛选「{query_filter}」显示 {visible} 条")

        if self._sort_col:
            self._sort_by(self._sort_col, silent=True)

    def _clear_filter(self):
        self._filter_query_var.set("全部")
        self._filter_search_var.set("")
        self._apply_filter()

    def _sort_by(self, col, silent=False):
        if not silent and self._sort_col == col:
            self._sort_desc = not self._sort_desc
        elif not silent:
            self._sort_col = col
            self._sort_desc = False

        if not self._sort_col:
            return

        # Update column headers to show sort indicator
        col_labels = {"query": "搜索词", "title": "标题", "reason": "AI 筛选理由"}
        for c in ("query", "title", "reason"):
            base = col_labels[c]
            if c == self._sort_col:
                self._tree.heading(c, text=base + (" ▼" if self._sort_desc else " ▲"))
            else:
                self._tree.heading(c, text=base)

        # Sort leaf items within each parent group
        col_idx = {"query": 0, "title": 1, "reason": 2}[self._sort_col]
        for parent_iid in self._tree.get_children(""):
            if not parent_iid.startswith("G_"):
                continue
            children = self._tree.get_children(parent_iid)
            items = []
            for iid in children:
                val = self._tree.set(iid, col)
                items.append((val.lower() if val else "", iid))
            items.sort(key=lambda x: x[0], reverse=self._sort_desc)
            for i, (_, iid) in enumerate(items):
                self._tree.move(iid, parent_iid, i)

    # ---- Queue methods ----

    def _open_queue_status(self):
        """Open the queue status dialog."""
        if not hasattr(self, "_status_dlg") or self._status_dlg is None:
            self._status_dlg = QueueStatusDialog(self.win, self)
        else:
            self._status_dlg.win.lift()

    def _queue_add(self):
        text = self._keyword_var.get().strip()
        if not text:
            return
        for kw in text.replace(",", "\n").replace("，", "\n").split("\n"):
            kw = kw.strip()
            if kw and kw not in self._queue_listbox.get(0, "end"):
                self._queue_listbox.insert("end", kw)

    def _queue_del(self):
        for i in reversed(self._queue_listbox.curselection()):
            self._queue_listbox.delete(i)

    def _queue_clear(self):
        self._queue_listbox.delete(0, "end")

    def _queue_start(self):
        keywords = list(self._queue_listbox.get(0, "end"))
        if not keywords:
            messagebox.showinfo("提示", "请先添加关键词到队列")
            return

        # Read engine selection (thread-safe: BooleanVar.get() returns plain bool)
        use_serper = self._use_serper.get()
        use_baidu = self._use_baidu.get()

        self._queue_running = True
        self._queue_stop = False
        self._queue_paused = False
        self._queue_start_btn.config(state="disabled")
        self._queue_pause_btn.config(state="normal")

        # Save checkpoint for crash recovery
        self._save_checkpoint(keywords, 0, [])

        # Count total operations based on engine selection (or language-based default)
        def _ops_for_kw(kw):
            if use_serper or use_baidu:
                return (2 if use_serper and use_baidu else 1)
            # Neither selected → fall back to language default
            return 2 if self._is_chinese(kw) else 1

        total_ops = sum(_ops_for_kw(kw) for kw in keywords)
        self._queue_progress["maximum"] = total_ops
        self._queue_progress["value"] = 0
        self._queue_op_count = 0

        # Open/update status dialog
        self._open_queue_status()
        if self._status_dlg:
            self._status_dlg.set_ready(len(keywords), total_ops)

        import threading
        t = threading.Thread(target=self._queue_run, args=(keywords, total_ops, use_serper, use_baidu), daemon=True)
        t.start()

    @staticmethod
    def _is_chinese(text):
        return any('一' <= ch <= '鿿' for ch in text)

    def _queue_run(self, keywords, total_ops, use_serper=True, use_baidu=True):
        from utils.smart_search import smart_search, baidu_smart_search
        import database as db

        # Build engine list from user selection; if neither selected, fall back to language-based default
        def _get_engines(kw):
            engines = []
            if use_baidu:
                engines.append(("baidu", "[百度]"))
            if use_serper:
                engines.append(("serper", "[Serper]"))
            if not engines:
                # Fallback: Chinese → both, non-Chinese → Serper only
                if self._is_chinese(kw):
                    engines.extend([("baidu", "[百度]"), ("serper", "[Serper]")])
                else:
                    engines.append(("serper", ""))
            return engines

        serper_down = False
        network_down = False
        skipped_serper = []
        total_kw = len(keywords)
        resume_i = 0
        resume_engine_idx = 0

        i = 0
        while i < len(keywords):
            kw = keywords[i]

            if self._queue_stop:
                break

            while self._queue_paused and not self._queue_stop:
                self.win.after(0, lambda: self._status_var.set("已暂停..."))
                import time
                time.sleep(0.5)

            if self._queue_stop:
                break

            engines = _get_engines(kw)
            # Resume from last engine if recovering from network pause
            start_engine = resume_engine_idx if i == resume_i and resume_engine_idx > 0 else 0

            all_engines_failed = True
            for ei in range(start_engine, len(engines)):
                engine, label = engines[ei]

                if self._queue_stop:
                    break

                # Check network_down flag — pause until user clicks 继续
                if network_down:
                    self._queue_paused = True
                    self.win.after(0, lambda: self._status_var.set(
                        "网络中断！联网后点「继续」"))
                    self.win.after(0, lambda: self._queue_pause_btn.config(text="继续"))
                    while self._queue_paused and not self._queue_stop:
                        import time
                        time.sleep(0.5)
                    if self._queue_stop:
                        break
                    network_down = False
                    self.win.after(0, lambda: self._queue_pause_btn.config(text="暂停"))

                # Skip Serper if proxy is down
                if engine == "serper" and serper_down:
                    skipped_serper.append(kw)
                    continue

                self._queue_op_count += 1
                query_label = f"{label}{kw}" if label else kw
                progress_text = f"{i+1}/{total_kw} [{engine}]: {kw}"
                self.win.after(0, lambda t=progress_text: self._status_var.set(t))
                self.win.after(0, lambda v=self._queue_op_count: self._queue_progress.configure(value=v))
                self.win.after(0, lambda t=progress_text: self._queue_progress_var.set(t))
                self.win.after(0, lambda p=self._queue_op_count, k=kw, e=engine:
                    self._status_dlg and self._status_dlg.update_progress(p, k, e))

                try:
                    if engine == "baidu":
                        results, stats = baidu_smart_search(kw, api_key=self._api_key)
                    else:
                        results, stats = smart_search(kw, api_key=self._api_key)
                    all_engines_failed = False
                    if results:
                        db.insert_filtered_results(query_label, results)
                        self.win.after(0, lambda ql=query_label, r=results:
                            self._add_batch_to_tree(ql, r))
                    self.win.after(0, lambda k=kw, e=engine, c=len(results):
                        self._status_dlg and self._status_dlg.log_done(k, e, c))
                except Exception as e:
                    err_msg = str(e)
                    is_timeout = "timeout" in err_msg.lower() or "timed out" in err_msg.lower()
                    print(f"[queue] error on '{kw}' [{engine}]: {e}")
                    self.win.after(0, lambda err=err_msg[:50], k=kw, eng=engine:
                        self._status_dlg and self._status_dlg.log_error(k, eng, err))

                    # Auto-retry once after 30s for timeout
                    if is_timeout and not self._queue_stop:
                        self.win.after(0, lambda k=kw, e=engine:
                            self._status_var.set(f"[{e}] {k} 超时，30秒后重试..."))
                        self.win.after(0, lambda:
                            self._status_dlg and self._status_dlg.log_pause(f"[{engine}] 超时，30秒后自动重试..."))
                        import time
                        for _ in range(30):
                            if self._queue_stop:
                                break
                            time.sleep(1)
                        if self._queue_stop:
                            break
                        self.win.after(0, lambda k=kw, e=engine:
                            self._status_var.set(f"[{e}] {k} 重试中..."))
                        try:
                            if engine == "baidu":
                                results, stats = baidu_smart_search(kw, api_key=self._api_key)
                            else:
                                results, stats = smart_search(kw, api_key=self._api_key)
                            all_engines_failed = False
                            if results:
                                db.insert_filtered_results(query_label, results)
                                self.win.after(0, lambda ql=query_label, r=results:
                                    self._add_batch_to_tree(ql, r))
                            self.win.after(0, lambda k=kw, e=engine, c=len(results):
                                self._status_dlg and self._status_dlg.log_done(k, e, c))
                            continue  # retry succeeded, skip error handling below
                        except Exception as e2:
                            err_msg = str(e2)
                            self.win.after(0, lambda err=err_msg[:50], k=kw, eng=engine:
                                self._status_dlg and self._status_dlg.log_error(k, eng, f"重试失败: {err}"))

                    if engine == "serper" and ("Connection refused" in err_msg or "Errno 61" in err_msg or is_timeout):
                        serper_down = True
                        skipped_serper.append(kw)
                        self.win.after(0, lambda: self._status_var.set(
                            f"Serper调用失败(超时/断线)，跳过Serper，继续跑百度..."))
                        self.win.after(0, lambda:
                            self._status_dlg and self._status_dlg.log_pause("Serper调用失败(超时/断线)，跳过Serper"))
                    elif engine == "baidu" and ("Connection" in err_msg or "resolve" in err_msg.lower() or is_timeout):
                        # Baidu also down → full network outage, set flag and retry
                        network_down = True
                        self.win.after(0, lambda:
                            self._status_dlg and self._status_dlg.log_pause("网络中断！联网后点继续"))
                        break  # break engine loop, will pause at retry

            if network_down and not self._queue_stop:
                continue  # retry same keyword
            # Save checkpoint after each keyword
            self._save_checkpoint(keywords, i + 1, skipped_serper)
            i += 1

        # Report
        msgs = []
        if skipped_serper:
            msgs.append(f"{len(skipped_serper)}个词跳过Serper")
        if msgs:
            self.win.after(0, lambda m="，".join(msgs): self._status_var.set(
                f"完成。{m}。代理恢复后可重新加入队列再跑"))

        self.win.after(0, lambda: self._status_dlg and self._status_dlg.set_done(
            total_kw, skipped_serper))
        self.win.after(0, self._update_filter_queries)
        self.win.after(0, self._queue_done)

    def _queue_done(self):
        self._queue_running = False
        self._queue_start_btn.config(state="normal")
        self._queue_pause_btn.config(state="disabled")
        self._clear_checkpoint()
        total = len(self._results)
        self._status_var.set(f"排队完成，共 {total} 条结果")
        self._stats_label.config(text=f"累计 {total} 条结果")

    def _queue_pause(self):
        if self._queue_paused:
            self._queue_paused = False
            self._queue_pause_btn.config(text="暂停")
        else:
            self._queue_paused = True
            self._queue_pause_btn.config(text="继续")

    _CHECKPOINT_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                                     "data", "queue_checkpoint.json")

    def _save_checkpoint(self, keywords, current_idx, skipped_serper):
        try:
            import json
            os.makedirs(os.path.dirname(self._CHECKPOINT_FILE), exist_ok=True)
            with open(self._CHECKPOINT_FILE, "w") as f:
                json.dump({
                    "keywords": keywords,
                    "current_idx": current_idx,
                    "skipped_serper": skipped_serper,
                }, f, ensure_ascii=False)
        except Exception:
            pass

    def _load_checkpoint(self):
        try:
            import json
            if os.path.exists(self._CHECKPOINT_FILE):
                with open(self._CHECKPOINT_FILE) as f:
                    return json.load(f)
        except Exception:
            pass
        return None

    def _clear_checkpoint(self):
        try:
            if os.path.exists(self._CHECKPOINT_FILE):
                os.remove(self._CHECKPOINT_FILE)
        except Exception:
            pass

    def _queue_stop_search(self):
        self._queue_stop = True
        self._queue_paused = False
        self._queue_running = False
        self._queue_start_btn.config(state="normal")
        self._queue_pause_btn.config(state="disabled")
        self._status_var.set("已停止")
        self._clear_checkpoint()

    def _clear_all(self):
        self._results = []
        self._query_groups = {}
        self._removed_info = []
        self._idx_counter = 0
        self._tree.delete(*self._tree.get_children())
        self._stats_label.config(text="")
        self._status_var.set("就绪")
        self._filter_query_var.set("全部")
        self._filter_query_combo["values"] = ["全部"]
        self._detail_text.configure(state="normal")
        self._detail_text.delete("1.0", "end")
        self._detail_text.configure(state="disabled")
        try:
            import database as db
            db.clear_filtered_results()
            db.log_operation("调用api", "清空全部", "清空所有筛选结果")
        except Exception as e:
            print(f"[SerperDialog] clear db error: {e}")

    def _delete_selected(self):
        """Mark selected results as cleaned (hide from list, keep in raw data)."""
        sel = self._tree.selection()
        if not sel:
            return
        count = len(sel)
        ok = messagebox.askyesno("确认", f"确定删除选中的 {count} 项？\n数据将标记为已清洗，可通过原始结果中的还原按钮恢复。")
        if not ok:
            return

        child_ids = set()
        group_queries = set()
        for iid in sel:
            if iid.startswith("G_"):
                query = self._tree.item(iid, "values")[0]
                group_queries.add(query)
            else:
                try:
                    child_ids.add(int(iid))
                except ValueError:
                    pass

        # Collect links and mark raw_search_results as cleaned
        try:
            import database as db
            links_to_mark = []
            for q in group_queries:
                for item in self._query_groups.get(q, []):
                    link = item.get("link", "")
                    if link:
                        links_to_mark.append(link)
                    db.delete_filtered_result(item["_idx"])
            for rid in child_ids:
                for r in self._results:
                    if r["_idx"] == rid:
                        link = r.get("link", "")
                        if link:
                            links_to_mark.append(link)
                        break
                db.delete_filtered_result(rid)
            db.mark_raw_results_cleaned_by_links(links_to_mark)
        except Exception as e:
            print(f"[SerperDialog] delete db error: {e}")

        # Update memory and tree
        for q in group_queries:
            for item in self._query_groups.pop(q, []):
                self._results = [r for r in self._results if r["_idx"] != item["_idx"]]
        self._results = [r for r in self._results if r["_idx"] not in child_ids]

        # Update remaining groups
        for q, items in list(self._query_groups.items()):
            new_items = [it for it in items if it["_idx"] not in child_ids]
            if new_items:
                self._query_groups[q] = new_items
            else:
                del self._query_groups[q]

        self._tree.delete(*sel)
        # Update parent node counts, delete empty groups
        for q_label in list(self._query_groups.keys()):
            parent_iid = f"G_{q_label}"
            if self._tree.exists(parent_iid):
                self._tree.item(parent_iid, text=f"  {q_label} ({len(self._query_groups[q_label])}条)")
        for q_label in (group_queries - set(self._query_groups.keys())):
            parent_iid = f"G_{q_label}"
            if self._tree.exists(parent_iid):
                self._tree.delete(parent_iid)

        self._stats_label.config(text=f"累计 {len(self._results)} 条结果")
        self._status_var.set(f"已删除 {count} 项，剩余 {len(self._results)} 条")
        self._update_filter_queries()

        self._detail_text.configure(state="normal")
        self._detail_text.delete("1.0", "end")
        self._detail_text.configure(state="disabled")

        # Audit log
        import database as db
        db.log_operation("调用api", "数据清洗", f"删除{count}项", count)

    def _clean_selected(self):
        """Mark selected results as 'cleaned' in database and remove from tree."""
        sel = self._tree.selection()
        if not sel:
            return
        # Collect leaf items only
        leaf_ids = []
        for iid in sel:
            if not iid.startswith("G_"):
                try:
                    leaf_ids.append(int(iid))
                except ValueError:
                    pass
        if not leaf_ids:
            return
        count = len(leaf_ids)
        ok = messagebox.askyesno("数据清洗", f"确定清洗选中的 {count} 条数据？\n清洗后的数据将在原始结果中标记为「已清洗」。")
        if not ok:
            return

        # Save undo data
        removed_items = [dict(r) for r in self._results if r["_idx"] in leaf_ids]
        self._undo_stack.append({"type": "clean", "items": removed_items})
        if len(self._undo_stack) > 10:
            self._undo_stack.pop(0)
        n = len(self._undo_stack)
        self._undo_btn.config(state="normal", text=f"撤销({n})")
        self._undo_log_var.set(f"上次: 清洗 {len(removed_items)}条")

        # Mark as cleaned in DB
        import database as db
        links_to_clean = []
        for r in self._results:
            if r["_idx"] in leaf_ids:
                links_to_clean.append(r.get("link", ""))
        db.mark_raw_results_batch_cleaned(links_to_clean)

        # Remove from memory and tree
        self._results = [r for r in self._results if r["_idx"] not in leaf_ids]
        for iid in sel:
            try:
                self._tree.delete(iid)
            except tk.TclError:
                pass
        # Update group counts
        for q_label in list(self._query_groups.keys()):
            self._query_groups[q_label] = [r for r in self._query_groups[q_label] if r["_idx"] not in leaf_ids]
            parent_iid = f"G_{q_label}"
            if not self._query_groups[q_label]:
                try:
                    self._tree.delete(parent_iid)
                except tk.TclError:
                    pass
                del self._query_groups[q_label]
            elif self._tree.exists(parent_iid):
                self._tree.item(parent_iid, text=f"  {q_label} ({len(self._query_groups[q_label])}条)")

        self._stats_label.config(text=f"累计 {len(self._results)} 条结果")
        self._status_var.set(f"已清洗 {count} 项，剩余 {len(self._results)} 条")
        self._update_filter_queries()

        self._detail_text.configure(state="normal")
        self._detail_text.delete("1.0", "end")
        self._detail_text.configure(state="disabled")

    def _deduplicate(self):
        """Fast exact dedup: same link or same title (case-insensitive). O(n)."""
        self._removed_info = []
        seen_links = {}
        seen_titles = {}
        removed_links = set()

        remaining = []
        for r in self._results:
            link = r["link"]
            title_lower = r["title"].strip().lower()
            kept = None
            reason = ""

            if link and link in seen_links:
                kept = seen_links[link]
                reason = f"链接重复: {link[:80]}"
            elif title_lower in seen_titles:
                kept = seen_titles[title_lower]
                reason = f"标题重复: {r['title'][:60]}"

            if kept:
                self._removed_info.append({
                    "removed_title": r["title"],
                    "removed_query": r["_query"],
                    "removed_idx": r["_idx"],
                    "kept_title": kept[1],
                    "kept_idx": kept[0],
                    "reason": reason,
                })
                removed_links.add(link)
            else:
                seen_links[link] = (r["_idx"], r["title"])
                seen_titles[title_lower] = (r["_idx"], r["title"])
                remaining.append(r)

        self._finish_dedup(removed_links)

    def _fuzzy_dedup_selected(self):
        """AI fuzzy dedup on SELECTED tree items only."""
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选中需要模糊去重的条目。")
            return

        selected_results = []
        for iid in sel:
            for r in self._results:
                if str(r.get("_idx", "")) == iid:
                    selected_results.append(r)
                    break

        if len(selected_results) < 2:
            messagebox.showinfo("提示", "请至少选中2条才能去重。")
            return

        import threading
        self._status_var.set(f"AI模糊去重: {len(selected_results)}条选中, 估计 {len(selected_results)*(len(selected_results)-1)//2} 对比...")
        self.win.update()
        t = threading.Thread(target=self._ai_dedup_thread, args=(selected_results, set()), daemon=True)
        t.start()

    def _ai_dedup_thread(self, remaining, removed_links=None):
        """Run O(n²) fuzzy matching + AI in background thread."""
        if removed_links is None:
            removed_links = set()
        from difflib import SequenceMatcher
        from utils.smart_search import ai_confirm_duplicates

        candidates = []
        pairs = []
        total = len(remaining)
        ten_pct = max(1, (total * (total - 1)) // 200)

        n = 0
        for i in range(total):
            ri = remaining[i]
            ti = ri["title"].strip().lower()
            for j in range(i + 1, total):
                n += 1
                rj = remaining[j]
                tj = rj["title"].strip().lower()
                sm = SequenceMatcher(None, ti, tj)
                ratio = sm.ratio()
                if ratio > 0.7:
                    candidates.append((
                        ri["title"], ri.get("snippet", ""),
                        rj["title"], rj.get("snippet", ""),
                    ))
                    pairs.append((ri, rj))
                elif ratio > 0.5:
                    match = sm.find_longest_match(0, len(ti), 0, len(tj))
                    common = ti[match.a:match.a + match.size].strip()
                    if match.size >= 8 and len(common) >= 8:
                        candidates.append((
                            ri["title"], ri.get("snippet", ""),
                            rj["title"], rj.get("snippet", ""),
                        ))
                        pairs.append((ri, rj))
                if n % ten_pct == 0:
                    self.win.after(0, lambda _n=n, _t=total:
                        self._status_var.set(
                            f"AI模糊去重: 对比中 {_n}/{(total*(total-1))//2} ..."))

        if candidates:
            self.win.after(0, lambda: self._status_var.set(f"AI复查 {len(candidates)} 对候选..."))
            dup_set = ai_confirm_duplicates(candidates, api_key=self._api_key)
            dup_removed = set()
            for k, (ri, rj) in enumerate(pairs):
                if k in dup_set:
                    if ri["_idx"] > rj["_idx"]:
                        removed, kept = ri, rj
                    else:
                        removed, kept = rj, ri
                    if removed["_idx"] not in dup_removed:
                        dup_removed.add(removed["_idx"])
                        self._removed_info.append({
                            "removed_title": removed["title"],
                            "removed_query": removed["_query"],
                            "removed_idx": removed["_idx"],
                            "kept_title": kept["title"],
                            "kept_idx": kept["_idx"],
                            "reason": f"AI确认重复: {removed['title'][:50]}",
                        })
                        removed_links.add(removed["link"])

        self.win.after(0, lambda: self._finish_dedup(removed_links))

    def _finish_dedup(self, removed_links):

        if not self._removed_info:
            messagebox.showinfo("去重", "没有发现重复条目")
            return

        removed_ids = {info["removed_idx"] for info in self._removed_info}

        # Save undo data
        removed_items = [dict(r) for r in self._results if r["_idx"] in removed_ids]
        self._undo_stack.append({
            "type": "dedup",
            "items": removed_items,
            "removed_links": list(removed_links),
            "removed_ids": removed_ids,
        })
        if len(self._undo_stack) > 10:
            self._undo_stack.pop(0)
        n = len(self._undo_stack)
        self._undo_btn.config(state="normal", text=f"撤销({n})")
        self._undo_log_var.set(f"上次: 去重 {len(removed_items)}条")

        self._results = [r for r in self._results if r["_idx"] not in removed_ids]
        # Rebuild query groups and tree
        self._query_groups = {}
        for r in self._results:
            q = r["_query"]
            if q not in self._query_groups:
                self._query_groups[q] = []
            self._query_groups[q].append(r)
        self._tree.delete(*self._tree.get_children())
        for q_label, items in self._query_groups.items():
            parent_iid = f"G_{q_label}"
            self._tree.insert("", "end", iid=parent_iid, open=False,
                              text=f"  {q_label} ({len(items)}条)",
                              values=("", q_label, "", ""))
            for item in items:
                self._tree.insert(parent_iid, "end", iid=str(item["_idx"]),
                                  text=str(item["_idx"]),
                                  values=("", item["title"],
                                          item.get("relevance_reason", "")))

        # Persist dedup-removed links and delete from filtered_results
        # Also mark dedup_status in raw_search_results
        try:
            import database as db
            db.insert_dedup_removed(list(removed_links))
            for rid in removed_ids:
                db.delete_filtered_result(rid)
            # Map removed_idx -> link from self._results for dedup_status update
            result_links = {r["_idx"]: r.get("link", "") for r in self._results}
            for info in self._removed_info:
                link = result_links.get(info.get("removed_idx"), "")
                reason = info.get("reason", "")
                if "AI确认" in reason:
                    dedup_status = "fuzzy_dedup"
                elif "标题重复" in reason:
                    dedup_status = "title_dedup"
                elif "链接重复" in reason:
                    dedup_status = "url_dedup"
                else:
                    dedup_status = "url_dedup"
                db.mark_raw_results_dedup_status(link, dedup_status)
        except Exception as e:
            print(f"[SerperDialog] dedup save error: {e}")

        dup_count = len(self._removed_info)
        self._stats_label.config(text=f"累计 {len(self._results)} 条结果")
        self._status_var.set(f"去重完成，移除 {dup_count} 条重复，剩余 {len(self._results)} 条")

        # Audit log
        db.log_operation("调用api", "去重", f"移除{dup_count}条重复", dup_count)

    def _show_removed_info(self):
        """Show a dialog listing all removed duplicate items and their reasons."""
        if not self._removed_info:
            messagebox.showinfo("移除信息", "暂无移除记录")
            return

        dlg = tk.Toplevel(self.win)
        dlg.title("去重移除详情")
        dlg.geometry("750x400")
        dlg.transient(self.win)

        columns = ("removed_title", "kept_idx", "reason")
        tree = ttk.Treeview(dlg, columns=columns, show="headings")
        tree.heading("removed_title", text="被移除的条目")
        tree.column("removed_title", width=300)
        tree.heading("kept_idx", text="保留条目#")
        tree.column("kept_idx", width=80)
        tree.heading("reason", text="原因")
        tree.column("reason", width=350)

        scroll_y = ttk.Scrollbar(dlg, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll_y.set)
        tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scroll_y.pack(side="right", fill="y", pady=5)

        for i, info in enumerate(self._removed_info):
            tree.insert("", "end", values=(
                info["removed_title"],
                info["kept_idx"],
                info["reason"],
            ))

        ttk.Label(dlg, text=f"共移除 {len(self._removed_info)} 条重复",
                  padding=5).pack(side="bottom")
        ttk.Button(dlg, text="关闭", command=dlg.destroy).pack(side="bottom", pady=3)

    def _undo_last(self):
        """Undo the last dedup or clean operation (up to 10 steps)."""
        if not self._undo_stack:
            messagebox.showinfo("撤销", "没有可撤销的操作。")
            return

        data = self._undo_stack.pop()
        if not self._undo_stack:
            self._undo_btn.config(state="disabled", text="撤销")
        else:
            self._undo_btn.config(text=f"撤销({len(self._undo_stack)})")

        op_type = data["type"]
        items = data["items"]
        count = len(items)

        import database as db

        if op_type == "dedup":
            # Restore to DB: re-insert filtered_results, remove from dedup_removed
            for r in items:
                try:
                    db.insert_filtered_results(r.get("_query", ""), [r])
                except Exception as e:
                    print(f"[undo] re-insert error: {e}")
            # Remove from dedup_removed
            removed_links = data.get("removed_links", [])
            if removed_links:
                try:
                    db.delete_dedup_removed_by_links(removed_links)
                except Exception as e:
                    print(f"[undo] dedup_removed delete error: {e}")

        elif op_type == "clean":
            # Restore raw results status back to active/ds_selected
            links = [r.get("link", "") for r in items if r.get("link")]
            titles = [r.get("title", "") for r in items if not r.get("link")]
            if links or titles:
                try:
                    db.unmark_raw_results_cleaned(links, titles)
                except Exception as e:
                    print(f"[undo] unmark error: {e}")

        # Restore to memory and tree
        for r in items:
            if r not in self._results:
                self._results.append(r)
            q = r.get("_query", "")
            if q not in self._query_groups:
                self._query_groups[q] = []
            if r not in self._query_groups[q]:
                self._query_groups[q].append(r)

        # Rebuild tree
        self._tree.delete(*self._tree.get_children())
        for q_label, group_items in self._query_groups.items():
            parent_iid = f"G_{q_label}"
            self._tree.insert("", "end", iid=parent_iid, open=False,
                              text=f"  {q_label} ({len(group_items)}条)",
                              values=("", q_label, "", ""))
            for item in group_items:
                self._tree.insert(parent_iid, "end", iid=str(item["_idx"]),
                                  text=str(item["_idx"]),
                                  values=("", item["title"],
                                          item.get("relevance_reason", "")))

        self._stats_label.config(text=f"累计 {len(self._results)} 条结果")
        self._status_var.set(f"已撤销 {op_type}，恢复 {count} 条")
        self._update_filter_queries()

    def _show_raw_results(self):
        """Show all raw search results grouped by query (Serper + Baidu only)."""
        import database as db
        queries = db.get_raw_result_queries()
        # Exclude scholar results (those belong to LiteratureSearchDialog)
        queries = [q for q in queries if q["engine"] != "scholar"]
        if not queries:
            messagebox.showinfo("原始结果", "暂无原始结果记录")
            return

        total_results = sum(q["count"] for q in queries)

        dlg = tk.Toplevel(self.win)
        dlg.title("原始搜索结果")
        dlg.geometry("1200x600")
        dlg.transient(self.win)

        paned = ttk.PanedWindow(dlg, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        # Left: query list
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=0)
        ttk.Label(left_frame, text="搜索记录", font=("", 10, "bold")).pack(anchor="w")
        query_tree = ttk.Treeview(left_frame, columns=("info",), show="headings",
                                   selectmode="extended")
        query_tree.heading("info", text="查询 / API / 条数")
        query_tree.column("info", width=280)
        q_scroll_y = ttk.Scrollbar(left_frame, orient="vertical", command=query_tree.yview)
        query_tree.configure(yscrollcommand=q_scroll_y.set)
        query_tree.pack(side="left", fill="both", expand=True)
        q_scroll_y.pack(side="right", fill="y")
        for q in queries:
            query_tree.insert("", "end", values=(f"[{q['engine']}] {q['query']} ({q['count']}条)",))

        # Right: results table
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        ttk.Label(right_frame, text="原始条目", font=("", 10, "bold")).pack(anchor="w")

        # Search box
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill="x", pady=(2, 0))
        ttk.Label(search_frame, text="搜索：").pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side="left", padx=5)
        ttk.Label(search_frame, text="（筛选标题）", font=("", 7), foreground="gray").pack(side="left")

        columns = ("query", "engine", "title", "status", "created_at")
        detail_tree = ttk.Treeview(right_frame, columns=columns, show="headings")
        detail_tree.heading("query", text="检索词")
        detail_tree.column("query", width=80)
        detail_tree.heading("engine", text="API")
        detail_tree.column("engine", width=50)
        detail_tree.heading("title", text="标题")
        detail_tree.column("title", width=350)
        detail_tree.heading("status", text="筛选状态")
        detail_tree.column("status", width=70)
        detail_tree.heading("created_at", text="检索时间")
        detail_tree.column("created_at", width=140)

        scroll_y = ttk.Scrollbar(right_frame, orient="vertical", command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=scroll_y.set)
        detail_tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        # Stash hidden fields for detail view, and all rows for filtering
        _raw_detail = {}
        _all_rows = []

        # Detail area for selected row (below paned, full width)
        detail_frame = ttk.LabelFrame(dlg, text=" 选中详情 ", padding=3)
        detail_frame.pack(fill="x", padx=5, pady=(3, 0))
        detail_text = tk.Text(detail_frame, height=6, wrap="word", state="disabled")
        detail_text.pack(fill="both")

        def _on_select(event):
            sel = detail_tree.selection()
            if not sel:
                return
            iid = sel[0]
            vals = detail_tree.item(iid, "values")
            extra = _raw_detail.get(iid, {})
            detail_text.configure(state="normal")
            detail_text.delete("1.0", "end")
            detail_text.insert("1.0", f"""检索词：{vals[0]}
API：{vals[1]}
标题：{vals[2]}

摘要：
{extra.get('snippet', '')}

链接：{extra.get('link', '')}
筛选状态：{vals[3]}
AI筛选理由：{extra.get('reason', '')}
检索时间：{vals[4]}""")
            detail_text.configure(state="disabled")

        detail_tree.bind("<<TreeviewSelect>>", _on_select)

        def _apply_search():
            keyword = search_var.get().strip().lower()
            detail_tree.delete(*detail_tree.get_children())
            _raw_detail.clear()
            for r in _all_rows:
                if not keyword or keyword in r["title"].lower():
                    iid = detail_tree.insert("", "end", values=(
                        r["query"], r["engine"],
                        r["title"],
                        r["status"], r["created_at"] or ""))
                    _raw_detail[iid] = {
                        "id": r["id"],
                        "snippet": r["snippet"] or "",
                        "link": r["link"] or "",
                        "reason": r["reason"] or "",
                    }

        search_var.trace_add("write", lambda *_: _apply_search())

        def on_select_query(event):
            sel = query_tree.selection()
            if not sel:
                return
            search_var.set("")
            _all_rows.clear()
            for sel_iid in sel:
                idx = query_tree.index(sel_iid)
                q = queries[idx]
                for r in db.get_raw_results_by_query(q["query"], q["engine"]):
                    _all_rows.append({
                        "id": r["id"],
                        "query": q["query"],
                        "engine": q["engine"],
                        "title": r["title"],
                        "status": r["status"],
                        "created_at": r["created_at"] or "",
                        "snippet": r["snippet"] or "",
                        "link": r["link"] or "",
                        "reason": r["reason"] or "",
                    })
            _apply_search()

        query_tree.bind("<<TreeviewSelect>>", on_select_query)

        bottom = ttk.Frame(dlg, padding=5)
        bottom.pack(fill="x")

        def _export():
            all_data = []
            for q in queries:
                for r in db.get_raw_results_by_query(q["query"], q["engine"]):
                    all_data.append({
                        "搜索词": q["query"],
                        "API": q["engine"],
                        "标题": r["title"],
                        "摘要": r["snippet"],
                        "链接": r["link"],
                        "筛选状态": r["status"],
                        "AI筛选理由": r["reason"],
                        "检索时间": r["created_at"],
                    })
            if not all_data:
                return
            df = pd.DataFrame(all_data)
            from utils.import_export import export_to_excel
            export_to_excel(df, title=f"原始结果_全部{len(queries)}条搜索记录")

        ttk.Button(bottom, text="导出Excel", command=_export).pack(side="left", padx=2)

        def _delete_selected():
            # First check if a query group (left side) is selected
            q_sel = query_tree.selection()
            d_sel = detail_tree.selection()
            if q_sel:
                # Delete entire query group
                idx = query_tree.index(q_sel[0])
                q = queries[idx]
                count = q["count"]
                if not messagebox.askyesno("确认删除", f"确定删除搜索记录「{q['query']}」的全部 {count} 条原始结果？\n此操作不可撤销。"):
                    return
                ids_to_del = [_raw_detail.get(iid, {}).get("id") for iid in detail_tree.get_children("")]
                ids_to_del = [x for x in ids_to_del if x]
                if ids_to_del:
                    db.delete_scholar_results_by_ids(ids_to_del)
                query_tree.delete(q_sel[0])
                detail_tree.delete(*detail_tree.get_children(""))
            elif d_sel:
                # Delete selected rows
                ids_to_del = [_raw_detail.get(iid, {}).get("id") for iid in d_sel]
                ids_to_del = [x for x in ids_to_del if x]
                if not ids_to_del:
                    return
                if not messagebox.askyesno("确认删除", f"确定删除选中的 {len(ids_to_del)} 条原始结果？\n此操作不可撤销。"):
                    return
                db.delete_raw_results_by_ids(ids_to_del)
                for iid in d_sel:
                    detail_tree.delete(iid)
            else:
                messagebox.showinfo("提示", "请先选中左侧搜索记录或右侧原始条目")

        ttk.Button(bottom, text="删除选中", command=_delete_selected).pack(side="left", padx=2)

        def _manual_select():
            d_sel = detail_tree.selection()
            if not d_sel:
                messagebox.showinfo("提示", "请先选中右侧要添加的条目")
                return
            reason_dlg = tk.Toplevel(dlg)
            reason_dlg.title("手动筛选理由")
            reason_dlg.geometry("420x120")
            reason_dlg.transient(dlg)
            ttk.Label(reason_dlg, text="请输入筛选理由（一句话说明为什么与海外中式建筑相关）：").pack(padx=10, pady=(10, 5))
            reason_var = tk.StringVar()
            reason_entry = ttk.Entry(reason_dlg, textvariable=reason_var, width=50)
            reason_entry.pack(padx=10, pady=5)
            reason_entry.focus_set()
            result = {"reason": None}

            def _ok():
                r = reason_var.get().strip()
                if not r:
                    messagebox.showwarning("提示", "请输入筛选理由")
                    return
                result["reason"] = r
                reason_dlg.destroy()

            btn_frame = ttk.Frame(reason_dlg)
            btn_frame.pack(pady=5)
            ttk.Button(btn_frame, text="确定", command=_ok).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="取消", command=reason_dlg.destroy).pack(side="left", padx=5)
            reason_entry.bind("<Return>", lambda e: _ok())
            dlg.wait_window(reason_dlg)
            if result["reason"] is None:
                return

            reason = result["reason"]
            all_by_id = {r["id"]: r for r in _all_rows}
            by_query = {}
            for iid in d_sel:
                extra = _raw_detail.get(iid, {})
                rid = extra.get("id")
                if not rid:
                    continue
                entry = all_by_id.get(rid)
                if not entry:
                    continue
                q = entry["query"]
                if q not in by_query:
                    by_query[q] = []
                by_query[q].append({
                    "title": entry["title"],
                    "snippet": entry.get("snippet", ""),
                    "link": entry.get("link", ""),
                    "relevance_reason": reason,
                })
            if not by_query:
                return
            for q, items in by_query.items():
                db.insert_filtered_results(q, items)
            on_select_query(None)
            self._refresh_from_db()

        ttk.Button(bottom, text="选中", command=_manual_select).pack(side="left", padx=2)

        def _restore_selected():
            d_sel = detail_tree.selection()
            if not d_sel:
                messagebox.showinfo("提示", "请先选中右侧要还原的条目")
                return
            ids_to_restore = []
            links_to_restore = []
            for iid in d_sel:
                extra = _raw_detail.get(iid, {})
                rid = extra.get("id")
                if rid:
                    ids_to_restore.append(rid)
                    links_to_restore.append(extra.get("link", ""))
            if not ids_to_restore:
                return
            all_by_id = {r["id"]: r for r in _all_rows}
            if not messagebox.askyesno("确认还原", f"确定还原选中的 {len(ids_to_restore)} 条记录？"):
                return
            db.restore_raw_results_full(ids_to_restore, links_to_restore, {})
            reinsert_by_query = {}
            for rid in ids_to_restore:
                entry = all_by_id.get(rid)
                if entry and entry.get("reason"):
                    q = entry["query"]
                    if q not in reinsert_by_query:
                        reinsert_by_query[q] = []
                    reinsert_by_query[q].append({
                        "title": entry["title"],
                        "snippet": entry.get("snippet", ""),
                        "link": entry.get("link", ""),
                        "relevance_reason": entry.get("reason", ""),
                    })
            for q, items in reinsert_by_query.items():
                db.insert_filtered_results(q, items)
            on_select_query(None)

        # Build id->entry map for use by buttons
        all_by_id = {r["id"]: r for r in _all_rows}
        ttk.Button(bottom, text="还原", command=_restore_selected).pack(side="left", padx=2)
        ttk.Label(bottom, text=f"共 {total_results} 条原始结果（{len(queries)} 次搜索）").pack(side="left", padx=10)
        ttk.Button(bottom, text="关闭", command=dlg.destroy).pack(side="right")

        # Auto-select first query
        if queries:
            query_tree.selection_set(query_tree.get_children()[0])
            on_select_query(None)

    def _on_select(self, event):
        sel = self._tree.selection()
        if not sel:
            self._expand_btn.configure(text="展开/合上")
            return
        iid = sel[0]
        self._update_expand_btn_text()
        # Parent node (query group) — expand/collapse
        if iid.startswith("G_"):
            return
        # Child node — show detail
        try:
            idx = int(iid)
        except ValueError:
            return
        for r in self._results:
            if r["_idx"] == idx:
                self._detail_text.configure(state="normal")
                self._detail_text.delete("1.0", "end")
                reason = r.get("relevance_reason", "")
                self._detail_text.insert("end",
                    f"搜索词: {r['_query']}\n"
                    f"标题: {r['title']}\n"
                    f"AI筛选理由: {reason}\n\n"
                    f"摘要: {r['snippet']}\n\n"
                    f"链接: {r['link']}")
                self._detail_text.configure(state="disabled")
                break

    def _update_expand_btn_text(self):
        g_items = [iid for iid in self._tree.selection() if iid.startswith("G_")]
        if not g_items:
            self._expand_btn.configure(text="展开/合上")
            return
        states = [self._tree.item(iid, "open") for iid in g_items]
        if all(states):
            self._expand_btn.configure(text="合上")
        elif not any(states):
            self._expand_btn.configure(text="展开")
        else:
            self._expand_btn.configure(text="展开/合上")

    def _expand_selected(self):
        """Expand/collapse selected query-group parent nodes."""
        for iid in self._tree.selection():
            if iid.startswith("G_"):
                self._tree.item(iid, open=not self._tree.item(iid, "open"))
        self._update_expand_btn_text()

    def _get_selected_results(self):
        ids = set()
        for iid in self._tree.selection():
            if not iid.startswith("G_"):
                try:
                    ids.add(int(iid))
                except ValueError:
                    pass
        return [r for r in self._results if r["_idx"] in ids]

    def _send_to_library(self):
        """Send selected search results to the literature database.
        Marks raw_search_results as '转为文献', and tags literature as '由灰色文献转入'."""
        sel = self._get_selected_results()
        if not sel:
            messagebox.showinfo("提示", "请先在结果列表中选中要入库的条目。")
            return
        import database as db
        saved = 0
        skipped = 0
        saved_links = []
        saved_ids = []
        for r in sel:
            title = (r.get("title") or "").strip()
            link = r.get("link", "")
            if not title:
                continue
            existing = db.search_literature_by_title(title)
            if existing:
                skipped += 1
                continue
            # Try to get richer metadata from raw_search_results
            authors = ""
            pub_date = ""
            engine = "Web搜索"
            try:
                raw = db.get_raw_result_by_link(link)
                if raw:
                    authors = raw.get("authors", "") or ""
                    pub_date = raw.get("year", "") or ""
                    engine = raw.get("engine", "") or "Web搜索"
            except Exception:
                pass
            # Build notes: AI筛选理由 + 灰色文献转入标记
            notes_parts = []
            reason = r.get("relevance_reason", "")
            if reason:
                notes_parts.append(f"AI筛选: {reason}")
            notes_parts.append("由灰色文献转入")
            db.insert_literature({
                "title": title,
                "authors": authors,
                "pub_date": str(pub_date),
                "source_url": link,
                "source_engine": engine,
                "summary": r.get("snippet", ""),
                "notes": " | ".join(notes_parts),
            })
            # Mark raw_search_results as 转为文献
            if link:
                try:
                    db.mark_raw_result_status(link, "转为文献")
                except Exception:
                    pass
            saved_links.append(link)
            saved_ids.append(r["_idx"])
            saved += 1
        # Remove saved items from tree and in-memory list immediately
        saved_ids_set = set(saved_ids)
        for iid in saved_ids_set:
            iid_str = str(iid)
            if self._tree.exists(iid_str):
                self._tree.delete(iid_str)
        self._results = [r for r in self._results if r["_idx"] not in saved_ids_set]
        # Clean up query groups and parent nodes
        for q_label in list(self._query_groups.keys()):
            group = self._query_groups[q_label]
            group[:] = [r for r in group if r["_idx"] not in saved_ids_set]
            if not group:
                parent_iid = f"G_{q_label}"
                if self._tree.exists(parent_iid):
                    self._tree.delete(parent_iid)
                del self._query_groups[q_label]
            else:
                parent_iid = f"G_{q_label}"
                if self._tree.exists(parent_iid):
                    self._tree.item(parent_iid, text=f"  {q_label} ({len(group)}条)")
        # Delete from filtered_results table so they don't come back on next load
        if saved_ids:
            try:
                db.delete_filtered_results_by_ids(saved_ids)
            except Exception:
                pass
        if saved_links:
            try:
                db.delete_filtered_results_by_links(saved_links)
            except Exception:
                pass
        self._stats_label.config(text=f"累计 {len(self._results)} 条结果")
        self._update_filter_queries()
        if saved:
            self._status_var.set(f"已入库 {saved} 条文献")
            msg = f"已将 {saved} 条文献保存到文献库。"
            if skipped:
                msg += f"\n跳过 {skipped} 条（标题重复）。"
            db.log_operation("调用api", "发送到文献库", f"入库{saved}条", saved)
            messagebox.showinfo("入库完成", msg)
        else:
            messagebox.showinfo("提示", f"没有新文献可入库。\n跳过 {skipped} 条（标题重复）。")

    @staticmethod
    def _open_in_chrome(url):
        import subprocess
        subprocess.run(["open", "-a", "/Applications/Google Chrome.app", url])

    def _open_selected(self):
        for r in self._get_selected_results():
            link = r.get("pdfUrl") or r.get("link")
            if link:
                self._open_in_chrome(link)

    def _open_all(self):
        if not self._results:
            return
        query_filter = self._filter_query_var.get()
        links = []
        for r in self._results:
            if query_filter != "全部" and r["_query"] != query_filter:
                continue
            links.append(r.get("pdfUrl") or r.get("link"))

        if not links:
            messagebox.showinfo("提示", "当前没有可打开的结果")
            return

        if len(links) > 15:
            ok = messagebox.askyesno("确认", f"将打开 {len(links)} 个链接，可能影响浏览器性能，继续？")
            if not ok:
                return

        for link in links:
            self._open_in_chrome(link)


class QueueStatusDialog:
    """Non-modal dialog showing queue progress: completed/skipped/errors."""

    def __init__(self, parent, serper_dialog, group=None, log_file=None):
        self._sd = serper_dialog
        self._group = group  # the group dict this dialog is tracking
        self._log = []  # [(time_str, status, message), ...]
        self._skipped_serper = []
        self._log_saved = False
        self._done = False

        # Log file path (use provided path, or create a Serper one)
        if log_file:
            self._log_file = log_file
        else:
            import datetime
            log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
            os.makedirs(log_dir, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            self._log_file = os.path.join(log_dir, f"serper_queue_{ts}.log")

        gname = group["name"] if group else ""
        self.win = tk.Toplevel(parent)
        self.win.title(f"队列状态 - {gname}")
        self.win.geometry("550x450")
        self.win.transient(parent)
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

        # Summary
        summary_frame = ttk.Frame(self.win, padding=5)
        summary_frame.pack(fill="x")
        self._summary_var = tk.StringVar(value="就绪，等待开始...")
        ttk.Label(summary_frame, textvariable=self._summary_var,
                  font=("", 11, "bold")).pack(anchor="w")

        # Progress bar
        self._progress = ttk.Progressbar(self.win, mode="determinate")
        self._progress.pack(fill="x", padx=5, pady=(0, 5))
        self._current_var = tk.StringVar(value="")
        ttk.Label(self.win, textvariable=self._current_var, font=("", 9)).pack(anchor="w", padx=5)

        # Log area
        log_frame = ttk.LabelFrame(self.win, text=" 运行日志 ", padding=5)
        log_frame.pack(fill="both", expand=True, padx=5, pady=(5, 0))
        self._log_text = tk.Text(log_frame, wrap="word", font=("", 9), height=15,
                                 state="disabled")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical",
                                   command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=log_scroll.set)
        self._log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        # Buttons
        btn_frame = ttk.Frame(self.win, padding=5)
        btn_frame.pack(fill="x")
        self._pause_btn = ttk.Button(btn_frame, text="暂停", state="disabled",
                                     command=self._toggle_pause)
        self._pause_btn.pack(side="left", padx=2)
        self._retry_btn = ttk.Button(btn_frame, text="补跑", state="disabled",
                                     command=self._retry_skipped)
        self._retry_btn.pack(side="left", padx=2)
        self._stop_btn = ttk.Button(btn_frame, text="停止", command=self._stop)
        self._stop_btn.pack(side="left", padx=2)
        ttk.Button(btn_frame, text="历史日志", command=self._sd._open_log_viewer).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="隐藏", command=self.win.withdraw).pack(side="right", padx=2)

    def _toggle_pause(self):
        g = self._group
        if not g:
            return
        if g.get("paused"):
            g["paused"] = False
            self._pause_btn.config(text="暂停")
            self._sd._refresh_group_controls()
            self._log_msg("继续", "队列已恢复")
        else:
            g["paused"] = True
            self._pause_btn.config(text="继续")
            self._sd._refresh_group_controls()
            self._log_msg("暂停", "队列已暂停")

    def _stop(self):
        g = self._group
        if g:
            g["stop"] = True
            g["paused"] = False
        self._log_msg("停止", "已请求停止...")

    def _retry_skipped(self):
        if not self._skipped_serper:
            return
        # Deduplicate and add to queue
        existing = set(self._sd._queue_listbox.get(0, "end"))
        added = 0
        for kw in self._skipped_serper:
            if kw not in existing:
                self._sd._queue_listbox.insert("end", kw)
                existing.add(kw)
                added += 1
        self._log_msg("补跑", f"已添加 {added} 个跳过Serper的关键词到队列")
        self._skipped_serper = []
        self._retry_btn.config(state="disabled")

    def _log_msg(self, status, msg):
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log.append((ts, status, msg))
        self._log_text.configure(state="normal")
        self._log_text.insert("end", f"[{ts}] [{status}] {msg}\n")
        self._log_text.see("end")
        self._log_text.configure(state="disabled")
        # Append to file immediately
        self._append_to_file(f"[{ts}] [{status}] {msg}")

    def _append_to_file(self, line):
        try:
            with open(self._log_file, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass

    def _save_full_log(self):
        if self._log_saved:
            return
        import datetime
        try:
            with open(self._log_file, "w", encoding="utf-8") as f:
                f.write(f"队列搜索日志\n")
                f.write(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"{'='*50}\n\n")
                for ts, status, msg in self._log:
                    f.write(f"[{ts}] [{status}] {msg}\n")
            self._log_saved = True
        except Exception:
            pass

    def _close(self):
        self._save_full_log()
        self.win.destroy()

    def _on_close(self):
        if self._done:
            self._close()
        else:
            self._save_full_log()
            self.win.withdraw()

    def set_ready(self, total, total_ops):
        self._progress["maximum"] = total
        self._progress["value"] = 0
        self._skipped_serper = []
        self._summary_var.set(f"共 {total} 个关键词，{total_ops} 次搜索操作")
        self._current_var.set("")
        self._pause_btn.config(state="normal")
        self._retry_btn.config(state="disabled")
        self._log_msg("开始", f"队列启动: {total}词 / {total_ops}操作")

    def update_progress(self, pos, current_kw, engine):
        self._progress["value"] = pos
        self._current_var.set(f"正在处理 [{engine}]: {current_kw}")

    def log_done(self, kw, engine, results_count):
        self._log_msg("完成", f"[{engine}] {kw} → {results_count}条")

    def log_skip(self, kw, reason):
        if kw not in self._skipped_serper:
            self._skipped_serper.append(kw)
        self._log_msg("跳过", f"[Serper] {kw} ({reason})")

    def log_error(self, kw, engine, error):
        self._log_msg("错误", f"[{engine}] {kw}: {error[:60]}")

    def log_pause(self, reason):
        self._log_msg("暂停", reason)
        self._current_var.set(reason)

    def set_done(self, total, skipped):
        self._progress["value"] = self._progress["maximum"]
        self._pause_btn.config(state="disabled")
        self._skipped_serper = skipped
        self._done = True
        parts = [f"完成 {total} 个关键词"]
        if skipped:
            parts.append(f"{len(skipped)}个失败可补跑")
            self._retry_btn.config(state="normal")
        self._summary_var.set("，".join(parts))
        self._current_var.set("队列完成")
        self._log_msg("完成", "队列全部完成")
        self._save_full_log()
        self._log_text.configure(state="normal")
        self._log_text.insert("end", f"\n--- 日志已保存至: {self._log_file} ---\n")
        self._log_text.see("end")
        self._log_text.configure(state="disabled")
        self._stop_btn.config(text="关闭", command=self._close)


class LiteratureSearchDialog:
    """Dialog for Google Scholar search via Serper.dev, designed for literature (PRISMA identification phase).
    Supports single search and batch queue search — results can be saved to the literature table."""

    def __init__(self, parent, keyword="", search_tab=None):
        self._search_tab = search_tab
        self._results = []          # flat list of all result dicts
        self._query_groups = {}     # query_label -> [results]
        self._idx_counter = 0
        self._groups = []          # [{"name": str, "keywords": list, "running": bool, "paused": bool, "stop": bool, "progress": int, "max": int}]
        self._active_group = None  # index of active group
        self._undo_stack = []
        self._status_dlg = None
        self._log_lock = threading.Lock()  # protect concurrent log file writes from multiple groups
        self._log_saved = False

        # Log file
        import datetime
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
        os.makedirs(log_dir, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self._log_file = os.path.join(log_dir, f"scholar_queue_{ts}.log")

        self.win = tk.Toplevel(parent)
        self.win.title("文献筛选 - Google Scholar")
        self.win.geometry("1100x680")
        self.win.transient(parent)
        self._build_ui()

        self._load_from_db()

        # Load saved queue groups (replaces default group if saved data exists)
        saved_groups = self._load_queue_keywords()
        # Remove default groups created by _build_ui
        self._groups.clear()
        for tab_id in self._group_notebook.tabs():
            self._group_notebook.forget(tab_id)
        # Create groups from saved data
        for gdata in saved_groups:
            g = {"name": gdata["name"], "keywords": gdata.get("keywords", []),
                 "running": False, "paused": False, "stop": False,
                 "progress": 0, "max": 0, "failed_keywords": []}
            self._groups.append(g)
            tab_frame = ttk.Frame(self._group_notebook, width=0, height=0)
            self._group_notebook.add(tab_frame, text=g["name"])
        if not self._groups:
            self._add_group(name="队列1")
        self._activate_group(0)

        # Check for saved queue checkpoints (crash recovery)
        all_cp = self._load_all_checkpoints()
        recovered_any = False
        if all_cp:
            cp_names = list(all_cp.keys())
            cp_info = "\n".join(
                f"  {name}: {cp['done']}/{cp['total']} 已完成, {len(cp['keywords'])} 剩余"
                for name, cp in all_cp.items()
            )
            if messagebox.askyesno("恢复队列",
                    f"检测到 {len(all_cp)} 个未完成的队列任务：\n{cp_info}\n\n是否恢复并继续？"):
                for gname, cp in all_cp.items():
                    remaining = cp.get("keywords", [])
                    if not remaining:
                        continue
                    # Find matching group or create new one
                    found = None
                    for g in self._groups:
                        if g["name"] == gname:
                            found = g
                            break
                    if not found:
                        self._add_group(name=gname)
                        found = self._groups[-1]
                    found["keywords"] = remaining
                    found["max"] = cp.get("total", len(remaining))
                    found["progress"] = cp.get("done", 0)
                    recovered_any = True
                # Refresh active group display
                self._activate_group(self._active_group or 0)
                self._save_queue_keywords()
                if recovered_any:
                    self._status_var.set(f"已恢复 {len(all_cp)} 个队列")
                    self.win.after(200, self._queue_start)

        # Save queue keywords on window close
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

        if keyword:
            self._keyword_var.set(keyword)
            self.win.after(100, self._do_search)

    def _build_ui(self):
        # Toolbar
        toolbar = ttk.Frame(self.win, padding=5)
        toolbar.pack(fill="x")

        ttk.Label(toolbar, text="关键词：").pack(side="left")
        self._keyword_var = tk.StringVar()
        kw_entry = ttk.Entry(toolbar, textvariable=self._keyword_var, width=28)
        kw_entry.pack(side="left", padx=5)
        kw_entry.bind("<Return>", lambda e: self._do_search())
        ttk.Button(toolbar, text="生成关键词",
                   command=self._open_keyword_gen).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Scholar搜索", command=self._do_search).pack(side="left", padx=2)
        ttk.Button(toolbar, text="队列日志", command=self._open_log_viewer).pack(side="left", padx=2)

        self._use_gemini = False
        self._screen_var = tk.StringVar(value="DS")
        screen_frame = ttk.Frame(toolbar)
        screen_frame.pack(side="right", padx=(10, 0))
        ttk.Label(screen_frame, text="AI筛选：").pack(side="left")
        ttk.Radiobutton(screen_frame, text="DS", variable=self._screen_var,
                        value="DS", command=self._on_screen_change).pack(side="left")
        ttk.Radiobutton(screen_frame, text="Gemini", variable=self._screen_var,
                        value="Gemini", command=self._on_screen_change).pack(side="left")

        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(toolbar, textvariable=self._status_var).pack(side="right")

        # Stats + action bar
        self._stats_frame = ttk.Frame(self.win, padding=(5, 0))
        self._stats_frame.pack(fill="x")
        self._stats_label = ttk.Label(self._stats_frame, text="")
        self._stats_label.pack(side="left")
        ttk.Button(self._stats_frame, text="原始结果",
                   command=self._show_raw_results).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="刷新",
                   command=self._refresh_from_db).pack(side="right", padx=2)
        self._expand_btn = ttk.Button(self._stats_frame, text="展开/合上",
                   command=self._expand_selected)
        self._expand_btn.pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="在浏览器打开选中",
                   command=self._open_selected).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="去重",
                   command=self._deduplicate).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="匹配文献库",
                   command=self._match_literature).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="数据清洗",
                   command=self._delete_selected).pack(side="right", padx=2)
        self._tag_var = tk.StringVar()
        ttk.Entry(self._stats_frame, textvariable=self._tag_var, width=10).pack(side="right", padx=2)
        ttk.Button(self._stats_frame, text="标记", command=self._tag_selected).pack(side="right", padx=2)
        self._undo_btn = ttk.Button(self._stats_frame, text="撤销",
                                     command=self._undo_last, state="disabled")
        self._undo_btn.pack(side="right", padx=2)
        self._undo_log_var = tk.StringVar(value="")
        self._undo_log_label = ttk.Label(self._stats_frame, textvariable=self._undo_log_var,
                                          font=("", 7), foreground="gray")
        self._undo_log_label.pack(side="right", padx=(0, 8))

        # Filter bar
        filter_frame = ttk.Frame(self.win, padding=(5, 2))
        filter_frame.pack(fill="x")
        ttk.Label(filter_frame, text="按检索词筛选：").pack(side="left")
        self._filter_query_var = tk.StringVar(value="全部")
        self._filter_query_combo = ttk.Combobox(filter_frame, textvariable=self._filter_query_var,
                                                 state="readonly", width=18)
        self._filter_query_combo.pack(side="left", padx=2)
        self._filter_query_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        ttk.Button(filter_frame, text="清除筛选", command=self._clear_filter, width=6).pack(side="left", padx=5)

        # Vertical paned: main area + detail area (resizable)
        vpaned = ttk.PanedWindow(self.win, orient="vertical")
        vpaned.pack(fill="both", expand=True, padx=5, pady=(5, 0))

        # Main area: queue panel (left) + treeview (right)
        main_paned = ttk.PanedWindow(vpaned, orient="horizontal")
        vpaned.add(main_paned, weight=2)

        # -- Queue panel (multi-group) --
        queue_frame = ttk.LabelFrame(main_paned, text=" 关键词队列 ", padding=5)
        main_paned.add(queue_frame, weight=0)

        # Group tab bar
        group_bar = ttk.Frame(queue_frame)
        group_bar.pack(fill="x")
        self._group_notebook = ttk.Notebook(group_bar)
        self._group_notebook.pack(side="left", fill="x", expand=True)
        self._group_notebook.bind("<<NotebookTabChanged>>", self._on_group_switch)
        ttk.Button(group_bar, text="＋", width=3, command=self._add_group).pack(side="right")

        # Group content area - shared widgets that reflect active group
        group_content = ttk.Frame(queue_frame)
        group_content.pack(fill="both", expand=True)

        listbox_frame = ttk.Frame(group_content)
        listbox_frame.pack(fill="both", expand=True)
        self._queue_listbox = tk.Listbox(listbox_frame, width=20, height=12,
                                          selectmode="extended", exportselection=False)
        self._queue_listbox.pack(side="left", fill="both", expand=True)
        q_scroll = ttk.Scrollbar(listbox_frame, orient="vertical", command=self._queue_listbox.yview)
        self._queue_listbox.configure(yscrollcommand=q_scroll.set)
        q_scroll.pack(side="right", fill="y")

        qbtn_frame = ttk.Frame(group_content)
        qbtn_frame.pack(fill="x", pady=(3, 0))
        ttk.Button(qbtn_frame, text="添加", command=self._queue_add, width=5).pack(side="left", padx=1)
        ttk.Button(qbtn_frame, text="删除", command=self._queue_del, width=5).pack(side="left", padx=1)
        ttk.Button(qbtn_frame, text="清空", command=self._queue_clear, width=5).pack(side="left", padx=1)
        ttk.Button(qbtn_frame, text="重命名", command=self._rename_group, width=6).pack(side="left", padx=1)
        ttk.Button(qbtn_frame, text="删组", command=self._remove_group, width=5).pack(side="right", padx=1)

        qctrl_frame = ttk.Frame(group_content)
        qctrl_frame.pack(fill="x", pady=(5, 0))
        self._queue_start_btn = ttk.Button(qctrl_frame, text="排队搜索",
                                            command=self._queue_start)
        self._queue_start_btn.pack(fill="x")
        self._queue_pause_btn = ttk.Button(qctrl_frame, text="暂停",
                                            command=self._queue_pause, state="disabled")
        self._queue_pause_btn.pack(fill="x", pady=(2, 0))
        ttk.Button(qctrl_frame, text="停止", command=self._queue_stop_search).pack(fill="x", pady=(2, 0))

        self._queue_progress = ttk.Progressbar(group_content, mode="determinate")
        self._queue_progress.pack(fill="x", pady=(5, 0))
        self._queue_progress_var = tk.StringVar(value="")
        ttk.Label(group_content, textvariable=self._queue_progress_var,
                  font=("", 8)).pack()

        # Initialize default group
        self._add_group(name="队列1")

        # -- Treeview area --
        tree_panel = ttk.Frame(main_paned)
        main_paned.add(tree_panel, weight=1)

        tree_frame = ttk.Frame(tree_panel)
        tree_frame.pack(fill="both", expand=True)

        columns = ("query", "title", "reason", "tags")
        self._tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings",
                                   selectmode="extended")
        self._tree.heading("#0", text="搜索词组 / #")
        self._tree.column("#0", width=180, minwidth=140)
        self._tree.heading("query", text="搜索词",
                           command=lambda: self._sort_by("query"))
        self._tree.column("query", width=0, minwidth=0, stretch=False)
        self._tree.heading("title", text="标题",
                           command=lambda: self._sort_by("title"))
        self._tree.column("title", width=420, minwidth=260)
        self._tree.heading("reason", text="筛选理由",
                           command=lambda: self._sort_by("reason"))
        self._tree.column("reason", width=280, minwidth=140)
        self._tree.heading("tags", text="标记",
                           command=lambda: self._sort_by("tags"))
        self._tree.column("tags", width=100, minwidth=60)

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        _make_tree_draggable(self._tree)

        # Detail area (in vertical paned, resizable)
        detail_frame = ttk.LabelFrame(vpaned, text=" 选中详情 ", padding=5)
        vpaned.add(detail_frame, weight=1)

        self._detail_text = tk.Text(detail_frame, height=10, wrap="word", state="disabled")
        detail_scroll = ttk.Scrollbar(detail_frame, orient="vertical",
                                       command=self._detail_text.yview)
        self._detail_text.configure(yscrollcommand=detail_scroll.set)
        self._detail_text.pack(side="left", fill="both", expand=True)
        detail_scroll.pack(side="right", fill="y")

        # Bottom bar
        bottom = ttk.Frame(self.win, padding=5)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="入库选中", command=self._save_to_library).pack(side="left", padx=2)
        ttk.Button(bottom, text="导出Excel", command=self._export_excel).pack(side="left", padx=2)
        self._bottom_count_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self._bottom_count_var).pack(side="right")
        ttk.Button(bottom, text="关闭", command=self.win.destroy).pack(side="right", padx=2)

    def _load_from_db(self):
        """Load scholar raw results from database into tree."""
        try:
            import database as db
            queries = db.get_scholar_result_queries()
            if not queries:
                return
            self._batch_parents = {}
            for q in queries:
                engine = q.get("screen_engine", "") or "Scholar"
                query_label = f"[{engine}] {q['query']}"
                rows = db.get_scholar_results_by_query(q["query"], engine if engine != "Scholar" else None)
                # Load both DS-selected and gemini_selected results
                target_statuses = ("ds_selected", "gemini_selected")
                active_rows = [r for r in rows if r.get("clean_status", "active") in target_statuses]
                if not active_rows:
                    continue
                # Ensure batch parent exists
                if engine not in self._batch_parents:
                    batch_iid = self._tree.insert("", "end", iid=f"B_{engine}",
                                                  open=True, text=f"  {engine}")
                    self._batch_parents[engine] = batch_iid
                if query_label not in self._query_groups:
                    self._query_groups[query_label] = []
                parent_iid = f"G_{query_label}"
                self._tree.insert(self._batch_parents[engine], "end", iid=parent_iid, open=False,
                                  text=f"  {query_label} ({len(active_rows)}条)",
                                  values=(query_label, "", ""))
                for r in active_rows:
                    self._idx_counter += 1
                    item = {
                        "title": r["title"],
                        "authors": r.get("authors", ""),
                        "year": r.get("year", ""),
                        "publication": r.get("publication", ""),
                        "snippet": r.get("snippet", ""),
                        "link": r.get("link", ""),
                        "pdfUrl": r.get("pdf_url", ""),
                        "citedBy": r.get("cited_by", ""),
                        "gemini_reason": r.get("gemini_reason", ""),
                        "tags": r.get("tags", ""),
                        "_screen_engine": r.get("screen_engine", ""),
                        "_query": query_label,
                        "_idx": self._idx_counter,
                        "created_at": r.get("created_at", ""),
                    }
                    self._results.append(item)
                    self._query_groups[query_label].append(item)
                    self._tree.insert(parent_iid, "end", iid=f"L_{self._idx_counter}",
                                      text=str(self._idx_counter),
                                      values=("",
                                              r["title"],
                                              r.get("gemini_reason", ""),
                                              r.get("tags", "")))
            total = len(self._results)
            self._stats_label.config(text=f"已加载 {total} 条Scholar结果")
            self._status_var.set(f"就绪（已加载 {total} 条）")
            self._update_filter_queries()
            self._update_bottom_count()
        except Exception as e:
            print(f"[LiteratureSearch] load db error: {e}")

    def _refresh_from_db(self):
        """Clear tree and reload from DB."""
        self._tree.delete(*self._tree.get_children())
        self._results.clear()
        self._query_groups.clear()
        self._idx_counter = 0
        self._load_from_db()
        self._update_bottom_count()
        self._status_var.set(f"已刷新，共 {len(self._results)} 条结果")

    # ---- Search ----

    def _screen_batch_with_ds(self, batch, api_key):
        """Send a batch of Scholar results to DeepSeek, return only those likely
        to contain overseas Chinese architecture cases (outside China).

        Returns [(result_dict, reason_str), ...].
        """
        import openai
        import json as _json

        lines = []
        for i, r in enumerate(batch, 1):
            lines.append(f"{i}. {r['title']}")
            if r.get('authors'):
                lines.append(f"   作者: {r['authors']}")
            if r.get('year'):
                lines.append(f"   年份: {r['year']}")
            if r.get('publication'):
                lines.append(f"   期刊: {r['publication']}")
            if r.get('snippet'):
                lines.append(f"   摘要: {r['snippet'][:200]}")

        prompt = """你是一个建筑学研究助手。请从以下Google Scholar搜索结果中，筛选出**可能包含中国之外的中式建筑案例**的文章。

中国之外的中式建筑包括：海外中国园林、海外中国寺庙、海外唐人街建筑、欧洲中国风建筑(Chinoiserie)、海外中式餐厅/茶馆建筑、海外中式文化中心/使馆建筑等。

排除中国境内的建筑（包括中国境内的西式建筑）。

请返回JSON，包含相关文章编号和筛选理由（一句话说明为什么相关）：
{"relevant": [1, 5], "reasons": {"1": "讨论了旧金山中国园林的设计案例", "5": "分析了伦敦唐人街的建筑特征"}}

""" + "\n".join(lines) + "\n\n只返回JSON："

        try:
            _base_url, _model, _ = config.get_llm("search_parse")
            client = openai.OpenAI(
                base_url=_base_url,
                api_key=api_key,
                timeout=120,
            )
            response = client.chat.completions.create(
                model=_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=400,
            )
            content = response.choices[0].message.content.strip()
            import re
            match = re.search(r'\{[\s\S]*\}', content)
            if match:
                parsed = _json.loads(match.group())
                if isinstance(parsed, dict):
                    relevant = parsed.get("relevant", [])
                    reasons = parsed.get("reasons", {})
                    out = []
                    for rid in relevant:
                        idx = rid - 1
                        if 0 <= idx < len(batch):
                            reason = reasons.get(str(rid), "")
                            out.append((batch[idx], reason))
                    return out
        except Exception as e:
            print(f"[DS screen] error: {e}")

        return []

    def _screen_batch_with_gemini(self, batch, api_key):
        """Send a batch of Scholar results to Gemini, return only those likely
        to contain overseas Chinese architecture cases (outside China).

        Returns [(result_dict, reason_str), ...].
        """
        import openai
        import json as _json

        lines = []
        for i, r in enumerate(batch, 1):
            lines.append(f"{i}. {r['title']}")
            if r.get('authors'):
                lines.append(f"   作者: {r['authors']}")
            if r.get('year'):
                lines.append(f"   年份: {r['year']}")
            if r.get('publication'):
                lines.append(f"   期刊: {r['publication']}")
            if r.get('snippet'):
                lines.append(f"   摘要: {r['snippet'][:200]}")

        prompt = """你是一个建筑学研究助手。请从以下Google Scholar搜索结果中，筛选出**可能包含中国之外的中式建筑案例**的文章。

中国之外的中式建筑包括：海外中国园林、海外中国寺庙、海外唐人街建筑、欧洲中国风建筑(Chinoiserie)、海外中式餐厅/茶馆建筑、海外中式文化中心/使馆建筑等。

排除中国境内的建筑（包括中国境内的西式建筑）。

请返回JSON，包含相关文章编号和筛选理由（一句话说明为什么相关）：
{"relevant": [1, 5], "reasons": {"1": "讨论了旧金山中国园林的设计案例", "5": "分析了伦敦唐人街的建筑特征"}}

""" + "\n".join(lines) + "\n\n只返回JSON："

        import datetime, os
        log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "gemini_debug.log")
        try:
            _base_url, _model, _ = config.get_llm("field_completion")
            client = openai.OpenAI(
                base_url=_base_url,
                api_key=api_key,
                timeout=120,
            )
            response = client.chat.completions.create(
                model=_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=2000,
            )
            content = response.choices[0].message.content.strip()
            with open(log_path, "a") as f:
                f.write(f"\n=== {datetime.datetime.now()} ===\n")
                f.write(f"Raw response:\n{content[:500]}\n")
            import re
            match = re.search(r'\{[\s\S]*\}', content)
            if match:
                parsed = _json.loads(match.group())
                if isinstance(parsed, dict):
                    relevant = parsed.get("relevant", [])
                    reasons = parsed.get("reasons", {})
                    with open(log_path, "a") as f:
                        f.write(f"Relevant indices: {relevant}\n")
                    out = []
                    for rid in relevant:
                        idx = rid - 1
                        if 0 <= idx < len(batch):
                            reason = reasons.get(str(rid), "")
                            out.append((batch[idx], reason))
                    return out
                else:
                    with open(log_path, "a") as f:
                        f.write(f"Parsed JSON not a dict: {type(parsed)}\n")
            else:
                with open(log_path, "a") as f:
                    f.write(f"No JSON found in response\n")
        except Exception as e:
            import traceback
            with open(log_path, "a") as f:
                f.write(f"\n=== {datetime.datetime.now()} ERROR ===\n")
                f.write(f"Error: {e}\n{traceback.format_exc()}\n")

        return []

    GEMINI_API_KEY = config.get_llm("field_completion")[2]

    def _screen_batch(self, batch, api_key):
        """Dispatch to DS or Gemini based on user selection."""
        if getattr(self, '_use_gemini', False):
            return self._screen_batch_with_gemini(batch, self.GEMINI_API_KEY)
        return self._screen_batch_with_ds(batch, api_key)

    def _on_screen_change(self):
        self._use_gemini = self._screen_var.get() == "Gemini"
        self._refresh_from_db()

    def _do_search(self):
        query = self._keyword_var.get().strip()
        if not query:
            messagebox.showinfo("提示", "请输入搜索关键词")
            return

        from utils.web_search import scholar_search

        self._scholar_log("开始", f"单次搜索: {query}")

        api_key = config.get_llm("search_parse")[2]
        all_raw = []       # All raw results (saved to DB for raw results view)
        all_relevant = []  # DS-screened results (shown in literature list)

        EMPTY_STOP = 2    # 连续N页无DS相关 → 停止
        MAX_PAGES = 20     # 安全上限

        page = 1
        empty_count = 0
        while page <= MAX_PAGES:

            self._status_var.set(f"Scholar搜索中: {query} (第{page}页) ...")
            self.win.update()

            try:
                batch = scholar_search(query, num=10, page=page)
            except Exception as e:
                err_msg = str(e)
                if page == 1:
                    if "timeout" in err_msg.lower() or "timed out" in err_msg.lower():
                        messagebox.showerror("Scholar搜索超时",
                            f"Scholar API 3分钟未响应，已超时。\n"
                            f"建议：检查代理是否运行，或稍后重试。")
                    elif "Connection refused" in err_msg or "Errno 61" in err_msg:
                        messagebox.showerror("Scholar连接失败",
                            f"代理连接被拒绝。\n请确认代理已启动。")
                    else:
                        messagebox.showerror("Scholar搜索失败", f"{err_msg}")
                    self._status_var.set("搜索失败")
                    return
                else:
                    print(f"[Scholar] page={page} error: {e}")
                    self._scholar_log("错误", f"单次搜索 [{query}] 第{page}页: {e}")
                    break

            if not batch:
                self._scholar_log("停止", f"单次搜索 [{query}] 第{page}页API无结果，停止翻页")
                break

            # Save ALL raw results to database immediately
            all_raw.extend(batch)
            engine = "Gemini" if self._use_gemini else "DS"
            try:
                import database as db
                db.insert_scholar_results(query, batch, f"Scholar·{engine}")
            except Exception as e:
                print(f"[LiteratureSearch] save raw error: {e}")
            self._status_var.set(f"{engine}筛选中: 第{page}页 ({len(batch)}条) ...")
            self.win.update()

            relevant = self._screen_batch(batch, api_key)
            print(f"[{engine} screen] page={page}: {len(batch)}条 -> {len(relevant)}条相关")

            if relevant:
                empty_count = 0
                # Mark DS-selected in database with reasons
                reasons = {}
                for r, reason in relevant:
                    key = r.get("link", "") or r.get("title", "")
                    if key:
                        reasons[key] = reason
                ds_links = [r.get("link", "") for r, _ in relevant if r.get("link")]
                ds_titles = [r.get("title", "") for r, _ in relevant if not r.get("link")]
                try:
                    db.mark_scholar_results_gemini_selected(ds_links, ds_titles, reasons)
                except Exception as e:
                    print(f"[LiteratureSearch] ds mark error: {e}")

                label = f"[Scholar·{engine}] {query}"
                self._add_batch_to_tree(label, relevant, f"Scholar·{engine}", group_name="单次搜索")
                all_relevant.extend(r for r, _ in relevant)
            else:
                empty_count += 1
                if empty_count >= EMPTY_STOP:
                    self._scholar_log("停止",
                        f"单次搜索 [{query}] 连续{EMPTY_STOP}页无{engine}相关，停止翻页")
                    break

            self._scholar_log("进度",
                f"单次搜索 [{query}] 第{page}页: {len(batch)}条原始 → {len(relevant)}条{engine}相关")
            page += 1

        total = len(self._results)
        self._scholar_log("完成",
            f"单次搜索 [{query}]: {len(all_raw)}条原始 → {len(all_relevant)}条{engine}相关（累计 {total} 条）")
        self._stats_label.config(text=(
            f"本次: {query} → Scholar {len(all_raw)}条 → {engine}筛选 {len(all_relevant)}条 | 累计: {total} 条"
        ))
        self._status_var.set(
            f"「{query}」完成，Scholar {len(all_raw)}条 → {engine}筛选后 +{len(all_relevant)}条（累计 {total} 条）")
        self._update_filter_queries()
        self._update_bottom_count()

    def _scholar_log(self, status, msg, log_file=None):
        """Write a log entry to the Scholar log file in real-time."""
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] [{status}] {msg}"
        target = log_file or self._log_file
        try:
            with self._log_lock:
                with open(target, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
        except Exception:
            pass

    # ---- Group management ----

    def _add_group(self, name=None):
        if name is None:
            name = f"队列{len(self._groups) + 1}"
        g = {"name": name, "keywords": [], "running": False,
             "paused": False, "stop": False, "progress": 0, "max": 0,
             "failed_keywords": []}
        self._groups.append(g)
        idx = len(self._groups) - 1
        # Create a dummy frame as tab content (notebook requires a widget)
        tab_frame = ttk.Frame(self._group_notebook, width=0, height=0)
        self._group_notebook.add(tab_frame, text=name)
        self._group_notebook.select(idx)
        self._activate_group(idx)
        self._save_queue_keywords()

    def _remove_group(self):
        if len(self._groups) <= 1:
            messagebox.showinfo("提示", "至少保留一个队列组")
            return
        g = self._get_active_group()
        if g and g["running"]:
            messagebox.showinfo("提示", "请先停止该组的搜索再删除")
            return
        idx = self._active_group
        self._groups.pop(idx)
        self._group_notebook.forget(idx)
        new_idx = min(idx, len(self._groups) - 1)
        self._group_notebook.select(new_idx)
        self._activate_group(new_idx)
        self._save_queue_keywords()

    def _rename_group(self):
        import tkinter.simpledialog as sd
        g = self._get_active_group()
        if not g:
            return
        new_name = sd.askstring("重命名", "输入新名称：", initialvalue=g["name"], parent=self.win)
        if new_name and new_name.strip():
            g["name"] = new_name.strip()
            self._group_notebook.tab(self._active_group, text=g["name"])
            self._save_queue_keywords()

    def _on_group_switch(self, event):
        new_idx = self._group_notebook.index("current")
        if new_idx == self._active_group:
            return
        # Save current group keywords
        if self._active_group is not None and self._active_group < len(self._groups):
            self._sync_listbox_to_group()
        self._activate_group(new_idx)

    def _activate_group(self, idx):
        self._active_group = idx
        g = self._groups[idx]
        # Load keywords into listbox
        self._queue_listbox.delete(0, "end")
        for kw in g["keywords"]:
            self._queue_listbox.insert("end", kw)
        # Update controls
        self._refresh_group_controls()

    def _get_active_group(self):
        if self._active_group is None or self._active_group >= len(self._groups):
            return None
        return self._groups[self._active_group]

    def _sync_listbox_to_group(self):
        g = self._get_active_group()
        if g:
            g["keywords"] = list(self._queue_listbox.get(0, "end"))

    def _sync_group_to_listbox(self):
        g = self._get_active_group()
        if g:
            self._queue_listbox.delete(0, "end")
            for kw in g["keywords"]:
                self._queue_listbox.insert("end", kw)

    def _refresh_group_controls(self):
        g = self._get_active_group()
        if not g:
            return
        if g["running"]:
            self._queue_start_btn.config(state="disabled")
            self._queue_pause_btn.config(state="normal")
            self._queue_pause_btn.config(text="继续" if g["paused"] else "暂停")
            self._queue_progress["maximum"] = g["max"]
            self._queue_progress["value"] = g["progress"]
            self._queue_progress_var.set(f"{g['progress']}/{g['max']}")
        else:
            self._queue_start_btn.config(state="normal")
            self._queue_pause_btn.config(state="disabled")
            self._queue_progress["value"] = 0
            self._queue_progress_var.set("")

    # ---- Queue operations (on active group) ----

    def _queue_add(self):
        text = self._keyword_var.get().strip()
        if not text:
            return
        g = self._get_active_group()
        if not g:
            return
        for kw in text.replace(",", "\n").replace("，", "\n").split("\n"):
            kw = kw.strip()
            if kw and kw not in self._queue_listbox.get(0, "end"):
                self._queue_listbox.insert("end", kw)
        self._sync_listbox_to_group()
        self._save_queue_keywords()

    def _queue_del(self):
        for i in reversed(self._queue_listbox.curselection()):
            self._queue_listbox.delete(i)
        self._sync_listbox_to_group()
        self._save_queue_keywords()

    def _queue_clear(self):
        self._queue_listbox.delete(0, "end")
        self._sync_listbox_to_group()
        self._save_queue_keywords()

    _QUEUE_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                                "data", "scholar_queue_keywords.json")

    def _save_queue_keywords(self):
        try:
            import json
            # Save all groups
            data = []
            for g in self._groups:
                # Sync current listbox to active group before saving
                if self._groups.index(g) == self._active_group:
                    self._sync_listbox_to_group()
                data.append({"name": g["name"], "keywords": g["keywords"]})
            with open(self._QUEUE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception:
            pass

    def _load_queue_keywords(self):
        try:
            import json
            if os.path.exists(self._QUEUE_FILE):
                with open(self._QUEUE_FILE, encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list) and data and isinstance(data[0], dict):
                    return data  # new multi-group format
                elif isinstance(data, list):
                    return [{"name": "队列1", "keywords": data}]
        except Exception:
            pass
        return [{"name": "队列1", "keywords": []}]

    def _on_close(self):
        self._sync_listbox_to_group()
        self._save_queue_keywords()
        self.win.destroy()

    def _save_raw_to_db(self, query, results, engine=""):
        try:
            import database as db
            db.insert_scholar_results(query, results, engine)
        except Exception as e:
            print(f"[LiteratureSearch] save raw error: {e}")

    def _mark_ds_selected(self, links, titles, reasons=None):
        try:
            import database as db
            db.mark_scholar_results_gemini_selected(links, titles, reasons)
        except Exception as e:
            print(f"[LiteratureSearch] ds mark error: {e}")

    def _queue_start(self):
        g = self._get_active_group()
        if not g:
            return
        if not self._queue_listbox.get(0, "end"):
            messagebox.showinfo("提示", "请先在队列中添加关键词")
            return

        self._sync_listbox_to_group()
        g["running"] = True
        g["paused"] = False
        g["stop"] = False
        g["failed_keywords"] = []
        g["max"] = len(g["keywords"])
        g["progress"] = 0

        # Per-group log file
        import datetime
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
        os.makedirs(log_dir, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = g["name"].replace("/", "_").replace(" ", "_")
        g["log_file"] = os.path.join(log_dir, f"scholar_{safe_name}_{ts}.log")

        self._refresh_group_controls()

        all_kw = list(g["keywords"])
        self._save_checkpoint(g, 0)
        self._save_queue_keywords()

        t = threading.Thread(target=self._queue_run, args=(g, all_kw), daemon=True)
        t.start()

    def _queue_run(self, group, keywords):
        from utils.web_search import scholar_search

        api_key = config.get_llm("search_parse")[2]
        group_name = group["name"]
        failed_keywords = []
        total = len(keywords)
        for i, kw in enumerate(keywords):
            if group["stop"]:
                remaining = keywords[i:]
                for rkw in remaining:
                    self._scholar_log("未处理", rkw, log_file=group.get("log_file"))
                self._clear_checkpoint(group)
                break

            while group["paused"] and not group["stop"]:
                import time
                time.sleep(0.5)

            if group["stop"]:
                remaining = keywords[i:]
                for rkw in remaining:
                    self._scholar_log("未处理", rkw, log_file=group.get("log_file"))
                self._clear_checkpoint(group)
                break

            if kw not in group["keywords"]:
                self._scholar_log("已删除", f"{kw} — 跳过", log_file=group.get("log_file"))
                self._save_checkpoint(group, i + 1)
                continue

            processed = i + 1
            group["progress"] = processed

            self.win.after(0, lambda _p=processed, _k=kw, _t=total, _gn=group_name: (
                self._queue_progress_var.set(f"[{_gn}] {_p}/{_t}"),
                self._queue_progress.configure(value=_p),
            ))

            all_relevant = []
            stop_reason = ""  # track why we stopped
            try:
                EMPTY_STOP = 2
                MAX_PAGES = 20

                page = 1
                empty_count = 0
                while page <= MAX_PAGES:
                    if group["stop"]:
                        stop_reason = "被停止"
                        break
                    batch = scholar_search(kw, num=10, page=page)
                    if not batch:
                        if page == 1:
                            stop_reason = "API无返回"
                        break
                    engine = "Gemini" if self._use_gemini else "DS"
                    self.win.after(0, lambda _k=kw, _r=list(batch), _e=f"Scholar·{engine}":
                        self._save_raw_to_db(_k, _r, _e))
                    relevant = self._screen_batch(batch, api_key)
                    if relevant:
                        empty_count = 0
                        all_relevant.extend(r for r, _ in relevant)
                        label = f"[{group_name}·Scholar·{engine}] {kw}"
                        reasons = {}
                        for r, reason in relevant:
                            key = r.get("link", "") or r.get("title", "")
                            if key:
                                reasons[key] = reason
                        ds_links = [r.get("link", "") for r, _ in relevant if r.get("link")]
                        ds_titles = [r.get("title", "") for r, _ in relevant if not r.get("link")]
                        self.win.after(0, lambda _l=ds_links, _t=ds_titles, _rs=reasons:
                            self._mark_ds_selected(_l, _t, _rs))
                        self.win.after(0, lambda _q=label, _r=list(relevant), _e=f"Scholar·{engine}", _gn=group_name: (
                            self._add_batch_to_tree(_q, _r, _e, group_name=_gn),
                            self._update_filter_queries(),
                            self._update_bottom_count(),
                        ))
                    else:
                        empty_count += 1
                        if empty_count >= EMPTY_STOP:
                            stop_reason = f"连续{EMPTY_STOP}页无AI相关"
                            break
                    page += 1
                if not stop_reason and page > MAX_PAGES:
                    stop_reason = "已达最大页数"
            except Exception as e:
                err = str(e)[:80]
                failed_keywords.append(kw)
                stop_reason = f"异常: {err}"
                self._scholar_log("失败", f"{kw} — {err}", log_file=group.get("log_file"))

            # One clean log line per keyword
            count = len(all_relevant)
            if stop_reason == "被停止":
                self._scholar_log("完成", f"{kw} — {count}条 (被中断)", log_file=group.get("log_file"))
            else:
                self._scholar_log("完成", f"{kw} — {count}条", log_file=group.get("log_file"))

            self.win.after(0, lambda _p=processed, _gn=group_name:
                self._stats_label.config(text=f"[{_gn}] 已处理{_p}个 | 累计: {len(self._results)} 条"))

            self._save_checkpoint(group, processed)

        group["running"] = False
        group["paused"] = False
        group["failed_keywords"] = failed_keywords
        # Keep only unprocessed + failed keywords
        processed_keywords = set(keywords[:i]) if group["stop"] else set(keywords)
        remaining = [kw for kw in group["keywords"] if kw in failed_keywords or kw not in processed_keywords]
        group["keywords"] = remaining
        self.win.after(0, lambda g=group, fk=failed_keywords: self._on_queue_done(g, fk))

    def _on_queue_done(self, group, failed_keywords=None):
        gname = group["name"]
        self._refresh_group_controls()
        # Refresh listbox if this group is active
        if self._get_active_group() is group:
            self._sync_group_to_listbox()
        self._save_queue_keywords()
        total = len(self._results)
        self._stats_label.config(text=f"[{gname}] 完成 | 累计: {total} 条")
        self._status_var.set(f"[{gname}] 队列完成，共 {total} 条结果")
        self._clear_checkpoint(group)
        self.win.lift()
        if failed_keywords:
            messagebox.showinfo("队列完成",
                f"[{gname}] Scholar队列搜索完成，共 {total} 条结果。\n失败: {len(failed_keywords)} 个关键词，已保留可重试")
        else:
            messagebox.showinfo("队列完成", f"[{gname}] Scholar队列搜索完成，共 {total} 条结果。")

    _CHECKPOINT_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                                     "data", "scholar_queue_checkpoint.json")

    def _save_checkpoint(self, group, done_count):
        """Save checkpoint for a specific group. Merges with other groups' checkpoints."""
        try:
            import json
            os.makedirs(os.path.dirname(self._CHECKPOINT_FILE), exist_ok=True)
            all_cp = self._load_all_checkpoints()
            keywords = group.get("keywords", [])
            remaining = keywords[done_count:]
            all_cp[group["name"]] = {
                "keywords": remaining,
                "done": done_count,
                "total": len(keywords),
            }
            with open(self._CHECKPOINT_FILE, "w") as f:
                json.dump(all_cp, f, ensure_ascii=False)
        except Exception:
            pass

    def _load_all_checkpoints(self):
        """Return {group_name: {keywords, done, total}} dict."""
        try:
            import json
            if os.path.exists(self._CHECKPOINT_FILE):
                with open(self._CHECKPOINT_FILE) as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    # Check if old single-group format
                    if "keywords" in data and "done" in data:
                        return {"队列1": data}
                    return data
        except Exception:
            pass
        return {}

    def _load_checkpoint(self):
        """Legacy compatibility: return first group's checkpoint."""
        all_cp = self._load_all_checkpoints()
        if all_cp:
            return list(all_cp.values())[0]
        return None

    def _clear_checkpoint(self, group):
        """Remove a group's checkpoint. Delete file if no groups remain."""
        try:
            all_cp = self._load_all_checkpoints()
            all_cp.pop(group["name"], None)
            if all_cp:
                import json
                with open(self._CHECKPOINT_FILE, "w") as f:
                    json.dump(all_cp, f, ensure_ascii=False)
            elif os.path.exists(self._CHECKPOINT_FILE):
                os.remove(self._CHECKPOINT_FILE)
        except Exception:
            pass

    def _queue_stop_search(self):
        g = self._get_active_group()
        if g:
            g["stop"] = True
            g["paused"] = False
            self._clear_checkpoint(g)

    def _queue_pause(self):
        g = self._get_active_group()
        if not g:
            return
        if g["paused"]:
            g["paused"] = False
            self._queue_pause_btn.config(text="暂停")
        else:
            g["paused"] = True
            self._queue_pause_btn.config(text="继续")

    # ---- Tree management ----

    def _add_batch_to_tree(self, query_label, results, engine="", group_name=""):
        """Add screened results to tree. results is [(result_dict, reason_str), ...]."""
        if not hasattr(self, '_batch_parents'):
            self._batch_parents = {}
        engine = engine or "Scholar"
        # Use group_name as top-level parent, fall back to engine
        batch_key = group_name or engine
        if batch_key not in self._batch_parents:
            batch_iid = self._tree.insert("", "end", iid=f"B_{batch_key}",
                                          open=True, text=f"  {batch_key}")
            self._batch_parents[batch_key] = batch_iid
        if query_label not in self._query_groups:
            self._query_groups[query_label] = []
            parent_iid = f"G_{query_label}"
            self._tree.insert(self._batch_parents[batch_key], "end", iid=parent_iid, open=False,
                              text=f"  {query_label}", values=(query_label, "", ""))
        else:
            parent_iid = f"G_{query_label}"
        for entry in results:
            if isinstance(entry, tuple) and len(entry) == 2:
                r, reason = entry
            else:
                r, reason = entry, ""
            self._idx_counter += 1
            r["_query"] = query_label
            r["_idx"] = self._idx_counter
            r["_gemini_reason"] = reason
            r["_screen_engine"] = engine
            if "created_at" not in r:
                import datetime
                r["created_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._results.append(r)
            self._query_groups[query_label].append(r)
            self._tree.insert(parent_iid, "end", iid=f"L_{self._idx_counter}",
                              text=str(self._idx_counter),
                              values=("",
                                      r.get("title", ""),
                                      reason,
                                      ""))
        count = len(self._query_groups[query_label])
        self._tree.item(parent_iid, text=f"  {query_label} ({count}条)")

    def _on_select(self, event):
        sel = self._tree.selection()
        if not sel:
            self._expand_btn.configure(text="展开/合上")
            return
        iid = sel[0]
        self._update_expand_btn_text()
        if not iid.startswith("L_"):
            return
        idx = int(iid.split("_")[1])
        r = None
        for item in self._results:
            if item.get("_idx") == idx:
                r = item
                break
        if not r:
            return
        detail = f"""标题：{r.get('title', '')}
作者：{r.get('authors', '')}
年份：{r.get('year', '')}
期刊/会议：{r.get('publication', '')}
被引次数：{r.get('citedBy', '')}
搜索时间：{r.get('created_at', '')}
筛选理由：{r.get('_gemini_reason', r.get('gemini_reason', ''))}
标记：{r.get('tags', '')}

摘要：
{r.get('snippet', '')}

PDF链接：{r.get('pdfUrl', '')}
网页链接：{r.get('link', '')}"""
        self._detail_text.configure(state="normal")
        self._detail_text.delete("1.0", "end")
        self._detail_text.insert("1.0", detail)
        self._detail_text.configure(state="disabled")

    def _expand_selected(self):
        """Expand/collapse selected batch or query-group parent nodes."""
        for iid in self._tree.selection():
            if iid.startswith("B_") or iid.startswith("G_"):
                self._tree.item(iid, open=not self._tree.item(iid, "open"))
        self._update_expand_btn_text()

    def _update_expand_btn_text(self):
        g_items = [iid for iid in self._tree.selection() if iid.startswith("B_") or iid.startswith("G_")]
        if not g_items:
            self._expand_btn.configure(text="展开/合上")
            return
        states = [self._tree.item(iid, "open") for iid in g_items]
        if all(states):
            self._expand_btn.configure(text="合上")
        elif not any(states):
            self._expand_btn.configure(text="展开")
        else:
            self._expand_btn.configure(text="展开/合上")

    # ---- Save to literature table ----

    def _save_to_library(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择要入库的文献")
            return
        import database as db
        saved = 0
        for iid in sel:
            if not iid.startswith("L_"):
                continue
            idx = int(iid.split("_")[1])
            for r in self._results:
                if r.get("_idx") == idx:
                    existing = db.search_literature_by_title(r.get("title", ""))
                    if existing:
                        # Skip duplicates silently
                        continue
                    db.insert_literature({
                        "title": r.get("title", ""),
                        "authors": r.get("authors", ""),
                        "pub_date": str(r.get("year", "")),
                        "source_url": r.get("link", "") or r.get("pdfUrl", ""),
                        "source_engine": "Google Scholar",
                        "summary": r.get("snippet", ""),
                        "notes": r.get("publication", ""),
                    })
                    saved += 1
                    break
        if saved:
            self._status_var.set(f"已入库 {saved} 条文献")
            db.log_operation("文献筛选", "入库选中", f"入库{saved}条文献", saved)
            messagebox.showinfo("入库完成", f"已将 {saved} 条文献保存到文献库。")
        else:
            messagebox.showinfo("提示", "没有新文献可入库（可能已存在）。")

    # ---- Re-fill DS reasons ----

    def _refill_reasons(self):
        """Re-screen existing ds_selected records to fill missing reasons."""
        import database as db
        import threading

        missing = db.get_scholar_results_missing_reason()
        if not missing:
            messagebox.showinfo("提示", "没有缺少筛选理由的记录。")
            return

        if not messagebox.askyesno("补充筛选理由",
                                   f"有 {len(missing)} 条记录缺少筛选理由。\n"
                                   f"将使用DeepSeek逐批补充，是否继续？"):
            return

        self._status_var.set(f"补充筛选中: 共 {len(missing)} 条 ...")
        self.win.update()

        api_key = config.get_llm("search_parse")[2]

        def run():
            import openai
            import json as _json
            import re

            batch_size = 10
            done = 0
            for i in range(0, len(missing), batch_size):
                batch = missing[i:i + batch_size]

                lines = []
                for j, r in enumerate(batch, 1):
                    lines.append(f"{j}. {r['title']}")
                    if r.get('authors'):
                        lines.append(f"   作者: {r['authors']}")
                    if r.get('year'):
                        lines.append(f"   年份: {r['year']}")
                    if r.get('publication'):
                        lines.append(f"   期刊: {r['publication']}")
                    if r.get('snippet'):
                        lines.append(f"   摘要: {r['snippet'][:200]}")

                prompt = (
                    "你是一个建筑学研究助手。以下是已被确认为包含海外中式建筑案例的文章。\n"
                    "请为每篇文章写一句筛选理由（一句话说明它为什么与海外中式建筑相关）。\n"
                    "返回JSON: {\"reasons\": {\"1\": \"理由\", \"2\": \"理由\", ...}}\n"
                    "只返回JSON。\n\n"
                    + "\n".join(lines)
                )

                try:
                    _base_url, _model, _ = config.get_llm("search_parse")
                    client = openai.OpenAI(
                        base_url=_base_url,
                        api_key=api_key,
                        timeout=120,
                    )
                    response = client.chat.completions.create(
                        model=_model,
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.1,
                        max_tokens=500,
                    )
                    content = response.choices[0].message.content.strip()
                    match = re.search(r'\{[\s\S]*\}', content)
                    if match:
                        parsed = _json.loads(match.group())
                        reasons = parsed.get("reasons", {})
                        for j_str, reason in reasons.items():
                            j = int(j_str) - 1
                            if 0 <= j < len(batch):
                                db.update_scholar_result_reason(batch[j]["id"], reason)
                                done += 1
                except Exception as e:
                    print(f"[refill_reasons] DS error: {e}")

                self.win.after(0, lambda _d=done: self._status_var.set(
                    f"补充筛选: {_d}/{len(missing)} 已完成"))

            # Reload tree from DB
            self.win.after(0, self._reload_tree)

        threading.Thread(target=run, daemon=True).start()

    def _reload_tree(self):
        """Clear tree and reload from database."""
        for item in self._tree.get_children():
            self._tree.delete(item)
        self._results.clear()
        self._query_groups.clear()
        self._idx_counter = 0
        self._load_from_db()
        self._status_var.set(f"就绪（已加载 {len(self._results)} 条）")

    # ---- Export ----

    def _export_excel(self):
        if not self._results:
            messagebox.showinfo("提示", "暂无结果")
            return
        data = []
        for r in self._results:
            data.append({
                "标题": r.get("title", ""),
                "作者": r.get("authors", ""),
                "年份": r.get("year", ""),
                "期刊/会议": r.get("publication", ""),
                "摘要": r.get("snippet", ""),
                "被引次数": r.get("citedBy", ""),
                "PDF链接": r.get("pdfUrl", ""),
                "网页链接": r.get("link", ""),
            })
        df = pd.DataFrame(data)
        from utils.import_export import export_to_excel
        export_to_excel(df, title="文献筛选结果")

    # ---- Deduplication ----

    def _deduplicate(self):
        if not self._results:
            return
        seen_titles = {}
        removed = 0
        removed_links = []
        removed_titles = []
        removed_items = []  # save for undo
        for r in self._results[:]:
            title = r.get("title", "").strip().lower()
            if title and title in seen_titles:
                iid = f"L_{r['_idx']}"
                try:
                    self._tree.delete(iid)
                except tk.TclError:
                    pass
                self._results.remove(r)
                removed_items.append(dict(r))
                link = r.get("link", "")
                if link:
                    removed_links.append(link)
                else:
                    removed_titles.append(r.get("title", ""))
                removed += 1
            elif title:
                seen_titles[title] = r["_idx"]
        # Save undo data
        if removed_items:
            self._undo_stack.append({
                "type": "dedup",
                "items": removed_items,
                "removed_links": removed_links,
                "removed_titles": removed_titles,
            })
            if len(self._undo_stack) > 10:
                self._undo_stack.pop(0)
            n = len(self._undo_stack)
        self._undo_btn.config(state="normal", text=f"撤销({n})")
        self._undo_log_var.set(f"上次: 去重 {len(removed_items)}条")
        # Sync to database
        import database as db
        if removed_links or removed_titles:
            db.mark_scholar_results_status(removed_links, removed_titles, "deduped")
        for qlabel in self._query_groups:
            self._query_groups[qlabel] = [r for r in self._query_groups[qlabel] if r in self._results]
            count = len(self._query_groups[qlabel])
            self._tree.item(f"G_{qlabel}", text=f"  {qlabel} ({count}条)")
        self._update_bottom_count()
        self._status_var.set(f"去重完成，移除 {removed} 条重复")
        db.log_operation("文献筛选", "去重", f"移除{removed}条重复", removed)
        messagebox.showinfo("去重完成", f"共移除 {removed} 条重复文献。")

    # ---- Delete ----

    def _delete_selected(self):
        sel = self._tree.selection()
        if not sel:
            return
        import database as db
        del_links = []
        del_titles = []
        for iid in sel:
            if iid.startswith("L_"):
                idx = int(iid.split("_")[1])
                for r in self._results:
                    if r.get("_idx") == idx:
                        link = r.get("link", "")
                        if link:
                            del_links.append(link)
                        else:
                            del_titles.append(r.get("title", ""))
                        break
                self._results = [r for r in self._results if r.get("_idx") != idx]
            try:
                self._tree.delete(iid)
            except tk.TclError:
                pass
        if del_links or del_titles:
            db.mark_scholar_results_status(del_links, del_titles, "cleaned")
        # Update group counts
        for qlabel in list(self._query_groups.keys()):
            self._query_groups[qlabel] = [r for r in self._query_groups[qlabel] if r in self._results]
            if not self._query_groups[qlabel]:
                try:
                    self._tree.delete(f"G_{qlabel}")
                except tk.TclError:
                    pass
                del self._query_groups[qlabel]
            else:
                count = len(self._query_groups[qlabel])
                self._tree.item(f"G_{qlabel}", text=f"  {qlabel} ({count}条)")
        self._update_bottom_count()
        self._update_filter_queries()
        # Audit log
        db.log_operation("文献筛选", "数据清洗", f"删除{len(sel)}项", len(sel))

    def _open_selected(self):
        sel = self._tree.selection()
        for iid in sel:
            if iid.startswith("L_"):
                idx = int(iid.split("_")[1])
                for r in self._results:
                    if r.get("_idx") == idx:
                        link = r.get("pdfUrl") or r.get("link")
                        if link:
                            import subprocess
                            subprocess.run(["open", "-a", "/Applications/Google Chrome.app", link])
                        break

    # ---- Clean ----

    def _tag_selected(self):
        """Tag selected Scholar results."""
        sel = self._tree.selection()
        if not sel:
            return
        leaf_ids = []
        for iid in sel:
            if iid.startswith("L_"):
                try:
                    leaf_ids.append(int(iid.split("_")[1]))
                except ValueError:
                    pass
        if not leaf_ids:
            return
        tag = self._tag_var.get().strip()
        if not tag:
            messagebox.showinfo("提示", "请输入标记内容。")
            return

        import database as db
        links = []
        titles = []
        for r in self._results:
            if r["_idx"] in leaf_ids:
                r["tags"] = tag
                link = r.get("link", "")
                if link:
                    links.append(link)
                else:
                    titles.append(r.get("title", ""))
        if links or titles:
            db.mark_scholar_results_tags(links, titles, tag)
        for iid in sel:
            if iid.startswith("L_"):
                self._tree.set(iid, "tags", tag)
        self._status_var.set(f"已标记 {len(leaf_ids)} 条")

    def _clean_selected(self):
        """Mark selected Scholar results as 'cleaned' in DB and remove from tree."""
        sel = self._tree.selection()
        if not sel:
            return
        leaf_ids = []
        for iid in sel:
            if iid.startswith("L_"):
                try:
                    leaf_ids.append(int(iid.split("_")[1]))
                except ValueError:
                    pass
        if not leaf_ids:
            return
        count = len(leaf_ids)
        ok = messagebox.askyesno("数据清洗", f"确定清洗选中的 {count} 条数据？\n清洗后的数据将在原始结果中标记为「已清洗」。")
        if not ok:
            return

        import database as db
        links_to_clean = []
        titles_to_clean = []
        removed_items = []  # save for undo
        for r in self._results:
            if r["_idx"] in leaf_ids:
                removed_items.append(dict(r))
                link = r.get("link", "")
                if link:
                    links_to_clean.append(link)
                else:
                    titles_to_clean.append(r.get("title", ""))

        # Save undo data
        self._undo_stack.append({
            "type": "clean",
            "items": removed_items,
            "cleaned_links": links_to_clean,
            "cleaned_titles": titles_to_clean,
        })
        if len(self._undo_stack) > 10:
            self._undo_stack.pop(0)
        n = len(self._undo_stack)
        self._undo_btn.config(state="normal", text=f"撤销({n})")
        self._undo_log_var.set(f"上次: 清洗 {len(removed_items)}条")

        if links_to_clean or titles_to_clean:
            db.mark_scholar_results_status(links_to_clean, titles_to_clean, "cleaned")

        self._results = [r for r in self._results if r["_idx"] not in leaf_ids]
        for iid in sel:
            try:
                self._tree.delete(iid)
            except tk.TclError:
                pass
        for qlabel in list(self._query_groups.keys()):
            self._query_groups[qlabel] = [r for r in self._query_groups[qlabel] if r["_idx"] not in leaf_ids]
            if not self._query_groups[qlabel]:
                try:
                    self._tree.delete(f"G_{qlabel}")
                except tk.TclError:
                    pass
                del self._query_groups[qlabel]
            else:
                self._tree.item(f"G_{qlabel}", text=f"  {qlabel} ({len(self._query_groups[qlabel])}条)")
        self._update_bottom_count()
        self._update_filter_queries()
        self._status_var.set(f"已清洗 {count} 项，剩余 {len(self._results)} 条")

    def _undo_last(self):
        """Undo the last dedup or clean operation in Scholar (up to 10 steps)."""
        if not self._undo_stack:
            messagebox.showinfo("撤销", "没有可撤销的操作。")
            return

        data = self._undo_stack.pop()
        if not self._undo_stack:
            self._undo_btn.config(state="disabled", text="撤销")
        else:
            self._undo_btn.config(text=f"撤销({len(self._undo_stack)})")

        op_type = data["type"]
        items = data["items"]
        count = len(items)

        import database as db

        if op_type == "dedup":
            removed_links = data.get("removed_links", [])
            removed_titles = data.get("removed_titles", [])
            if removed_links or removed_titles:
                try:
                    db.mark_scholar_results_status(removed_links, removed_titles, "active")
                except Exception as e:
                    print(f"[undo] schol dedup undo error: {e}")

        elif op_type == "clean":
            cleaned_links = data.get("cleaned_links", [])
            cleaned_titles = data.get("cleaned_titles", [])
            if cleaned_links or cleaned_titles:
                try:
                    db.mark_scholar_results_status(cleaned_links, cleaned_titles, "ds_selected")
                except Exception as e:
                    print(f"[undo] schol clean undo error: {e}")

        # Restore to memory
        for r in items:
            if r not in self._results:
                self._results.append(r)
                self._idx_counter = max(self._idx_counter, r.get("_idx", 0))
            q = r.get("_query", "")
            if q not in self._query_groups:
                self._query_groups[q] = []
            if r not in self._query_groups[q]:
                self._query_groups[q].append(r)

        # Rebuild tree
        self._tree.delete(*self._tree.get_children())
        for q_label, group_items in self._query_groups.items():
            parent_iid = f"G_{q_label}"
            self._tree.insert("", "end", iid=parent_iid, open=False,
                              text=f"  {q_label} ({len(group_items)}条)",
                              values=(q_label, "", ""))
            for item in group_items:
                iid = f"L_{item['_idx']}"
                self._tree.insert(parent_iid, "end", iid=iid,
                                  text=str(item["_idx"]),
                                  values=("", item.get("title", ""),
                                          item.get("_gemini_reason", item.get("gemini_reason", ""))))

        self._update_bottom_count()
        self._update_filter_queries()
        self._status_var.set(f"已撤销，恢复 {count} 条")

    # ---- Raw results ----

    def _show_raw_results(self):
        """Show raw Scholar search results from database."""
        import database as db
        queries = db.get_scholar_result_queries()
        if not queries:
            messagebox.showinfo("原始结果", "暂无Scholar原始结果记录")
            return
        total_results = sum(q["count"] for q in queries)

        dlg = tk.Toplevel(self.win)
        dlg.title("Scholar原始搜索结果")
        dlg.geometry("1100x550")
        dlg.transient(self.win)

        paned = ttk.PanedWindow(dlg, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        # Left: batch tree (batch → query)
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=0)
        left_header = ttk.Frame(left_frame)
        left_header.pack(fill="x")
        ttk.Label(left_header, text="批次 / 搜索词", font=("", 10, "bold")).pack(side="left")

        # Build batch groups
        _batch_groups = {}  # engine_label -> {"queries": [...], "total": N}
        for q in queries:
            engine = q.get("screen_engine", "") or "Scholar"
            if engine not in _batch_groups:
                _batch_groups[engine] = {"queries": [], "total": 0}
            _batch_groups[engine]["queries"].append(q)
            _batch_groups[engine]["total"] += q["count"]

        query_tree = ttk.Treeview(left_frame, show="tree", selectmode="extended")
        query_tree.column("#0", width=250)
        q_scroll_y = ttk.Scrollbar(left_frame, orient="vertical", command=query_tree.yview)
        query_tree.configure(yscrollcommand=q_scroll_y.set)
        query_tree.pack(side="left", fill="both", expand=True)
        q_scroll_y.pack(side="right", fill="y")

        for batch_label, bg in _batch_groups.items():
            batch_iid = query_tree.insert("", "end", open=False,
                text=f"  {batch_label} ({len(bg['queries'])}个搜索词, {bg['total']}条)")
            for q in bg["queries"]:
                query_tree.insert(batch_iid, "end",
                    text=f"  {q['query']} ({q['count']}条)",
                    values=(q["query"], q.get("screen_engine", "")))

        _expand_btn = ttk.Button(left_header, text="展开")
        _expand_btn.pack(side="right", padx=1)
        def _toggle_tree():
            children = query_tree.get_children("")
            if not children:
                return
            any_open = any(query_tree.item(c, "open") for c in children)
            new_state = not any_open
            for c in children:
                query_tree.item(c, open=new_state)
            _expand_btn.configure(text="合上" if new_state else "展开")
        _expand_btn.configure(command=_toggle_tree)

        # Right: results table
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        right_header = ttk.Frame(right_frame)
        right_header.pack(fill="x")
        ttk.Label(right_header, text="原始条目", font=("", 10, "bold")).pack(side="left")
        _hint_var = tk.StringVar(value="← 请点击左侧搜索词查看结果")
        _hint_label = ttk.Label(right_header, textvariable=_hint_var,
                                foreground="gray", font=("", 9))
        _hint_label.pack(side="left", padx=15)

        # Search box
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill="x", pady=(2, 0))
        ttk.Label(search_frame, text="搜索：").pack(side="left")
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side="left", padx=5)
        ttk.Label(search_frame, text="（筛选标题）", font=("", 7), foreground="gray").pack(side="left")

        columns = ("title", "status", "created_at")
        detail_tree = ttk.Treeview(right_frame, columns=columns, show="headings")
        detail_tree.heading("title", text="标题")
        detail_tree.column("title", width=420)
        detail_tree.heading("status", text="状态")
        detail_tree.column("status", width=100)
        detail_tree.heading("created_at", text="检索时间")
        detail_tree.column("created_at", width=140)

        scroll_y = ttk.Scrollbar(right_frame, orient="vertical", command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=scroll_y.set)
        detail_tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        # Stash hidden fields for detail view, and all rows for filtering
        _raw_detail = {}
        _all_rows = []

        # Detail area for selected row (below paned, full width)
        detail_frame = ttk.LabelFrame(dlg, text=" 选中详情 ", padding=3)
        detail_frame.pack(fill="x", padx=5, pady=(3, 0))
        detail_text = tk.Text(detail_frame, height=6, wrap="word", state="disabled")
        detail_text.pack(fill="both")

        def _on_select(event):
            sel = detail_tree.selection()
            if not sel:
                return
            iid = sel[0]
            vals = detail_tree.item(iid, "values")
            extra = _raw_detail.get(iid, {})
            detail_text.configure(state="normal")
            detail_text.delete("1.0", "end")
            detail_text.insert("1.0", f"""标题：{vals[0]}
作者：{extra.get('authors', '')}
年份：{extra.get('year', '')}
期刊/会议：{extra.get('publication', '')}

摘要：
{extra.get('snippet', '')}

链接：{extra.get('link', '')}
PDF链接：{extra.get('pdf_url', '')}
被引次数：{extra.get('cited_by', '')}
状态：{vals[1]}
检索时间：{vals[2]}""")
            detail_text.configure(state="disabled")

        detail_tree.bind("<<TreeviewSelect>>", _on_select)

        def _apply_search():
            keyword = search_var.get().strip().lower()
            detail_tree.delete(*detail_tree.get_children())
            _raw_detail.clear()
            for r in _all_rows:
                if not keyword or keyword in r["title"].lower():
                    iid = detail_tree.insert("", "end", values=(
                        r["title"],
                        r["status"], r["created_at"] or ""))
                    _raw_detail[iid] = {
                        "id": r["id"],
                        "title": r["title"],
                        "authors": r.get("authors", ""),
                        "year": r.get("year", ""),
                        "publication": r.get("publication", ""),
                        "snippet": r["snippet"] or "",
                        "link": r["link"] or "",
                        "pdf_url": r.get("pdf_url", ""),
                        "cited_by": r.get("cited_by", ""),
                        "clean_status": r.get("clean_status", "active"),
                    }

        search_var.trace_add("write", lambda *_: _apply_search())

        def _get_children_iids(parent_iid):
            """Recursively get all child iids under a parent."""
            result = []
            for child in query_tree.get_children(parent_iid):
                result.append(child)
                result.extend(_get_children_iids(child))
            return result

        def on_select_query(event):
            sel = query_tree.selection()
            if not sel:
                return
            search_var.set("")
            _all_rows.clear()
            rows = []
            for sel_iid in sel:
                vals = query_tree.item(sel_iid, "values")
                children = query_tree.get_children(sel_iid)
                if children:
                    continue
                if vals:
                    q_name = vals[0]
                    q_engine = vals[1] if len(vals) > 1 else None
                    for r in db.get_scholar_results_by_query(q_name, q_engine):
                        r["_query_name"] = q_name
                        rows.append(r)
            for r in rows:
                status = r.get("clean_status", "active")
                status_map = {"active": "未选中", "ds_selected": "已选中",
                              "cleaned": "已清洗", "depuded": "已去重"}
                display_status = status_map.get(status, status)
                tags = r.get("tags", "")
                if tags:
                    display_status += f" [{tags}]"
                _all_rows.append({
                    "id": r["id"],
                    "title": r["title"],
                    "status": display_status,
                    "created_at": r["created_at"] or "",
                    "authors": r.get("authors", ""),
                    "year": r.get("year", ""),
                    "publication": r.get("publication", ""),
                    "snippet": r["snippet"] or "",
                    "link": r["link"] or "",
                    "pdf_url": r.get("pdf_url", ""),
                    "cited_by": r.get("cited_by", ""),
                    "clean_status": r.get("clean_status", "active"),
                    "tags": r.get("tags", ""),
                })
            if rows:
                _hint_var.set("")
            else:
                _hint_var.set("← 该搜索词无结果")
            _apply_search()

        query_tree.bind("<<TreeviewSelect>>", on_select_query)

        bottom = ttk.Frame(dlg, padding=5)
        bottom.pack(fill="x")

        def _collect_selected_queries(sel_iids):
            """Collect (query, screen_engine) pairs from selected tree items."""
            result = []
            for sel_iid in sel_iids:
                vals = query_tree.item(sel_iid, "values")
                children = query_tree.get_children(sel_iid)
                if children:
                    continue
                if vals:
                    result.append((vals[0], vals[1] if len(vals) > 1 else None))
            return result

        def _export_excel():
            q_sel = query_tree.selection()
            if q_sel:
                pairs = _collect_selected_queries(q_sel)
                rows = []
                labels = []
                for q_name, q_engine in pairs:
                    rows.extend(db.get_scholar_results_by_query(q_name, q_engine))
                    labels.append(q_name)
                label = "_".join(labels)[:50]
            else:
                rows = []
                for q in queries:
                    rows.extend(db.get_scholar_results_by_query(q["query"], q.get("screen_engine")))
                label = "全部"
            if not rows:
                messagebox.showinfo("提示", "没有可导出的数据")
                return
            status_map = {"active": "未选中", "ds_selected": "已选中",
                          "cleaned": "已清洗", "deduped": "已去重"}
            data = []
            for r in rows:
                data.append({
                    "标题": r.get("title", ""),
                    "作者": r.get("authors", ""),
                    "年份": r.get("year", ""),
                    "期刊/会议": r.get("publication", ""),
                    "摘要": r.get("snippet", ""),
                    "网页链接": r.get("link", ""),
                    "PDF链接": r.get("pdf_url", ""),
                    "被引次数": r.get("cited_by", ""),
                    "状态": status_map.get(r.get("clean_status", "active"), r.get("clean_status", "")),
                    "检索时间": r.get("created_at", ""),
                })
            df = pd.DataFrame(data)
            from utils.import_export import export_to_excel
            export_to_excel(df, title=f"Scholar原始结果_{label}")

        def _delete_selected():
            q_sel = query_tree.selection()
            d_sel = detail_tree.selection()
            if q_sel:
                pairs = _collect_selected_queries(q_sel)
                total_count = sum(q["count"] for q in queries if (q["query"], q.get("screen_engine")) in [(p[0], p[1] or "") for p in pairs])
                query_names = "、".join(p[0] for p in pairs)
                if not messagebox.askyesno("确认删除", f"确定删除 {len(pairs)} 条搜索记录「{query_names}」的全部结果？\n此操作不可撤销。"):
                    return
                for q_name, q_engine in pairs:
                    ids_to_del = [r["id"] for r in db.get_scholar_results_by_query(q_name, q_engine)]
                    if ids_to_del:
                        db.delete_scholar_results_by_ids(ids_to_del)
                for sel_iid in q_sel:
                    query_tree.delete(sel_iid)
                detail_tree.delete(*detail_tree.get_children(""))
            elif d_sel:
                ids_to_del = [_raw_detail.get(iid, {}).get("id") for iid in d_sel]
                ids_to_del = [x for x in ids_to_del if x]
                if not ids_to_del:
                    return
                if not messagebox.askyesno("确认删除", f"确定删除选中的 {len(ids_to_del)} 条原始结果？\n此操作不可撤销。"):
                    return
                db.delete_scholar_results_by_ids(ids_to_del)
                for iid in d_sel:
                    detail_tree.delete(iid)
            else:
                messagebox.showinfo("提示", "请先选中左侧搜索记录或右侧原始条目")

        ttk.Button(bottom, text="删除选中", command=_delete_selected).pack(side="left", padx=2)

        def _manual_select():
            d_sel = detail_tree.selection()
            if not d_sel:
                messagebox.showinfo("提示", "请先选中右侧要添加的条目")
                return
            reason_dlg = tk.Toplevel(dlg)
            reason_dlg.title("手动筛选理由")
            reason_dlg.geometry("420x120")
            reason_dlg.transient(dlg)
            ttk.Label(reason_dlg, text="请输入筛选理由（一句话说明为什么与海外中式建筑相关）：").pack(padx=10, pady=(10, 5))
            reason_var = tk.StringVar()
            reason_entry = ttk.Entry(reason_dlg, textvariable=reason_var, width=50)
            reason_entry.pack(padx=10, pady=5)
            reason_entry.focus_set()
            result = {"reason": None}

            def _ok():
                r = reason_var.get().strip()
                if not r:
                    messagebox.showwarning("提示", "请输入筛选理由")
                    return
                result["reason"] = r
                reason_dlg.destroy()

            btn_frame = ttk.Frame(reason_dlg)
            btn_frame.pack(pady=5)
            ttk.Button(btn_frame, text="确定", command=_ok).pack(side="left", padx=5)
            ttk.Button(btn_frame, text="取消", command=reason_dlg.destroy).pack(side="left", padx=5)
            reason_entry.bind("<Return>", lambda e: _ok())
            dlg.wait_window(reason_dlg)
            if result["reason"] is None:
                return

            reason = result["reason"]
            links = []
            titles = []
            reasons_map = {}
            for iid in d_sel:
                extra = _raw_detail.get(iid, {})
                link = extra.get("link", "")
                title = extra.get("title", "")
                if link:
                    links.append(link)
                    reasons_map[link] = reason
                if title:
                    titles.append(title)
                    reasons_map[title] = reason
            if not links and not titles:
                return
            db.mark_scholar_results_gemini_selected(links, titles, reasons_map)
            on_select_query(None)
            self._refresh_from_db()

        ttk.Button(bottom, text="选中", command=_manual_select).pack(side="left", padx=2)

        def _restore_selected():
            d_sel = detail_tree.selection()
            if not d_sel:
                messagebox.showinfo("提示", "请先选中右侧要还原的条目")
                return
            ids_to_restore = [_raw_detail.get(iid, {}).get("id") for iid in d_sel]
            ids_to_restore = [x for x in ids_to_restore if x]
            if not ids_to_restore:
                return
            if not messagebox.askyesno("确认还原", f"确定还原选中的 {len(ids_to_restore)} 条记录？\n状态将恢复为「未选中」。"):
                return
            db.restore_scholar_results_by_ids(ids_to_restore)
            on_select_query(None)

        def _show_saved_raw_matches():
            batches = db.get_lit_match_batches("raw")
            if not batches:
                messagebox.showinfo("提示", "没有已保存的匹配记录。")
                return

            m_win = tk.Toplevel(dlg)
            m_win.title(f"原始结果匹配记录 - {len(batches)} 次")
            m_win.geometry("500x400")
            m_win.transient(dlg)

            ttk.Label(m_win, text=f"原始结果匹配文献库 — 共 {len(batches)} 次匹配记录，双击查看：",
                      font=("", 10)).pack(padx=15, pady=(15, 5))

            tree_frame = ttk.Frame(m_win)
            tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

            cols = ("time", "count")
            tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
            tree.heading("time", text="匹配时间")
            tree.heading("count", text="匹配数量")
            tree.column("time", width=250)
            tree.column("count", width=100)

            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)

            for b in batches:
                tree.insert("", "end", values=(b["batch_id"], f"{b['cnt']} 对"))

            def _open_batch(event):
                sel = tree.selection()
                if not sel:
                    return
                idx = tree.index(sel[0])
                batch_id = batches[idx]["batch_id"]
                saved = db.get_lit_matches("raw", batch_id=batch_id)
                if saved:
                    _show_raw_batch_results(batch_id, saved)

            tree.bind("<Double-1>", _open_batch)

            btn_row = ttk.Frame(m_win)
            btn_row.pack(fill="x", padx=15, pady=10)
            ttk.Button(btn_row, text="重新匹配",
                       command=lambda: [m_win.destroy(), db.clear_lit_matches("raw"), _match_lit()]
                       ).pack(side="left", padx=2)
            ttk.Button(btn_row, text="关闭", command=m_win.destroy).pack(side="right", padx=2)

        def _show_raw_batch_results(batch_id, saved):
            r_win = tk.Toplevel(dlg)
            r_win.title(f"原始结果匹配 - {batch_id}")
            r_win.geometry("1050x550")
            r_win.transient(dlg)

            ttk.Label(r_win, text=f"匹配时间：{batch_id}，共 {len(saved)} 对：",
                      font=("", 10)).pack(padx=15, pady=(15, 5))

            m_frame = ttk.Frame(r_win)
            m_frame.pack(fill="both", expand=True, padx=15, pady=5)

            cols = ("lit_title", "raw_title", "similarity", "authors", "pub_date")
            m_tree = ttk.Treeview(m_frame, columns=cols, show="headings")
            m_tree.heading("lit_title", text="文献库标题")
            m_tree.heading("raw_title", text="原始结果标题")
            m_tree.heading("similarity", text="相似度")
            m_tree.heading("authors", text="作者")
            m_tree.heading("pub_date", text="发表时间")
            m_tree.column("lit_title", width=300)
            m_tree.column("raw_title", width=300)
            m_tree.column("similarity", width=60)
            m_tree.column("authors", width=130)
            m_tree.column("pub_date", width=80)

            m_vsb = ttk.Scrollbar(m_frame, orient="vertical", command=m_tree.yview)
            m_hsb = ttk.Scrollbar(m_frame, orient="horizontal", command=m_tree.xview)
            m_tree.configure(yscrollcommand=m_vsb.set, xscrollcommand=m_hsb.set)
            m_tree.grid(row=0, column=0, sticky="nsew")
            m_vsb.grid(row=0, column=1, sticky="ns")
            m_hsb.grid(row=1, column=0, sticky="ew")
            m_frame.grid_rowconfigure(0, weight=1)
            m_frame.grid_columnconfigure(0, weight=1)

            for m in saved:
                sim = m["similarity"]
                m_tree.insert("", "end", values=(
                    m["lit_title"],
                    m["matched_title"],
                    f"{sim:.0%}" if sim == 1.0 else f"{sim:.0%}",
                    m["lit_authors"] or "",
                    m["lit_pub_date"] or "",
                ))

            m_detail = tk.Text(r_win, wrap="word", font=("", 10), height=6, state="disabled")
            m_detail.pack(fill="x", padx=15, pady=5)

            def _m_select(event):
                sel = m_tree.selection()
                if not sel:
                    return
                idx = m_tree.index(sel[0])
                if idx < len(saved):
                    m = saved[idx]
                    sim = m["similarity"]
                    m_detail.configure(state="normal")
                    m_detail.delete("1.0", "end")
                    m_detail.insert("end", f"文献库标题：{m['lit_title']}\n")
                    m_detail.insert("end", f"原始结果标题：{m['matched_title']}\n")
                    m_detail.insert("end", f"相似度：{sim:.0%}\n" if sim == 1.0 else f"相似度：{sim:.0%}\n")
                    m_detail.insert("end", f"作者：{m['lit_authors'] or ''}\n")
                    m_detail.insert("end", f"发表时间：{m['lit_pub_date'] or ''}\n")
                    m_detail.configure(state="disabled")

            m_tree.bind("<<TreeviewSelect>>", _m_select)

            ttk.Button(r_win, text="关闭", command=r_win.destroy).pack(pady=10)

        def _match_lit():
            sel = query_tree.selection()
            if not sel:
                messagebox.showinfo("提示", "请先在左侧选择搜索记录。")
                return

            rows = []
            for sel_iid in sel:
                idx = query_tree.index(sel_iid)
                q = queries[idx]
                rows.extend(db.get_scholar_results_by_query(q["query"]))

            # Deduplicate raw results by title
            seen_titles = set()
            uniq_rows = []
            for r in rows:
                t = (r.get("title") or "").strip()
                if t and t not in seen_titles:
                    seen_titles.add(t)
                    uniq_rows.append(r)
            rows = uniq_rows

            if not rows:
                messagebox.showinfo("提示", "所选搜索记录没有原始结果。")
                return

            if db.has_lit_matches("raw"):
                _show_saved_raw_matches()
                return

            lit_records = db.get_all_literature()
            if not lit_records:
                messagebox.showinfo("提示", "文献库中没有文献。")
                return

            import difflib, threading

            # Progress popup
            popup = tk.Toplevel(dlg)
            popup.title("匹配中...")
            popup.geometry("350x100")
            popup.transient(dlg)
            popup.grab_set()
            ttk.Label(popup, text=f"正在匹配 {len(lit_records)} 条文献 × {len(rows)} 条原始结果...",
                      font=("", 10)).pack(padx=20, pady=(15, 5))
            prog = ttk.Progressbar(popup, mode="indeterminate")
            prog.pack(padx=20, fill="x", pady=(0, 15))
            prog.start()

            result_holder = [None]
            done_flag = [False]

            def _run():
                try:
                    match_list = []
                    total = len(lit_records)
                    for j, lit in enumerate(lit_records):
                        lit_title = (lit.get("title") or "").strip()
                        if not lit_title:
                            if j % 10 == 0:
                                result_holder[0] = ("progress", j + 1, total)
                            continue
                        lit_lower = lit_title.lower()
                        for r in rows:
                            rt = (r.get("title") or "").strip()
                            if not rt:
                                continue
                            rt_lower = rt.lower()
                            if lit_lower == rt_lower:
                                match_list.append((lit, r, 1.0))
                            elif lit_lower in rt_lower or rt_lower in lit_lower:
                                match_list.append((lit, r, 0.85))
                            else:
                                ratio = difflib.SequenceMatcher(None, lit_lower, rt_lower).ratio()
                                if ratio > 0.5:
                                    match_list.append((lit, r, round(ratio, 2)))
                        if j % 10 == 0:
                            result_holder[0] = ("progress", j + 1, total)
                    match_list.sort(key=lambda x: x[2], reverse=True)
                    result_holder[0] = ("ok", rows, lit_records, match_list)
                except Exception as e:
                    result_holder[0] = ("err", str(e))
                done_flag[0] = True

            threading.Thread(target=_run, daemon=True).start()

            def _poll():
                if done_flag[0]:
                    popup.destroy()
                    status = result_holder[0]
                    if status[0] == "err":
                        messagebox.showerror("匹配出错", f"匹配过程出错：{status[1]}")
                        return
                    _, rows_val, lit_recs, match_list = status
                    db.save_lit_matches("raw", match_list)
                    if not match_list:
                        messagebox.showinfo("匹配结果",
                            f"未找到匹配。\n文献库共 {len(lit_recs)} 条，原始结果 {len(rows_val)} 条。")
                        return

                    _show_match_result(dlg, match_list, lit_recs, rows_val)
                elif result_holder[0] and result_holder[0][0] == "progress":
                    _, cur, tot = result_holder[0]
                    popup.title(f"匹配中... {cur}/{tot}")
                    dlg.after(100, _poll)
                else:
                    dlg.after(100, _poll)

            dlg.after(100, _poll)

        def _show_match_result(parent, matches, lit_records, rows):
            m_win = tk.Toplevel(parent)
            m_win.title(f"原始结果匹配文献库 - {len(matches)} 对")
            m_win.geometry("1050x550")
            m_win.transient(parent)

            ttk.Label(m_win, text=f"文献库 {len(lit_records)} 条，原始结果 {len(rows)} 条，匹配 {len(matches)} 对：",
                      font=("", 10)).pack(padx=15, pady=(15, 5))

            m_frame = ttk.Frame(m_win)
            m_frame.pack(fill="both", expand=True, padx=15, pady=5)

            cols = ("lit_title", "raw_title", "similarity", "authors", "pub_date")
            m_tree = ttk.Treeview(m_frame, columns=cols, show="headings")
            m_tree.heading("lit_title", text="文献库标题")
            m_tree.heading("raw_title", text="原始结果标题")
            m_tree.heading("similarity", text="相似度")
            m_tree.heading("authors", text="作者")
            m_tree.heading("pub_date", text="发表时间")
            m_tree.column("lit_title", width=300)
            m_tree.column("raw_title", width=300)
            m_tree.column("similarity", width=60)
            m_tree.column("authors", width=130)
            m_tree.column("pub_date", width=80)

            m_vsb = ttk.Scrollbar(m_frame, orient="vertical", command=m_tree.yview)
            m_hsb = ttk.Scrollbar(m_frame, orient="horizontal", command=m_tree.xview)
            m_tree.configure(yscrollcommand=m_vsb.set, xscrollcommand=m_hsb.set)
            m_tree.grid(row=0, column=0, sticky="nsew")
            m_vsb.grid(row=0, column=1, sticky="ns")
            m_hsb.grid(row=1, column=0, sticky="ew")
            m_frame.grid_rowconfigure(0, weight=1)
            m_frame.grid_columnconfigure(0, weight=1)

            for lit, raw, sim in matches:
                m_tree.insert("", "end", values=(
                    lit.get("title", ""),
                    raw.get("title", ""),
                    f"{sim:.0%}" if sim == 1.0 else f"{sim:.0%}",
                    lit.get("authors", "") or "",
                    lit.get("pub_date", "") or "",
                ))

            m_detail = tk.Text(m_win, wrap="word", font=("", 10), height=6, state="disabled")
            m_detail.pack(fill="x", padx=15, pady=5)

            def _m_select(event):
                sel = m_tree.selection()
                if not sel:
                    return
                idx = m_tree.index(sel[0])
                if idx < len(matches):
                    lit, raw, sim = matches[idx]
                    m_detail.configure(state="normal")
                    m_detail.delete("1.0", "end")
                    m_detail.insert("end", f"文献库标题：{lit.get('title', '')}\n")
                    m_detail.insert("end", f"原始结果标题：{raw.get('title', '')}\n")
                    m_detail.insert("end", f"相似度：{sim:.0%}\n" if sim == 1.0 else f"相似度：{sim:.0%}\n")
                    m_detail.insert("end", f"作者：{lit.get('authors', '') or ''}\n")
                    m_detail.insert("end", f"发表时间：{lit.get('pub_date', '') or ''}\n")
                    m_detail.insert("end", f"原始结果作者：{raw.get('authors', '') or ''}\n")
                    m_detail.insert("end", f"原始结果年份：{raw.get('year', '') or ''}\n")
                    m_detail.insert("end", f"原始结果链接：{raw.get('link', '') or ''}\n")
                    m_detail.configure(state="disabled")

            m_tree.bind("<<TreeviewSelect>>", _m_select)

            ttk.Button(m_win, text="关闭", command=m_win.destroy).pack(pady=10)

        ttk.Button(bottom, text="还原", command=_restore_selected).pack(side="left", padx=2)
        ttk.Button(bottom, text="匹配文献库", command=_match_lit).pack(side="left", padx=2)
        ttk.Button(bottom, text="导出Excel", command=_export_excel).pack(side="left", padx=2)
        ttk.Label(bottom, text=f"共 {total_results} 条原始结果（{len(queries)} 次搜索）").pack(side="left", padx=10)
        ttk.Button(bottom, text="关闭", command=dlg.destroy).pack(side="right")

        # Auto-select first query
        if queries:
            query_tree.selection_set(query_tree.get_children()[0])
            on_select_query(None)

    # ---- Filtering ----

    def _update_filter_queries(self):
        current = self._filter_query_var.get()
        queries = ["全部"] + sorted(self._query_groups.keys())
        self._filter_query_combo["values"] = queries
        if current in queries:
            self._filter_query_var.set(current)
        else:
            self._filter_query_var.set("全部")

    def _apply_filter(self):
        query = self._filter_query_var.get()
        for item in self._tree.get_children(""):
            if query == "全部" or self._tree.item(item, "values")[0] == query or \
               query in self._tree.item(item, "text", ""):
                self._tree.reattach(item, "", "end")
            else:
                self._tree.detach(item)

    def _clear_filter(self):
        self._filter_query_var.set("全部")
        self._apply_filter()

    # ---- Sorting ----

    def _sort_by(self, col):
        reverse = getattr(self, "_sort_reverse", False)
        self._sort_reverse = not reverse

        col_map = {"query": 0, "title": 1, "reason": 2}
        # Update column headers
        col_labels = {"query": "搜索词", "title": "标题", "reason": "筛选理由"}
        for c in ("query", "title", "reason"):
            base = col_labels[c]
            if c == col:
                self._tree.heading(c, text=base + (" ▼" if reverse else " ▲"))
            else:
                self._tree.heading(c, text=base)

        # Sort leaf items within each parent group
        for parent_iid in self._tree.get_children(""):
            if not parent_iid.startswith("G_"):
                continue
            children = self._tree.get_children(parent_iid)
            items = []
            for iid in children:
                val = self._tree.set(iid, col)
                items.append((val.lower() if val else "", iid))
            items.sort(key=lambda x: x[0], reverse=reverse)
            for i, (_, iid) in enumerate(items):
                self._tree.move(iid, parent_iid, i)

    # ---- Bottom count ----

    def _update_bottom_count(self):
        self._bottom_count_var.set(f"共 {len(self._results)} 条结果")

    # ---- Keyword gen ----

    def _open_keyword_gen(self):
        from ui.keyword_gen import KeywordGenDialog
        def on_keywords(keywords):
            if keywords:
                self._keyword_var.set(keywords[0])
                for kw in keywords:
                    if kw not in self._queue_listbox.get(0, "end"):
                        self._queue_listbox.insert("end", kw)
        KeywordGenDialog(self.win, callback=on_keywords)

    # ---- Log viewer ----

    def _open_log_viewer(self):
        if hasattr(self, "_status_dlg") and self._status_dlg is not None:
            try:
                self._status_dlg.win.deiconify()
                self._status_dlg.win.lift()
                return
            except tk.TclError:
                self._status_dlg = None

        import os
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
        log_files = sorted(
            [f for f in os.listdir(log_dir) if f.endswith(".log")],
            reverse=True
        ) if os.path.isdir(log_dir) else []

        dlg = tk.Toplevel(self.win)
        dlg.title("队列日志")
        dlg.geometry("650x500")
        dlg.transient(self.win)

        sel_frame = ttk.Frame(dlg, padding=5)
        sel_frame.pack(fill="x")
        ttk.Label(sel_frame, text="日志文件：").pack(side="left")
        file_var = tk.StringVar()
        file_combo = ttk.Combobox(sel_frame, textvariable=file_var,
                                   values=log_files, state="readonly", width=50)
        file_combo.pack(side="left", padx=5)
        if log_files:
            file_combo.current(0)

        def _refresh():
            nonlocal log_files
            log_files = sorted(
                [f for f in os.listdir(log_dir) if f.endswith(".log")],
                reverse=True
            ) if os.path.isdir(log_dir) else []
            file_combo["values"] = log_files
            if log_files:
                file_combo.current(0)
                _load()

        def _load():
            fname = file_var.get()
            if not fname:
                return
            fpath = os.path.join(log_dir, fname)
            log_text.configure(state="normal")
            log_text.delete("1.0", "end")
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    content = f.read()
                log_text.insert("1.0", content)
            except Exception as e:
                log_text.insert("1.0", f"读取日志失败: {e}")
            log_text.configure(state="disabled")
            log_text.see("end")

        file_combo.bind("<<ComboboxSelected>>", lambda e: _load())
        ttk.Button(sel_frame, text="刷新", command=_refresh).pack(side="left", padx=2)
        ttk.Button(sel_frame, text="补跑失败词", command=lambda: self._retry_from_log(file_var.get())).pack(side="left", padx=2)

        log_frame = ttk.Frame(dlg, padding=5)
        log_frame.pack(fill="both", expand=True)
        log_text = tk.Text(log_frame, wrap="word", font=("", 9))
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
        log_text.configure(yscrollcommand=log_scroll.set)
        log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        ttk.Button(dlg, text="关闭", command=dlg.destroy).pack(pady=5)

        if log_files:
            _load()
        else:
            log_text.configure(state="normal")
            log_text.insert("1.0", "暂无日志文件。\n\n队列搜索的日志会自动保存到 data/logs/ 目录。")
            log_text.configure(state="disabled")

    def _retry_from_log(self, log_filename):
        """Parse failed/incomplete keywords from log file and add them back to queue."""
        if not log_filename:
            return
        import os, re as _re
        log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "logs")
        fpath = os.path.join(log_dir, log_filename)
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception:
            messagebox.showerror("错误", f"无法读取日志文件: {log_filename}")
            return
        # Collect completed keywords to exclude them
        completed = set()
        failed = []
        for line in content.split("\n"):
            if "[完成]" in line:
                m = _re.search(r'\[完成\] (.+?) —', line)
                if m:
                    completed.add(m.group(1).strip())
            elif "[失败]" in line:
                m = _re.search(r'\[失败\] (.+?) —', line)
                if m:
                    kw = m.group(1).strip()
                    if kw not in failed:
                        failed.append(kw)
            elif "[无结果]" in line:
                m = _re.search(r'\[无结果\] (.+?) —', line)
                if m:
                    kw = m.group(1).strip()
                    if kw not in failed and kw not in completed:
                        failed.append(kw)
            elif "[未处理]" in line:
                m = _re.search(r'\[未处理\] (.+)$', line)
                if m:
                    kw = m.group(1).strip()
                    if kw not in failed and kw not in completed:
                        failed.append(kw)
            # Old format compatibility
            elif "[错误]" in line or "[跳过]" in line:
                m = _re.search(r'\] \[(?:Scholar|Serper|百度)\] (.+?):', line)
                if m:
                    kw = m.group(1).strip()
                    if kw not in failed and kw not in completed:
                        failed.append(kw)
        if not failed:
            messagebox.showinfo("提示", "该日志中没有找到失败或未完成的关键词。")
            return
        existing = set(self._queue_listbox.get(0, "end"))
        added = 0
        for kw in failed:
            if kw not in existing:
                self._queue_listbox.insert("end", kw)
                existing.add(kw)
                added += 1
        messagebox.showinfo("补跑完成", f"从日志中提取 {len(failed)} 个失败词，成功添加 {added} 个到队列。")

    # ---- Queue status dialog ----

    def _open_queue_status(self, group=None):
        if hasattr(self, "_status_dlg") and self._status_dlg is not None:
            try:
                self._status_dlg.win.winfo_exists()
                self._status_dlg.win.lift()
                return
            except tk.TclError:
                self._status_dlg = None
        self._status_dlg = QueueStatusDialog(self.win, self, group=group, log_file=self._log_file)

    def _match_literature(self, force=False):
        """Cross-reference literature library titles against current search results."""
        if not self._results:
            messagebox.showinfo("提示", "当前没有搜索结果。请先执行搜索。")
            return

        if not force and db.has_lit_matches("main"):
            self._show_saved_matches("main")
            return

        import difflib, threading

        lit_records = db.get_all_literature()
        if not lit_records:
            messagebox.showinfo("提示", "文献库中没有文献。")
            return

        search_titles = set()
        for r in self._results:
            t = (r.get("title") or "").strip()
            if t:
                search_titles.add(t)

        search_list = list(search_titles)

        # Progress popup
        popup = tk.Toplevel(self.win)
        popup.title("匹配中...")
        popup.geometry("350x100")
        popup.transient(self.win)
        popup.grab_set()
        ttk.Label(popup, text=f"正在匹配 {len(lit_records)} 条文献 × {len(search_list)} 条搜索结果...",
                  font=("", 10)).pack(padx=20, pady=(15, 5))
        prog = ttk.Progressbar(popup, mode="indeterminate")
        prog.pack(padx=20, fill="x", pady=(0, 15))
        prog.start()

        result_holder = [None]
        done_flag = [False]

        def _run():
            try:
                match_list = []
                total = len(lit_records)
                for j, lit in enumerate(lit_records):
                    lit_title = (lit.get("title") or "").strip()
                    if not lit_title:
                        if j % 10 == 0:
                            result_holder[0] = ("progress", j + 1, total)
                        continue
                    lit_lower = lit_title.lower()
                    for st in search_list:
                        st_lower = st.lower()
                        if lit_lower == st_lower:
                            match_list.append((lit, st, 1.0))
                        elif lit_lower in st_lower or st_lower in lit_lower:
                            match_list.append((lit, st, 0.85))
                        else:
                            ratio = difflib.SequenceMatcher(None, lit_lower, st_lower).ratio()
                            if ratio > 0.5:
                                match_list.append((lit, st, round(ratio, 2)))
                    if j % 10 == 0:
                        result_holder[0] = ("progress", j + 1, total)
                match_list.sort(key=lambda x: x[2], reverse=True)
                result_holder[0] = ("ok", match_list, lit_records, search_list)
            except Exception as e:
                result_holder[0] = ("err", str(e))
            done_flag[0] = True

        threading.Thread(target=_run, daemon=True).start()

        def _poll():
            if done_flag[0]:
                popup.destroy()
                status = result_holder[0]
                if status[0] == "err":
                    messagebox.showerror("匹配出错", f"匹配过程出错：{status[1]}")
                    return
                _, match_list, lit_recs, s_list = status
                db.save_lit_matches("main", match_list)
                if not match_list:
                    messagebox.showinfo("匹配结果",
                        f"未找到匹配。\n文献库共 {len(lit_recs)} 条，搜索结果共 {len(s_list)} 条。")
                    return
                self._show_lit_match_result(match_list, lit_recs, s_list)
            elif result_holder[0] and result_holder[0][0] == "progress":
                _, cur, tot = result_holder[0]
                popup.title(f"匹配中... {cur}/{tot}")
                self.win.after(100, _poll)
            else:
                self.win.after(100, _poll)

        self.win.after(100, _poll)

    def _show_lit_match_result(self, matches, lit_records, search_titles):
        win = tk.Toplevel(self.win)
        win.title(f"匹配结果 - {len(matches)} 条")
        win.geometry("1000x550")
        win.transient(self.win)

        ttk.Label(win, text=f"文献库中共 {len(lit_records)} 条，搜索结果共 {len(search_titles)} 条，匹配 {len(matches)} 对：",
                  font=("", 10)).pack(padx=15, pady=(15, 5))

        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

        cols = ("lit_title", "search_title", "similarity", "authors", "pub_date")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        tree.heading("lit_title", text="文献库标题")
        tree.heading("search_title", text="搜索结果标题")
        tree.heading("similarity", text="相似度")
        tree.heading("authors", text="作者")
        tree.heading("pub_date", text="发表时间")
        tree.column("lit_title", width=300)
        tree.column("search_title", width=300)
        tree.column("similarity", width=60)
        tree.column("authors", width=130)
        tree.column("pub_date", width=80)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        for lit, st, sim in matches:
            tree.insert("", "end", values=(
                lit.get("title", ""),
                st,
                f"{sim:.0%}" if sim == 1.0 else f"{sim:.0%}",
                lit.get("authors", "") or "",
                lit.get("pub_date", "") or "",
            ))

        def _on_select(event):
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            if idx < len(matches):
                lit, st, sim = matches[idx]
                detail_text.config(state="normal")
                detail_text.delete("1.0", "end")
                detail_text.insert("end", f"文献库标题：{lit.get('title', '')}\n")
                detail_text.insert("end", f"搜索结果标题：{st}\n")
                detail_text.insert("end", f"相似度：{sim:.0%}\n" if sim == 1.0 else f"相似度：{sim:.0%}\n")
                detail_text.insert("end", f"作者：{lit.get('authors', '') or ''}\n")
                detail_text.insert("end", f"发表时间：{lit.get('pub_date', '') or ''}\n")
                detail_text.insert("end", f"期刊：{lit.get('journal', '') or ''}\n")
                detail_text.insert("end", f"出版社：{lit.get('publisher', '') or ''}\n")
                detail_text.insert("end", f"链接：{lit.get('source_url', '') or ''}\n")
                detail_text.configure(state="disabled")

        tree.bind("<<TreeviewSelect>>", _on_select)

        detail_text = tk.Text(win, wrap="word", font=("", 10), height=6, state="disabled")
        detail_text.pack(fill="x", padx=15, pady=5)

        btn_row = ttk.Frame(win)
        btn_row.pack(fill="x", padx=15, pady=10)
        ttk.Button(btn_row, text="关闭", command=win.destroy).pack(side="right", padx=2)

    def _show_saved_matches(self, match_type):
        batches = db.get_lit_match_batches(match_type)
        if not batches:
            messagebox.showinfo("提示", "没有已保存的匹配记录。")
            return

        win = tk.Toplevel(self.win)
        win.title(f"匹配记录 - {len(batches)} 次")
        win.geometry("500x400")
        win.transient(self.win)

        label_text = "搜索结果匹配文献库" if match_type == "main" else "原始结果匹配文献库"
        ttk.Label(win, text=f"{label_text} — 共 {len(batches)} 次匹配记录，双击查看：",
                  font=("", 10)).pack(padx=15, pady=(15, 5))

        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

        cols = ("time", "count")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        tree.heading("time", text="匹配时间")
        tree.heading("count", text="匹配数量")
        tree.column("time", width=250)
        tree.column("count", width=100)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        for b in batches:
            tree.insert("", "end", values=(b["batch_id"], f"{b['cnt']} 对"))

        def _open_batch(event):
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            batch_id = batches[idx]["batch_id"]
            saved = db.get_lit_matches(match_type, batch_id=batch_id)
            if saved:
                self._show_batch_results(win, batch_id, saved, match_type)

        tree.bind("<Double-1>", _open_batch)

        btn_row = ttk.Frame(win)
        btn_row.pack(fill="x", padx=15, pady=10)
        if match_type == "main":
            ttk.Button(btn_row, text="重新匹配",
                       command=lambda: [win.destroy(), db.clear_lit_matches("main"), self._match_literature(force=True)]
                       ).pack(side="left", padx=2)
        ttk.Button(btn_row, text="关闭", command=win.destroy).pack(side="right", padx=2)

    def _show_batch_results(self, parent_win, batch_id, saved, match_type):
        win = tk.Toplevel(parent_win)
        win.title(f"匹配结果 - {batch_id}")
        win.geometry("1000x550")
        win.transient(parent_win)

        ttk.Label(win, text=f"匹配时间：{batch_id}，共 {len(saved)} 对：",
                  font=("", 10)).pack(padx=15, pady=(15, 5))

        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

        cols = ("lit_title", "matched_title", "similarity", "authors", "pub_date")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        tree.heading("lit_title", text="文献库标题")
        tree.heading("matched_title", text="搜索结果标题")
        tree.heading("similarity", text="相似度")
        tree.heading("authors", text="作者")
        tree.heading("pub_date", text="发表时间")
        tree.column("lit_title", width=300)
        tree.column("matched_title", width=300)
        tree.column("similarity", width=60)
        tree.column("authors", width=130)
        tree.column("pub_date", width=80)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        for m in saved:
            sim = m["similarity"]
            tree.insert("", "end", values=(
                m["lit_title"],
                m["matched_title"],
                f"{sim:.0%}" if sim == 1.0 else f"{sim:.0%}",
                m["lit_authors"] or "",
                m["lit_pub_date"] or "",
            ))

        detail_text = tk.Text(win, wrap="word", font=("", 10), height=6, state="disabled")
        detail_text.pack(fill="x", padx=15, pady=5)

        def _on_select(event):
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            if idx < len(saved):
                m = saved[idx]
                sim = m["similarity"]
                detail_text.config(state="normal")
                detail_text.delete("1.0", "end")
                detail_text.insert("end", f"文献库标题：{m['lit_title']}\n")
                detail_text.insert("end", f"搜索结果标题：{m['matched_title']}\n")
                detail_text.insert("end", f"相似度：{sim:.0%}\n" if sim == 1.0 else f"相似度：{sim:.0%}\n")
                detail_text.insert("end", f"作者：{m['lit_authors'] or ''}\n")
                detail_text.insert("end", f"发表时间：{m['lit_pub_date'] or ''}\n")
                detail_text.configure(state="disabled")

        tree.bind("<<TreeviewSelect>>", _on_select)

        btn_row = ttk.Frame(win)
        btn_row.pack(fill="x", padx=15, pady=10)
        ttk.Button(btn_row, text="关闭", command=win.destroy).pack(side="right", padx=2)


class LiteratureLibraryDialog:
    def __init__(self, parent, api_key="", search_tab=None):
        self._api_key = api_key
        self._search_tab = search_tab
        self._current_id = None
        self.win = tk.Toplevel(parent)
        self.win.title("文献库")
        self.win.geometry("900x700")
        self.win.transient(parent)
        self._build_ui()
        self._refresh()

    def _build_ui(self):
        toolbar = ttk.Frame(self.win, padding=5)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="手动添加", command=self._manual_add).pack(side="left", padx=2)
        ttk.Button(toolbar, text="上传文件", command=self._upload_file).pack(side="left", padx=2)
        ttk.Button(toolbar, text="刷新", command=self._refresh).pack(side="left", padx=2)
        ttk.Button(toolbar, text="删除选中", command=self._delete_selected).pack(side="left", padx=2)
        ttk.Button(toolbar, text="输入文献信息", command=self._open_bib_input_dialog).pack(side="left", padx=2)
        ttk.Label(toolbar, text="搜索：").pack(side="left", padx=(15, 0))
        self._search_var = tk.StringVar()
        search_entry = ttk.Entry(toolbar, textvariable=self._search_var, width=18)
        search_entry.pack(side="left", padx=3)
        search_entry.bind("<Return>", lambda e: self._refresh())
        ttk.Button(toolbar, text="搜索", command=self._refresh).pack(side="left", padx=2)

        vpaned = ttk.PanedWindow(self.win, orient="vertical")
        vpaned.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        top_paned = ttk.PanedWindow(vpaned, orient="horizontal")
        vpaned.add(top_paned, weight=1)

        left_frame = ttk.Frame(top_paned)
        top_paned.add(left_frame, weight=1)
        columns = ("title", "authors", "pub_date", "source_engine", "created_at")
        self._tree = ttk.Treeview(left_frame, columns=columns,
                                   show="headings", selectmode="extended")
        self._tree.heading("title", text="标题")
        self._tree.heading("authors", text="作者")
        self._tree.heading("pub_date", text="发表日期")
        self._tree.heading("source_engine", text="来源引擎")
        self._tree.heading("created_at", text="收录时间")
        self._tree.column("title", width=200)
        self._tree.column("authors", width=90)
        self._tree.column("pub_date", width=90)
        self._tree.column("source_engine", width=80)
        self._tree.column("created_at", width=110)
        lit_scroll = ttk.Scrollbar(left_frame, orient="vertical",
                                    command=self._tree.yview)
        self._tree.configure(yscrollcommand=lit_scroll.set)
        self._tree.pack(side="left", fill="both", expand=True)
        lit_scroll.pack(side="right", fill="y")
        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        self._tree.bind("<Double-1>", self._open_edit_dialog)

        right_frame = ttk.Frame(top_paned)
        top_paned.add(right_frame, weight=1)
        row1 = ttk.Frame(right_frame)
        row1.pack(fill="x", pady=1)
        ttk.Label(row1, text="标题：", width=5).pack(side="left")
        self._title_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self._title_var, width=35).pack(side="left")
        row2 = ttk.Frame(right_frame)
        row2.pack(fill="x", pady=1)
        ttk.Label(row2, text="作者：", width=5).pack(side="left")
        self._authors_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self._authors_var, width=20).pack(side="left")
        ttk.Label(row2, text="发表日期：").pack(side="left", padx=(5, 0))
        self._pub_date_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self._pub_date_var, width=12).pack(side="left")
        row3 = ttk.Frame(right_frame)
        row3.pack(fill="x", pady=1)
        ttk.Label(row3, text="URL：", width=5).pack(side="left")
        self._url_var = tk.StringVar()
        url_entry = ttk.Entry(row3, textvariable=self._url_var, font=("", 8))
        url_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row3, text="下载", command=self._download_url).pack(side="left", padx=2)
        row3b = ttk.Frame(right_frame)
        row3b.pack(fill="x", pady=1)
        ttk.Label(row3b, text="来源：", width=5).pack(side="left")
        self._source_engine_var = tk.StringVar()
        ttk.Entry(row3b, textvariable=self._source_engine_var, width=15).pack(side="left")
        row3c = ttk.Frame(right_frame)
        row3c.pack(fill="x", pady=1)
        ttk.Label(row3c, text="文件：", width=5).pack(side="left")
        self._file_path_var = tk.StringVar()
        file_entry = ttk.Entry(row3c, textvariable=self._file_path_var, font=("", 8), state="readonly")
        file_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row3c, text="上传", command=self._attach_file).pack(side="left", padx=1)
        ttk.Button(row3c, text="打开", command=self._open_file).pack(side="left")
        ttk.Label(right_frame, text="摘要：", font=("", 9)).pack(anchor="w", pady=(5, 0))
        self._summary_text = tk.Text(right_frame, wrap="word", font=("", 9), height=2)
        self._summary_text.pack(fill="x")
        ttk.Label(right_frame, text="备注：", font=("", 9)).pack(anchor="w")
        self._notes_text = tk.Text(right_frame, wrap="word", font=("", 9), height=2)
        self._notes_text.pack(fill="x")

        btn_bar = ttk.Frame(right_frame)
        btn_bar.pack(fill="x", pady=(5, 0))
        ttk.Button(btn_bar, text="发送到AI提取", command=self._send_to_ai_extract).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="保存修改", command=self._save_edits).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="在浏览器中打开", command=self._open_url).pack(side="left", padx=2)

        bottom_frame = ttk.Frame(vpaned)
        vpaned.add(bottom_frame, weight=1)
        ttk.Label(bottom_frame, text="正文内容：", font=("", 9)).pack(anchor="w")
        self._content_text = tk.Text(bottom_frame, wrap="word", font=("", 10), height=8)
        content_scroll = ttk.Scrollbar(bottom_frame, orient="vertical",
                                        command=self._content_text.yview)
        self._content_text.configure(yscrollcommand=content_scroll.set)
        self._content_text.pack(side="left", fill="both", expand=True)
        content_scroll.pack(side="right", fill="y")

    def _refresh(self):
        for item in self._tree.get_children():
            self._tree.delete(item)
        search = self._search_var.get().strip() or None
        rows = db.get_all_literature(search=search)
        for r in rows:
            self._tree.insert("", "end", iid=str(r["id"]), values=(
                r.get("title", ""),
                r.get("authors", "") or "",
                r.get("pub_date", "") or "",
                r.get("source_engine", "") or "",
                r.get("created_at", "") or "",
            ))

    def _on_select(self, event):
        sel = self._tree.selection()
        if not sel:
            return
        lit_id = int(sel[0])
        row = db.get_literature_by_id(lit_id)
        if not row:
            return
        self._current_id = lit_id
        self._title_var.set(row.get("title", ""))
        self._authors_var.set(row.get("authors", "") or "")
        self._pub_date_var.set(row.get("pub_date", "") or "")
        self._url_var.set(row.get("source_url", "") or "")
        self._source_engine_var.set(row.get("source_engine", "") or "")
        self._file_path_var.set(row.get("file_path", "") or "")
        self._summary_text.delete("1.0", "end")
        self._summary_text.insert("1.0", row.get("summary", "") or "")
        self._notes_text.delete("1.0", "end")
        self._notes_text.insert("1.0", row.get("notes", "") or "")
        self._content_text.delete("1.0", "end")
        self._content_text.insert("1.0", row.get("content", "") or "")

    def _manual_add(self):
        """Open a dialog for manually adding a literature entry."""
        dlg = tk.Toplevel(self.win)
        dlg.title("手动添加文献")
        dlg.geometry("600x520")
        dlg.transient(self.win)

        f = ttk.Frame(dlg, padding=10)
        f.pack(fill="both", expand=True)

        # Title
        ttk.Label(f, text="标题 *").pack(anchor="w")
        title_var = tk.StringVar()
        ttk.Entry(f, textvariable=title_var, font=("", 10)).pack(fill="x", pady=(0, 5))

        # Authors + pub_date
        r1 = ttk.Frame(f)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="作者：").pack(side="left")
        authors_var = tk.StringVar()
        ttk.Entry(r1, textvariable=authors_var, width=25).pack(side="left", padx=(0, 10))
        ttk.Label(r1, text="发表日期：").pack(side="left")
        pub_date_var = tk.StringVar()
        ttk.Entry(r1, textvariable=pub_date_var, width=12).pack(side="left")

        # Source URL
        ttk.Label(f, text="来源链接").pack(anchor="w", pady=(5, 0))
        url_var = tk.StringVar()
        ttk.Entry(f, textvariable=url_var, font=("", 9)).pack(fill="x", pady=(0, 5))

        # Source engine
        ttk.Label(f, text="来源引擎").pack(anchor="w")
        engine_var = tk.StringVar(value="手动录入")
        engines = ["手动录入", "浏览器扩展", "Google Scholar", "Serper搜索", "百度搜索", "PDF上传", "由灰色文献转入"]
        engine_combo = ttk.Combobox(f, textvariable=engine_var, values=engines, width=20)
        engine_combo.pack(anchor="w", pady=(0, 5))

        # File attachment
        ttk.Label(f, text="附件文件").pack(anchor="w")
        file_row = ttk.Frame(f)
        file_row.pack(fill="x", pady=(0, 5))
        file_var = tk.StringVar()
        file_entry = ttk.Entry(file_row, textvariable=file_var, font=("", 8), state="readonly")
        file_entry.pack(side="left", fill="x", expand=True)

        def _pick_file():
            path = filedialog.askopenfilename(filetypes=[
                ("All supported", "*.pdf *.docx *.doc *.txt *.md *.csv *.xlsx *.pptx"),
                ("PDF files", "*.pdf"),
                ("Word documents", "*.docx *.doc"),
                ("Text files", "*.txt *.md *.csv"),
                ("All files", "*.*"),
            ])
            if not path:
                return
            import shutil
            import uuid
            ext = os.path.splitext(path)[1].lower()
            lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
            os.makedirs(lit_dir, exist_ok=True)
            dest_name = f"{uuid.uuid4().hex}{ext}"
            shutil.copy2(path, os.path.join(lit_dir, dest_name))
            file_var.set(dest_name)

        ttk.Button(file_row, text="选择文件", command=_pick_file).pack(side="left", padx=3)

        # Summary
        ttk.Label(f, text="摘要").pack(anchor="w")
        summary_text = tk.Text(f, wrap="word", font=("", 9), height=3)
        summary_text.pack(fill="x", pady=(0, 5))

        # Notes
        ttk.Label(f, text="备注").pack(anchor="w")
        notes_text = tk.Text(f, wrap="word", font=("", 9), height=2)
        notes_text.pack(fill="x", pady=(0, 5))

        def _do_add():
            title = title_var.get().strip()
            if not title:
                messagebox.showwarning("提示", "标题不能为空。", parent=dlg)
                return
            lid = db.insert_literature({
                "title": title,
                "authors": authors_var.get().strip(),
                "pub_date": pub_date_var.get().strip(),
                "source_url": url_var.get().strip(),
                "source_engine": engine_var.get().strip(),
                "summary": summary_text.get("1.0", "end-1c").strip(),
                "notes": notes_text.get("1.0", "end-1c").strip(),
                "file_path": file_var.get().strip(),
            })
            if lid:
                self._refresh()
                dlg.destroy()
                messagebox.showinfo("添加完成", f"文献「{title}」已添加到文献库。", parent=self.win)
            else:
                messagebox.showerror("错误", "添加失败。", parent=dlg)

        btn_bar = ttk.Frame(f)
        btn_bar.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_bar, text="添加", command=_do_add).pack(side="right", padx=3)
        ttk.Button(btn_bar, text="取消", command=dlg.destroy).pack(side="right", padx=3)

    def _upload_file(self):
        """Upload a file as a new literature entry. Supports PDF, DOCX, TXT, MD, CSV, etc."""
        path = filedialog.askopenfilename(filetypes=[
            ("All supported", "*.pdf *.docx *.doc *.txt *.md *.csv *.xlsx *.pptx"),
            ("PDF files", "*.pdf"),
            ("Word documents", "*.docx *.doc"),
            ("Text files", "*.txt *.md *.csv"),
            ("All files", "*.*"),
        ])
        if not path:
            return
        filename = os.path.basename(path)
        ext = os.path.splitext(path)[1].lower()
        text = ""
        title = filename
        authors = ""
        pub_date = ""
        source_engine = "文件上传"

        if ext == ".pdf":
            try:
                from utils.pdf_extractor import extract_text_and_metadata
                text, metadata = extract_text_and_metadata(path)
                title = metadata.get("title") or metadata.get("file_name", filename).replace(".pdf", "")
                authors = metadata.get("author", "")
                pub_date = metadata.get("pub_date", "")
                source_engine = "PDF上传"
            except Exception as e:
                messagebox.showerror("PDF提取失败", f"无法读取PDF文件：{e}")
                return
        elif ext in (".txt", ".md", ".csv"):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    text = f.read()
                title = filename
            except Exception as e:
                messagebox.showerror("文件读取失败", f"无法读取文件：{e}")
                return
        elif ext in (".docx",):
            try:
                from docx import Document
                doc = Document(path)
                text = "\n".join(p.text for p in doc.paragraphs)
                title = filename
            except Exception as e:
                messagebox.showwarning("文本提取失败", f"无法从docx提取文本，仅保存文件。\n{e}")
        else:
            messagebox.showinfo("提示", f"不支持对 {ext} 格式进行文本提取，文件将仅作为附件保存。")

        summary = ""
        if self._api_key and text.strip():
            try:
                from utils.ai_extractor import generate_summary
                summary = generate_summary(text, self._api_key) or ""
            except Exception:
                pass

        import shutil
        import uuid
        lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
        os.makedirs(lit_dir, exist_ok=True)
        dest_name = f"{uuid.uuid4().hex}{ext}"
        shutil.copy2(path, os.path.join(lit_dir, dest_name))
        lid = db.insert_literature({
            "title": title, "authors": authors,
            "pub_date": pub_date,
            "source_engine": source_engine,
            "content": text, "summary": summary,
            "file_path": dest_name,
        })
        if lid:
            self._refresh()
            messagebox.showinfo("导入完成", f"文件已导入文献库。\n标题：{title}")
        else:
            messagebox.showerror("导入失败", "无法保存到文献库。")

    def _attach_file(self):
        """Attach a file to the currently selected literature entry."""
        if not self._current_id:
            messagebox.showinfo("提示", "请先在左侧列表中选择一条文献。")
            return
        path = filedialog.askopenfilename(filetypes=[
            ("All supported", "*.pdf *.docx *.doc *.txt *.md *.csv *.xlsx *.pptx"),
            ("PDF files", "*.pdf"),
            ("Word documents", "*.docx *.doc"),
            ("Text files", "*.txt *.md *.csv"),
            ("All files", "*.*"),
        ])
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        text = ""
        if ext == ".pdf":
            try:
                from utils.pdf_extractor import extract_text_and_metadata
                text, metadata = extract_text_and_metadata(path)
            except Exception as e:
                messagebox.showerror("PDF提取失败", f"无法读取PDF文件：{e}")
                return
        elif ext in (".txt", ".md", ".csv"):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    text = f.read()
            except Exception as e:
                messagebox.showerror("文件读取失败", f"无法读取文件：{e}")
                return
        elif ext in (".docx",):
            try:
                from docx import Document
                doc = Document(path)
                text = "\n".join(p.text for p in doc.paragraphs)
            except Exception as e:
                messagebox.showwarning("文本提取失败", f"无法从docx提取文本，仅保存文件。\n{e}")

        import shutil
        import uuid
        lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
        os.makedirs(lit_dir, exist_ok=True)
        dest_name = f"{uuid.uuid4().hex}{ext}"
        shutil.copy2(path, os.path.join(lit_dir, dest_name))

        update_data = {"file_path": dest_name}
        if text.strip():
            update_data["content"] = text
        db.update_literature(self._current_id, update_data)
        self._file_path_var.set(dest_name)
        self._refresh()
        messagebox.showinfo("完成", f"文件已附加到当前文献。")

    def _open_file(self):
        """Open the literature file storage folder, selecting the attached file."""
        lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
        os.makedirs(lit_dir, exist_ok=True)
        file_path = self._file_path_var.get().strip()
        full_path = os.path.join(lit_dir, file_path) if file_path else ""
        import subprocess
        import platform
        try:
            if platform.system() == "Darwin":
                if full_path and os.path.exists(full_path):
                    subprocess.run(["open", "-R", full_path])
                else:
                    subprocess.run(["open", lit_dir])
            elif platform.system() == "Windows":
                if full_path and os.path.exists(full_path):
                    subprocess.run(["explorer", "/select,", full_path])
                else:
                    os.startfile(lit_dir)
            else:
                if full_path and os.path.exists(full_path):
                    subprocess.run(["xdg-open", os.path.dirname(full_path)])
                else:
                    subprocess.run(["xdg-open", lit_dir])
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def _download_url(self):
        """Download file from the URL field and attach to current literature.
        Uses requests to follow redirects and detect real Content-Type."""
        if not self._current_id:
            messagebox.showinfo("提示", "请先在左侧列表中选择一条文献。")
            return
        url = self._url_var.get().strip()
        if not url:
            messagebox.showinfo("提示", "URL字段为空，请先填入文件下载链接。")
            return
        import requests as req_lib
        import shutil
        import uuid
        import threading

        # Guess filename from URL
        raw_name = url.rstrip("/").split("/")[-1] or "downloaded_file"
        if "?" in raw_name:
            raw_name = raw_name.split("?")[0]

        # --- progress popup ---
        popup = tk.Toplevel(self.win)
        popup.title("下载中")
        popup.geometry("400x120")
        popup.transient(self.win)
        popup.resizable(False, False)
        f = ttk.Frame(popup, padding=15)
        f.pack(fill="both", expand=True)
        status_var = tk.StringVar(value=f"正在下载: {raw_name}")
        ttk.Label(f, textvariable=status_var, font=("", 10)).pack(pady=(0, 5))
        progress = ttk.Progressbar(f, mode="indeterminate")
        progress.pack(fill="x")
        progress.start(10)

        result_holder = [None]
        done_flag = [False]

        def _download():
            try:
                headers = {
                    "User-Agent": (
                        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/120.0.0.0 Safari/537.36"
                    ),
                    "Accept": "text/html,application/pdf,application/xhtml+xml,*/*",
                }
                resp = req_lib.get(url, headers=headers, allow_redirects=True,
                                   stream=True, timeout=60)

                # Determine extension from Content-Type
                ct = resp.headers.get("Content-Type", "").lower()
                ext_map = {
                    "application/pdf": ".pdf",
                    "application/msword": ".doc",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
                    "text/plain": ".txt",
                    "text/markdown": ".md",
                    "text/csv": ".csv",
                    "application/vnd.ms-excel": ".xls",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
                    "application/vnd.ms-powerpoint": ".ppt",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
                }
                ext = ".pdf"
                for mime, extension in ext_map.items():
                    if mime in ct:
                        ext = extension
                        break
                if ext == ".pdf" and os.path.splitext(raw_name)[1]:
                    ext = os.path.splitext(raw_name)[1].lower()

                filename = os.path.splitext(raw_name)[0] + ext if os.path.splitext(raw_name)[1] else raw_name + ext

                data = resp.content

                lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
                os.makedirs(lit_dir, exist_ok=True)
                dest_name = f"{uuid.uuid4().hex}{ext}"
                dest_path = os.path.join(lit_dir, dest_name)
                with open(dest_path, "wb") as fw:
                    fw.write(data)

                update_data = {"file_path": dest_name}
                if ext == ".pdf":
                    try:
                        from utils.pdf_extractor import extract_text_and_metadata
                        text, _ = extract_text_and_metadata(dest_path)
                        if text.strip():
                            update_data["content"] = text
                    except Exception:
                        pass
                elif ext in (".txt", ".md", ".csv"):
                    try:
                        text = data.decode("utf-8")
                        update_data["content"] = text
                    except Exception:
                        pass
                db.update_literature(self._current_id, update_data)
                result_holder[0] = ("ok", dest_name, filename, update_data.get("content", ""))
            except Exception as e:
                result_holder[0] = ("err", str(e))
            done_flag[0] = True

        threading.Thread(target=_download, daemon=True).start()

        def _poll():
            if done_flag[0]:
                popup.destroy()
                res = result_holder[0]
                if res and res[0] == "ok":
                    dest_name, filename, thread_content = res[1], res[2], res[3]
                    # Verify DB was updated; retry from main thread if not
                    row = db.get_literature_by_id(self._current_id)
                    if row and not row.get("file_path"):
                        update_data = {"file_path": dest_name}
                        if thread_content:
                            update_data["content"] = thread_content
                        db.update_literature(self._current_id, update_data)
                    self._refresh()
                    row = db.get_literature_by_id(self._current_id)
                    if row:
                        self._file_path_var.set(row.get("file_path", "") or "")
                        self._content_text.delete("1.0", "end")
                        self._content_text.insert("1.0", row.get("content", "") or "")
                    messagebox.showinfo("下载完成", f"「{filename}」已下载并附加到当前文献。", parent=self.win)
                else:
                    err_msg = res[1] if res else "未知错误"
                    messagebox.showerror("下载失败", err_msg, parent=self.win)
            else:
                self.win.after(200, _poll)

        self.win.after(200, _poll)

    def _open_detail_dialog(self):
        """Open the edit/detail dialog for the selected literature."""
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在左侧列表中选择一条文献。")
            return
        lit_id = int(sel[0])
        row = db.get_literature_by_id(lit_id)
        if row:
            LiteratureEditDialog(self.win, row, on_save=lambda: self._refresh())

    def _delete_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择要删除的文献。")
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(sel)} 条文献吗？"):
            for iid in sel:
                db.delete_literature(int(iid))
            self._refresh()
            self._current_id = None

    def _send_to_ai_extract(self):
        if not self._current_id:
            messagebox.showinfo("提示", "请先选择一条文献。")
            return
        from utils.perf_monitor import Tracker
        with Tracker("文献库-发送到AI提取"):
            self._send_to_ai_extract_impl()

    def _send_to_ai_extract_impl(self):
        row = db.get_literature_by_id(self._current_id)
        if not row:
            messagebox.showinfo("提示", "文献不存在。")
            return
        content = row.get("content", "") or ""
        if not content.strip():
            file_path = row.get("file_path", "") or ""
            if not file_path:
                messagebox.showinfo("提示", "该文献没有内容，也没有附加文件。")
                return
            lit_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "literature")
            full_path = os.path.join(lit_dir, file_path)
            if not os.path.exists(full_path):
                messagebox.showinfo("提示", f"附加文件不存在：{file_path}")
                return
            ext = os.path.splitext(file_path)[1].lower()
            try:
                if ext == ".pdf":
                    from utils.pdf_extractor import extract_text_and_metadata
                    content, _ = extract_text_and_metadata(full_path)
                elif ext in (".txt", ".md", ".csv"):
                    with open(full_path, "r", encoding="utf-8") as f:
                        content = f.read()
                elif ext in (".docx",):
                    from docx import Document
                    doc = Document(full_path)
                    content = "\n".join(p.text for p in doc.paragraphs)
                else:
                    messagebox.showinfo("提示", f"暂不支持 {ext} 格式的文本提取。")
                    return
            except Exception as e:
                messagebox.showerror("提取失败", f"无法从文件提取文本：{e}")
                return
            if not content.strip():
                messagebox.showinfo("提示", "文件中未提取到文字内容。")
                return
            db.update_literature(self._current_id, {"content": content})
            self._content_text.delete("1.0", "end")
            self._content_text.insert("1.0", content)
        if self._search_tab:
            title = row.get("title", "") or ""
            source_url = row.get("source_url", "") or ""
            if source_url:
                source_label = f"文献库: {title} | {source_url}"
            else:
                source_label = f"文献库: {title}"

            current = self._search_tab.ai_text_input.get("1.0", "end-1c").strip()
            if current:
                self._search_tab.ai_text_input.insert("end", f"\n\n--- [来源] {source_label} ---\n\n{content}")
            else:
                self._search_tab.ai_text_input.insert("end", f"--- [来源] {source_label} ---\n\n{content}")
            self._search_tab.ai_status_var.set(f"已从文献库加载：{title}")

    def _open_url(self, event=None):
        url = self._url_var.get().strip()
        if not url:
            messagebox.showinfo("提示", "没有可打开的链接。")
            return
        import subprocess
        subprocess.run(["open", "-a", "/Applications/Google Chrome.app", url])

    def _save_edits(self):
        if not self._current_id:
            return
        from utils.perf_monitor import Tracker
        with Tracker("文献库-保存修改"):
            self._save_edits_impl()

    def _save_edits_impl(self):
        data = {
            "title": self._title_var.get().strip(),
            "authors": self._authors_var.get().strip(),
            "pub_date": self._pub_date_var.get().strip(),
            "source_url": self._url_var.get().strip(),
            "source_engine": self._source_engine_var.get().strip(),
            "summary": self._summary_text.get("1.0", "end-1c").strip(),
            "notes": self._notes_text.get("1.0", "end-1c").strip(),
            "content": self._content_text.get("1.0", "end-1c").strip(),
        }
        db.update_literature(self._current_id, data)
        self._refresh()

    def _open_edit_dialog(self, event):
        sel = self._tree.selection()
        if not sel:
            return
        lit_id = int(sel[0])
        row = db.get_literature_by_id(lit_id)
        if row:
            LiteratureEditDialog(self.win, row, on_save=lambda: self._refresh())

    def _open_bib_input_dialog(self):
        win = tk.Toplevel(self.win)
        win.title("输入文献信息")
        win.geometry("1050x720")
        win.transient(self.win)

        ttk.Label(win, text="粘贴参考文献列表，AI将自动识别每条的作者、题目、期刊、出版社、城市、国家等信息：",
                  font=("", 10)).pack(padx=15, pady=(15, 2))

        text = tk.Text(win, wrap="word", font=("", 11), height=10)
        text.pack(fill="x", padx=15, pady=5)

        btn_row1 = ttk.Frame(win)
        btn_row1.pack(fill="x", padx=15, pady=(0, 5))
        parsed_entries = []

        RESULT_COLS = ("authors", "title", "pub_date", "journal", "publisher", "city", "country", "translator")
        RESULT_HEADINGS = {
            "authors": "作者", "title": "题目", "pub_date": "发表时间",
            "journal": "期刊名", "publisher": "出版社", "city": "城市",
            "country": "国家", "translator": "翻译人",
        }

        def _do_parse():
            input_text = text.get("1.0", "end-1c").strip()
            if not input_text:
                messagebox.showinfo("提示", "请先粘贴参考文献信息。")
                return

            import threading
            from utils.ai_extractor import _call_yunwu
            api_key = self._api_key or config.get_llm("search_parse")[2]

            prompt = (
                "你是一个参考文献解析器。请将以下参考文献列表解析为JSON数组。\n\n"
                "参考文献可能是任何语言（中文、英文、日文、韩文、法文、德文、俄文、西班牙文等），你都能识别和解析。\n\n"
                "每条文献提取以下字段（没有的字段用null）：\n"
                "- authors: 作者（保留原文姓名，多个用分号分隔）\n"
                "- title: 题目（保留原文语言，不要翻译）\n"
                "- pub_date: 发表时间（提取年份，如\"1836\"）\n"
                "- journal: 期刊名（保留原文）\n"
                "- publisher: 出版社（保留原文）\n"
                "- city: 城市\n"
                "- country: 国家\n"
                "- translator: 翻译人\n"
                "- source_url: 链接/DOI\n\n"
                "要求：\n"
                "1. 输出纯JSON数组，不要markdown代码块\n"
                "2. 每条参考文献对应数组中的一个对象\n"
                "3. 尽量提取所有可用信息，不要遗漏\n"
                "4. 所有字段保留原文语言，不要翻译成英文或中文\n"
                "5. 中文文献保留中文，日文文献保留日文，法文文献保留法文，以此类推\n\n"
                "参考文献列表：\n" + input_text
            )

            result_holder = [None]
            done_flag = [False]

            popup = tk.Toplevel(win)
            popup.title("解析中...")
            popup.geometry("300x80")
            popup.transient(win)
            ttk.Label(popup, text="AI 正在解析参考文献，请稍候...").pack(padx=20, pady=15)
            prog = ttk.Progressbar(popup, mode="indeterminate")
            prog.pack(padx=20, fill="x")
            prog.start()

            def _run():
                try:
                    result = _call_yunwu(prompt, api_key, temperature=0.1, max_tokens=8000, timeout=120)
                    result_holder[0] = ("ok", result)
                except Exception as e:
                    result_holder[0] = ("err", str(e))
                done_flag[0] = True

            threading.Thread(target=_run, daemon=True).start()

            def _poll():
                if done_flag[0]:
                    popup.destroy()
                    status, data = result_holder[0]
                    if status == "err":
                        messagebox.showerror("解析失败", f"AI解析出错：{data}")
                        return

                    if isinstance(data, list):
                        entries = data
                    elif isinstance(data, dict):
                        entries = [data]
                    else:
                        messagebox.showerror("解析失败", f"AI返回格式异常：{str(data)[:200]}")
                        return

                    parsed_entries.clear()
                    parsed_entries.extend(entries)

                    for item in result_tree.get_children():
                        result_tree.delete(item)
                    for entry in entries:
                        vals = [str(entry.get(c, "") or "") for c in RESULT_COLS]
                        result_tree.insert("", "end", values=vals)

                    messagebox.showinfo("解析完成", f"AI识别出 {len(entries)} 条文献，请核对后点击「确认入库」。",
                                       parent=win)
                else:
                    win.after(200, _poll)

            win.after(200, _poll)

        ttk.Button(btn_row1, text="AI 解析", command=_do_parse).pack(side="left", padx=2)

        ttk.Label(win, text="解析结果（双击单元格可编辑）：",
                  font=("", 10, "bold")).pack(padx=15, pady=(10, 2))

        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

        result_tree = ttk.Treeview(tree_frame, columns=RESULT_COLS, show="headings")
        for col in RESULT_COLS:
            result_tree.heading(col, text=RESULT_HEADINGS[col])
            result_tree.column(col, width=110, minwidth=60)
        result_tree.column("title", width=180)
        result_tree.column("authors", width=130)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=result_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=result_tree.xview)
        result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        result_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        def _on_double_click(event):
            region = result_tree.identify_region(event.x, event.y)
            if region != "cell":
                return
            col_id = result_tree.identify_column(event.x)
            item = result_tree.identify_row(event.y)
            if not item:
                return
            col_idx = int(col_id.replace("#", "")) - 1
            if col_idx < 0 or col_idx >= len(RESULT_COLS):
                return
            idx = result_tree.index(item)
            old_val = result_tree.set(item, RESULT_COLS[col_idx])

            edit = tk.Toplevel(win)
            edit.title("编辑")
            edit.geometry("420x120")
            edit.transient(win)
            evar = tk.StringVar(value=old_val)
            ttk.Label(edit, text=f"编辑 {RESULT_HEADINGS[RESULT_COLS[col_idx]]}：").pack(padx=10, pady=(10, 2))
            entry = ttk.Entry(edit, textvariable=evar, width=52)
            entry.pack(padx=10, pady=5, fill="x")
            entry.select_range(0, "end")
            entry.focus_set()

            def _save_edit():
                new_val = evar.get().strip()
                result_tree.set(item, RESULT_COLS[col_idx], new_val)
                if idx < len(parsed_entries):
                    parsed_entries[idx][RESULT_COLS[col_idx]] = new_val
                edit.destroy()

            entry.bind("<Return>", lambda e: _save_edit())
            ttk.Button(edit, text="确定", command=_save_edit).pack(pady=5)

        result_tree.bind("<Double-1>", _on_double_click)

        btn_row2 = ttk.Frame(win)
        btn_row2.pack(fill="x", padx=15, pady=10)

        def _do_save():
            if not parsed_entries:
                messagebox.showinfo("提示", "没有可入库的条目。请先点击「AI解析」。")
                return
            saved = 0
            for entry in parsed_entries:
                title = (entry.get("title") or "").strip()
                if not title:
                    continue
                data = {}
                for col in RESULT_COLS:
                    val = (entry.get(col) or "").strip()
                    if val:
                        data[col] = val
                data["title"] = title
                if entry.get("source_url"):
                    data["source_url"] = (entry["source_url"] or "").strip()
                data["source_engine"] = "手动录入"
                db.insert_literature(data)
                saved += 1
            messagebox.showinfo("完成", f"已入库 {saved} 条文献。")
            win.destroy()
            self._refresh()

        ttk.Button(btn_row2, text="确认入库", command=_do_save).pack(side="right", padx=3)
        ttk.Button(btn_row2, text="取消", command=win.destroy).pack(side="right", padx=3)


import tkinter as tk
from tkinter import ttk, messagebox
import database as db
from ui.search_tab import SearchTab
from ui.organize_tab import OrganizeTab
from ui.analysis_tab import AnalysisTab
from ui.chat_assistant import ChatAssistant


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("海外中式建筑数据库")
        self.root.geometry("1500x850")

        self._chat_visible = False
        self._setup_style()
        self._build_menu()
        self._build_notebook()
        self._build_statusbar()

        db.init_db()

        from utils.perf_monitor import start_watchdog
        start_watchdog(self.root)

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="初始化数据库", command=self._reinit_db)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="API 配置…", command=self._open_api_settings)
        menubar.add_cascade(label="设置", menu=settings_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="切换聊天助手 (Ctrl+B)", command=self._toggle_chat)
        menubar.add_cascade(label="视图", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self._show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

    def _build_notebook(self):
        self.pane = ttk.PanedWindow(self.root, orient="horizontal")
        self.pane.pack(fill="both", expand=True, padx=5, pady=5)

        # Left: notebook
        self.notebook = ttk.Notebook(self.pane)
        self.search_tab = SearchTab(self.notebook)
        self.organize_tab = OrganizeTab(self.notebook)
        self.analysis_tab = AnalysisTab(self.notebook)

        self.notebook.add(self.search_tab.frame, text="  案例查询  ")
        self.notebook.add(self.organize_tab.frame, text="  数据整理  ")
        self.notebook.add(self.analysis_tab.frame, text="  数据分析  ")

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self.pane.add(self.notebook, weight=1)

        # Right: chat assistant panel (hidden by default, Ctrl+B to show)
        self.chat_assistant = ChatAssistant(self.pane, self)

        # Keyboard shortcut to toggle chat panel
        self.root.bind("<Control-b>", lambda e: self._toggle_chat())
        self.root.bind("<Command-b>", lambda e: self._toggle_chat())

    def _on_tab_changed(self, event):
        tab_idx = self.notebook.index("current")
        if tab_idx != 1:  # Not on organize tab — clear verify_receiver
            from ui.search_tab import _ExtractServerHandler
            _ExtractServerHandler.verify_receiver = None
        if tab_idx == 2:  # Analysis tab
            self.analysis_tab._lazy_load_chart()

    def _toggle_chat(self):
        if self._chat_visible:
            self.pane.forget(self.chat_assistant)
            self._chat_visible = False
        else:
            self.pane.add(self.chat_assistant, weight=0)
            self._chat_visible = True

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="就绪")
        statusbar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        statusbar.pack(side="bottom", fill="x")

    def _reinit_db(self):
        if messagebox.askyesno("确认", "重新初始化数据库会清空所有数据，确定继续？"):
            db.init_db()
            messagebox.showinfo("完成", "数据库已重新初始化。")
            self.search_tab.refresh()

    def _show_about(self):
        messagebox.showinfo("关于", "海外中式建筑数据库 v1.0\n\n用于海外中式建筑案例的查询、整理与数据分析。")

    def _open_api_settings(self):
        from ui.api_settings import ApiSettingsDialog
        ApiSettingsDialog(self.root)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = MainWindow()
    app.run()

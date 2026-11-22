"""API 配置设置弹窗：集中编辑所有 LLM 通道与专用服务的密钥。

LLM 通道可自定义 base_url + model + api_key；专用服务仅填 api_key。
"""

import tkinter as tk
from tkinter import ttk, messagebox

import config


class ApiSettingsDialog:
    def __init__(self, parent):
        self.cfg = config.load_config()
        self._vars = {}  # id -> {"base_url": var, "model": var, "api_key": var}

        self.win = tk.Toplevel(parent)
        self.win.title("API 配置")
        self.win.geometry("720x600")
        self.win.transient(parent)
        self.win.resizable(False, False)

        self._build()
        self.win.grab_set()

    def _build(self):
        body = ttk.Frame(self.win, padding=12)
        body.pack(fill="both", expand=True)

        hint = ttk.Label(
            body,
            text="LLM 通道走 OpenAI 兼容协议，可填写任意兼容服务（DeepSeek / OpenAI / Qwen / 中转等）。\n"
                 "专用服务为固定 API，仅需填写密钥。所有配置保存在本地，不会随代码公开。",
            foreground="#666",
        )
        hint.pack(anchor="w", pady=(0, 10))

        # 可滚动区域，容纳 4 个 LLM 通道 + 3 个专用服务
        canvas = tk.Canvas(body, highlightthickness=0)
        scroll = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # ---- LLM 通道 ----
        ttk.Label(inner, text="大模型通道（可自定义服务地址 / 模型名）",
                  font=("", 10, "bold")).pack(anchor="w", pady=(0, 4))
        for ch in config.LLM_CHANNELS:
            self._add_llm_channel(inner, ch)

        # ---- 专用服务 ----
        ttk.Label(inner, text="专用服务（固定 API，仅填密钥）",
                  font=("", 10, "bold")).pack(anchor="w", pady=(14, 4))
        for svc in config.SPECIAL_SERVICES:
            self._add_service(inner, svc)

        # ---- 按钮 ----
        btns = ttk.Frame(self.win, padding=(12, 0, 12, 12))
        btns.pack(fill="x")
        ttk.Button(btns, text="保存", command=self._save).pack(side="right")
        ttk.Button(btns, text="取消", command=self.win.destroy).pack(side="right", padx=6)

    def _add_llm_channel(self, parent, ch):
        cid = ch["id"]
        box = ttk.LabelFrame(parent, text=ch["name"], padding=8)
        box.pack(fill="x", padx=0, pady=4)

        saved = self.cfg.get(cid, {})
        v_url = tk.StringVar(value=saved.get("base_url", ch["default_base_url"]))
        v_model = tk.StringVar(value=saved.get("model", ch["default_model"]))
        v_key = tk.StringVar(value=saved.get("api_key", ""))
        self._vars[cid] = {"base_url": v_url, "model": v_model, "api_key": v_key}

        if ch.get("description"):
            ttk.Label(box, text=ch["description"], foreground="#888").pack(
                anchor="w", pady=(0, 4))

        row1 = ttk.Frame(box)
        row1.pack(fill="x", pady=(0, 4))
        ttk.Label(row1, text="服务地址", width=10).pack(side="left")
        ttk.Entry(row1, textvariable=v_url).pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Label(row1, text="模型名", width=8).pack(side="left")
        ttk.Entry(row1, textvariable=v_model, width=22).pack(side="left")

        row2 = ttk.Frame(box)
        row2.pack(fill="x")
        ttk.Label(row2, text="密钥", width=10).pack(side="left")
        ttk.Entry(row2, textvariable=v_key, show="*").pack(side="left", fill="x", expand=True)

    def _add_service(self, parent, svc):
        sid = svc["id"]
        box = ttk.LabelFrame(parent, text=svc["name"], padding=8)
        box.pack(fill="x", padx=0, pady=4)

        v_key = tk.StringVar(value=self.cfg.get(sid, {}).get("api_key", ""))
        self._vars[sid] = {"api_key": v_key}

        row = ttk.Frame(box)
        row.pack(fill="x")
        ttk.Label(row, text="密钥", width=10).pack(side="left")
        ttk.Entry(row, textvariable=v_key, show="*").pack(side="left", fill="x", expand=True)

    def _save(self):
        cfg = {}
        for ch in config.LLM_CHANNELS:
            v = self._vars[ch["id"]]
            cfg[ch["id"]] = {
                "base_url": v["base_url"].get().strip(),
                "model": v["model"].get().strip(),
                "api_key": v["api_key"].get().strip(),
            }
        for svc in config.SPECIAL_SERVICES:
            cfg[svc["id"]] = {"api_key": self._vars[svc["id"]]["api_key"].get().strip()}
        config.save_config(cfg)
        messagebox.showinfo("已保存", "API 配置已保存到本地。")
        self.win.destroy()

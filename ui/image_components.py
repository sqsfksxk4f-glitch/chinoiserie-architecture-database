"""
Shared image thumbnail components used by search_tab and organize_tab.
"""

import io
import json
import os
import ssl
import hashlib
import threading
import tkinter as tk
from tkinter import ttk
from concurrent.futures import ThreadPoolExecutor, as_completed
import urllib.request
import urllib.error
from PIL import Image, ImageTk

# ---------------------------------------------------------------------------
#  Global cache & shared resources
# ---------------------------------------------------------------------------

_cache_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "img_cache")
_CACHE_TTL = 86400  # 1 day


_MAX_WORKERS = 3

_fetch_pool = ThreadPoolExecutor(max_workers=_MAX_WORKERS)


def _ssl_context():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _url_hash(url):
    return hashlib.md5(url.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
#  Disk cache for thumbnails (avoids re-downloading every refresh)
# ---------------------------------------------------------------------------

class _ThumbCache:
    """Simple disk cache that stores thumbnail bytes keyed by URL hash."""

    @staticmethod
    def get(url):
        os.makedirs(_cache_dir, exist_ok=True)
        path = os.path.join(_cache_dir, _url_hash(url))
        if os.path.exists(path):
            mtime = os.path.getmtime(path)
            if (__import__("time").time() - mtime) < _CACHE_TTL:
                with open(path, "rb") as f:
                    return f.read()
        return None

    @staticmethod
    def put(url, data):
        os.makedirs(_cache_dir, exist_ok=True)
        path = os.path.join(_cache_dir, _url_hash(url))
        with open(path, "wb") as f:
            f.write(data)


# ---------------------------------------------------------------------------
#  URL parsing helper
# ---------------------------------------------------------------------------

def parse_image_urls(image_urls_text, image_path=None):
    """Return a deduplicated list of image source strings (URLs or local paths)."""
    urls = []
    if image_path:
        # image_path may be a JSON-encoded list (from _refresh_image_preview),
        # a dbimg:// reference, a file path, or a URL
        try:
            parsed = json.loads(image_path)
            if isinstance(parsed, list):
                urls.extend(str(p) for p in parsed)
            elif isinstance(parsed, str) and parsed.strip():
                urls.append(parsed.strip())
        except (json.JSONDecodeError, TypeError):
            if os.path.exists(image_path) or image_path.startswith("dbimg://"):
                urls.append(image_path)
    if image_urls_text:
        try:
            parsed = json.loads(image_urls_text)
            if isinstance(parsed, list):
                urls.extend(parsed)
            elif isinstance(parsed, str) and parsed.strip():
                urls.append(parsed.strip())
        except (json.JSONDecodeError, TypeError):
            for u in str(image_urls_text).replace("\n", ",").split(","):
                u = u.strip()
                if u:
                    urls.append(u)
    # Deduplicate while preserving order
    # Also skip HTTP URLs that clearly aren't images (e.g. blog posts)
    _NON_IMAGE_EXTS = {".html", ".htm", ".php", ".asp", ".aspx", ".jsp", ".xml", ".json"}
    seen = set()
    deduped = []
    for u in urls:
        if u in seen:
            continue
        seen.add(u)
        if u.startswith(("http://", "https://")):
            low = u.lower().split("?")[0].split("#")[0]
            if any(low.endswith(ext) for ext in _NON_IMAGE_EXTS):
                continue
        deduped.append(u)
    return deduped


def merged_image_sources(case):
    """Return (image_urls_str, image_path_str, all_urls_list) for consistent display.

    image_path is the single source of truth — it reflects the user's curated
    image set from the verification panel.  Do NOT merge extra images from the
    database; that would undo the user's deletions.
    """
    image_urls = case.get("image_urls", "") or ""
    image_path = case.get("image_path", "") or ""

    all_urls = parse_image_urls(image_urls, image_path)
    return image_urls, image_path, all_urls


# ---------------------------------------------------------------------------
#  Core fetch helpers
# ---------------------------------------------------------------------------

def fetch_image_bytes(src, timeout=15):
    """Return (bytes, is_cached:bool) or (None, False) on failure.

    Supports:
      - http:// / https:// URLs (with disk cache)
      - local file paths
      - dbimg://<id> protocol (reads BLOB from case_images table)
    """
    if src.startswith("dbimg://"):
        import database as db
        try:
            img_id = int(src.replace("dbimg://", ""))
            record = db.get_case_image(img_id)
            if record and record.get("image_data"):
                return record["image_data"], True
        except Exception:
            pass
        return None, False

    if src.startswith("fileimg://"):
        import database as db
        try:
            file_name = src.replace("fileimg://", "")
            fpath = os.path.join(db.IMAGE_DIR, file_name)
            if os.path.exists(fpath):
                with open(fpath, "rb") as f:
                    return f.read(), True
        except Exception:
            pass
        return None, False

    if src.startswith(("http://", "https://")):
        cached = _ThumbCache.get(src)
        if cached is not None:
            return cached, True
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
            "Referer": "https://weibo.com/",
            "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
        }
        req = urllib.request.Request(src, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as resp:
                data = resp.read()
        except urllib.error.HTTPError as e:
            if e.code == 403 and headers["Referer"] != "":
                # Retry with empty Referer (some servers reject Referer)
                headers["Referer"] = ""
                req = urllib.request.Request(src, headers=headers)
                with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as resp:
                    data = resp.read()
            else:
                raise
        _ThumbCache.put(src, data)
        return data, False
    else:
        if os.path.exists(src):
            with open(src, "rb") as f:
                return f.read(), True
        return None, False


def open_image_from_bytes(data):
    return Image.open(io.BytesIO(data))


def open_image_file(path):
    return Image.open(path)


# ---------------------------------------------------------------------------
#  Thumbnail strip widget  (shared)
# ---------------------------------------------------------------------------

class ImageThumbnailStrip(ttk.Frame):
    """Horizontal scrolling strip of image thumbnails (shown in detail / edit dialogs)."""

    THUMB_SIZE = 200

    def __init__(self, parent, image_urls_text, image_path=None, on_delete=None, on_reorder=None, case_id=None, on_crop_save=None, image_metadata=None):
        super().__init__(parent)
        self.image_urls_text = image_urls_text
        self.image_path = image_path
        self.on_delete = on_delete  # callback(src) when delete button clicked
        self.on_reorder = on_reorder  # callback(new_ordered_dbimg_list) when images are reordered
        self.case_id = case_id
        self.on_crop_save = on_crop_save  # callback(cropped_pil_image, src) when crop is saved
        self.image_metadata = image_metadata or {}  # {dbimg://<id>: {author, image_time}}
        self._photos = []
        self._dbimg_order = []  # ordered list of dbimg:// src strings
        self._build()
        self._load_async()

    def _build(self):
        self._canvas = tk.Canvas(self, height=self.THUMB_SIZE + 70, highlightthickness=0)
        h_scroll = ttk.Scrollbar(self, orient="horizontal", command=self._canvas.xview)
        self._inner = ttk.Frame(self._canvas)
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._canvas.configure(xscrollcommand=h_scroll.set)
        self._inner.bind("<Configure>", lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox("all")))
        self._canvas.pack(fill="x", expand=True, pady=2)
        h_scroll.pack(fill="x")

    def _load_async(self):
        urls = parse_image_urls(self.image_urls_text, self.image_path)
        # Remember dbimg order from image_path (used for reorder buttons)
        self._dbimg_order = [u for u in urls if u.startswith("dbimg://")]
        if not urls:
            ttk.Label(self._inner, text="暂无图片", foreground="gray",
                      font=("", 9)).pack(padx=10, pady=10)
            return
        self._loading = ttk.Label(self._inner, text=f"加载 {len(urls)} 张图片中...",
                                  foreground="gray")
        self._loading.pack(padx=10, pady=10)
        threading.Thread(target=self._fetch_all, args=(urls,), daemon=True).start()

    def _fetch_all(self, urls):
        results = []
        futures = {}
        with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as pool:
            for src in urls:
                futures[pool.submit(fetch_image_bytes, src, 15)] = src

            for future in as_completed(futures):
                src = futures[future]
                try:
                    data, _cached = future.result()
                    if data is None:
                        continue
                    img = open_image_from_bytes(data)
                    img.thumbnail((self.THUMB_SIZE, self.THUMB_SIZE), Image.LANCZOS)
                    results.append((img, src))
                except Exception:
                    pass

        # Keep original order
        results.sort(key=lambda x: urls.index(x[1]) if x[1] in urls else len(urls))
        # Convert PIL Images to bytes to avoid GIL issues when creating
        # ImageTk.PhotoImage on the main thread
        safe_results = []
        for img, src in results:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            safe_results.append((buf.getvalue(), src))
        self.after(0, self._show_thumbnails, safe_results)

    def _show_thumbnails(self, results):
        if not self._inner.winfo_exists():
            return
        if hasattr(self, "_loading") and self._loading:
            self._loading.destroy()
            self._loading = None
        if not results:
            ttk.Label(self._inner, text="无法加载图片", foreground="gray",
                      font=("", 9)).pack(padx=10, pady=10)
            return
        for i, (img_bytes, src) in enumerate(results):
            img = Image.open(io.BytesIO(img_bytes))
            photo = ImageTk.PhotoImage(img)
            self._photos.append(photo)

            frame = ttk.Frame(self._inner)
            frame.pack(side="left", padx=4, pady=4)

            lbl = ttk.Label(frame, image=photo, cursor="hand2")
            lbl.image = photo
            lbl.pack()

            src_short = src if len(src) <= 35 else src[:32] + "..."
            ttk.Label(frame, text=src_short, font=("", 7), foreground="gray").pack()

            # Show author and image_time for all images
            meta = self.image_metadata.get(src, {})
            author_val = (meta.get("author") or "").strip()
            time_val = (meta.get("image_time") or "").strip()
            if author_val:
                ttk.Label(frame, text=f"作者: {author_val}", font=("", 7), foreground="#555").pack()
            if time_val:
                ttk.Label(frame, text=f"拍摄: {time_val}", font=("", 7), foreground="#555").pack()

            # Reorder + delete buttons for dbimg:// images
            if src.startswith("dbimg://") and self.on_reorder:
                btn_row = ttk.Frame(frame)
                btn_row.pack(pady=(2, 0))
                idx = self._dbimg_order.index(src) if src in self._dbimg_order else -1
                left_btn = ttk.Button(btn_row, text="←", width=2,
                                      command=lambda s=src: self._move_left(s))
                left_btn.pack(side="left", padx=1)
                right_btn = ttk.Button(btn_row, text="→", width=2,
                                       command=lambda s=src: self._move_right(s))
                right_btn.pack(side="left", padx=1)
                if self.on_delete:
                    del_btn = ttk.Button(btn_row, text="×", width=2,
                                         command=lambda s=src, f=frame: self._on_delete_click(s, f))
                    del_btn.pack(side="left", padx=1)
            elif self.on_delete and src.startswith("dbimg://"):
                del_btn = ttk.Button(frame, text="× 删除", width=5,
                                     command=lambda s=src, f=frame: self._on_delete_click(s, f))
                del_btn.pack()

            lbl.bind("<Button-1>", lambda e, s=src: self._open_full(s))
            frame.bind("<Button-1>", lambda e, s=src: self._open_full(s))

    def _move_left(self, src):
        if src not in self._dbimg_order:
            return
        idx = self._dbimg_order.index(src)
        if idx <= 0:
            return
        self._dbimg_order[idx], self._dbimg_order[idx - 1] = self._dbimg_order[idx - 1], self._dbimg_order[idx]
        if self.on_reorder:
            self.on_reorder(list(self._dbimg_order))

    def _move_right(self, src):
        if src not in self._dbimg_order:
            return
        idx = self._dbimg_order.index(src)
        if idx >= len(self._dbimg_order) - 1:
            return
        self._dbimg_order[idx], self._dbimg_order[idx + 1] = self._dbimg_order[idx + 1], self._dbimg_order[idx]
        if self.on_reorder:
            self.on_reorder(list(self._dbimg_order))

    def _on_delete_click(self, src, frame):
        if self.on_delete:
            self.on_delete(src)
        frame.destroy()

    def _open_full(self, src):
        FullImageViewer(self.winfo_toplevel(), src,
                        case_id=self.case_id,
                        on_crop_save=self.on_crop_save,
                        on_metadata_update=self._on_metadata_update)

    def _on_metadata_update(self):
        """Called when image metadata is edited in FullImageViewer. Rebuild strip."""
        self.image_metadata = self._load_all_metadata()
        for w in self._inner.winfo_children():
            w.destroy()
        self._photos.clear()
        urls = parse_image_urls(self.image_urls_text, self.image_path)
        self._dbimg_order = [u for u in urls if u.startswith("dbimg://")]
        if urls:
            threading.Thread(target=self._fetch_all, args=(urls,), daemon=True).start()

    def _load_all_metadata(self):
        """Load image metadata from both case_images table and raw_cases.image_metadata JSON."""
        result = {}
        if not self.case_id:
            return result
        import database as db
        # Load from case_images table (for dbimg:// images)
        try:
            images = db.get_images_for_raw_case(self.case_id)
            for img in images:
                key = f"dbimg://{img['id']}"
                result[key] = {"author": img.get("author", ""), "image_time": img.get("image_time", "")}
        except Exception:
            pass
        # Load from raw_cases.image_metadata JSON (for all images including URLs)
        try:
            conn = db.get_connection()
            row = conn.execute("SELECT image_metadata FROM raw_cases WHERE id = ?", (self.case_id,)).fetchone()
            conn.close()
            if row and row["image_metadata"]:
                import json
                all_meta = json.loads(row["image_metadata"])
                for src, meta in all_meta.items():
                    if src not in result:
                        result[src] = meta
                    else:
                        # dbimg entries from case_images take priority, but fill gaps
                        if not result[src].get("author") and meta.get("author"):
                            result[src]["author"] = meta["author"]
                        if not result[src].get("image_time") and meta.get("image_time"):
                            result[src]["image_time"] = meta["image_time"]
        except Exception:
            pass
        return result
        for w in self._inner.winfo_children():
            w.destroy()
        self._photos.clear()
        urls = parse_image_urls(self.image_urls_text, self.image_path)
        self._dbimg_order = [u for u in urls if u.startswith("dbimg://")]
        if urls:
            threading.Thread(target=self._fetch_all, args=(urls,), daemon=True).start()


# ---------------------------------------------------------------------------
#  Full-size image viewer  (shared)
# ---------------------------------------------------------------------------

class FullImageViewer:
    """Popup window that loads and displays a full-resolution image, with optional crop."""

    def __init__(self, parent, src, case_id=None, on_crop_save=None, on_metadata_update=None):
        self.src = src
        self.case_id = case_id
        self.on_crop_save = on_crop_save
        self.on_metadata_update = on_metadata_update
        self._img_id = None
        self._meta_frame = None
        self.win = tk.Toplevel(parent)
        self.win.title("图片查看")
        self.win.geometry("600x500")
        self.win.transient(parent)

        self._crop_mode = False
        self._crop_start = None
        self._crop_rect = None
        self._crop_coords = None  # (x1, y1, x2, y2) on displayed image
        self._photo = None
        self._pil_image = None
        self._canvas = None
        self._save_btn = None
        self._crop_btn = None
        self._cancel_crop_btn = None
        self._status_label = None

        ttk.Label(self.win, text="加载中...", font=("", 11)).pack(expand=True)
        self.win.bind("<Escape>", lambda e: self.win.destroy())
        threading.Thread(target=self._load_and_show, daemon=True).start()

    def _load_and_show(self):
        try:
            data, _ = fetch_image_bytes(self.src, timeout=30)
            if data is None:
                raise FileNotFoundError(self.src)
            img = open_image_from_bytes(data)
            self._pil_image = img

            screen_w = self.win.winfo_screenwidth()
            screen_h = self.win.winfo_screenheight()
            max_w = int(screen_w * 0.85)
            max_h = int(screen_h * 0.85)
            img.thumbnail((max_w, max_h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self._photo = photo

            def show():
                for w in self.win.winfo_children():
                    w.destroy()

                win_w = min(img.width + 20, int(screen_w * 0.9))
                win_h = min(img.height + 180, int(screen_h * 0.9))
                self.win.geometry(f"{win_w}x{win_h}")

                # Scrollable area (image + metadata)
                outer = ttk.Frame(self.win)
                outer.pack(side="top", fill="both", expand=True)

                scroll_canvas = tk.Canvas(outer, highlightthickness=0)
                scrollbar = ttk.Scrollbar(outer, orient="vertical", command=scroll_canvas.yview)
                scroll_frame = ttk.Frame(scroll_canvas)

                scroll_frame.bind("<Configure>",
                    lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")))
                scroll_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
                scroll_canvas.configure(yscrollcommand=scrollbar.set)

                scroll_canvas.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")

                # Mouse wheel scrolling
                def _on_mousewheel(event):
                    scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                scroll_canvas.bind("<Enter>", lambda e: scroll_canvas.bind_all("<MouseWheel>", _on_mousewheel))
                scroll_canvas.bind("<Leave>", lambda e: scroll_canvas.unbind_all("<MouseWheel>"))

                # Image canvas
                canvas = tk.Canvas(scroll_frame, width=img.width, height=img.height)
                canvas.pack(side="top")
                canvas.create_image(0, 0, anchor="nw", image=photo)
                canvas.image = photo
                canvas.bind("<Button-1>", lambda e: self.win.destroy() if not self._crop_mode else None)
                self._canvas = canvas

                # Metadata editing panel
                self._meta_frame = ttk.LabelFrame(scroll_frame, text="图片信息", padding=5)
                self._meta_frame.pack(side="top", fill="x", padx=5, pady=(5, 5))

                # Load metadata
                author_val = ""
                time_val = ""
                if self.case_id and self.src:
                    import database as db
                    if self.src.startswith("dbimg://"):
                        try:
                            img_id = int(self.src.replace("dbimg://", ""))
                            self._img_id = img_id
                            rec = db.get_case_image(img_id)
                            if rec:
                                author_val = (rec.get("author") or "").strip()
                                time_val = (rec.get("image_time") or "").strip()
                        except Exception:
                            pass
                    if not author_val and not time_val:
                        meta = db.get_image_metadata(self.case_id, self.src)
                        author_val = (meta.get("author") or "").strip()
                        time_val = (meta.get("image_time") or "").strip()

                author_row = ttk.Frame(self._meta_frame)
                author_row.pack(fill="x", pady=2)
                ttk.Label(author_row, text="作者:", width=8, font=("", 9)).pack(side="left")
                author_var = tk.StringVar(value=author_val)
                author_entry = ttk.Entry(author_row, textvariable=author_var, font=("", 9))
                author_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

                time_row = ttk.Frame(self._meta_frame)
                time_row.pack(fill="x", pady=2)
                ttk.Label(time_row, text="拍摄时间:", width=8, font=("", 9)).pack(side="left")
                time_var = tk.StringVar(value=time_val)
                time_entry = ttk.Entry(time_row, textvariable=time_var, font=("", 9))
                time_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

                tip = ttk.Label(self._meta_frame, text="填入作者和拍摄时间后点击保存",
                                font=("", 7), foreground="#999")
                tip.pack(side="left", pady=(3, 0))
                save_meta_btn = ttk.Button(self._meta_frame, text="保存信息",
                                           command=lambda av=author_var, tv=time_var: self._save_metadata(av.get(), tv.get()))
                save_meta_btn.pack(side="right", pady=(3, 0))

                # Button bar (fixed at bottom)
                btn_bar = ttk.Frame(self.win, height=40)
                btn_bar.pack(side="bottom", fill="x", pady=(5, 5))
                btn_bar.pack_propagate(False)

                self._status_label = ttk.Label(btn_bar, text="",
                                               font=("", 9), foreground="#666")
                self._status_label.pack(side="left", padx=5)

                self._save_btn = ttk.Button(btn_bar, text="保存", state="disabled",
                                            command=self._do_crop)
                self._save_btn.pack(side="right", padx=3)

                self._cancel_crop_btn = ttk.Button(btn_bar, text="取消裁剪", state="disabled",
                                                   command=self._exit_crop_mode)
                self._cancel_crop_btn.pack(side="right", padx=3)

                self._crop_btn = ttk.Button(btn_bar, text="裁剪",
                                            command=self._enter_crop_mode)
                self._crop_btn.pack(side="right", padx=3)

                ttk.Button(btn_bar, text="关闭",
                           command=self.win.destroy).pack(side="right", padx=3)

            self.win.after(0, show)
        except Exception as e:
            def show_err():
                for w in self.win.winfo_children():
                    w.destroy()
                ttk.Label(self.win, text=f"加载失败: {e}", wraplength=400).pack(expand=True)
            self.win.after(0, show_err)

    def _enter_crop_mode(self):
        self._crop_mode = True
        self._canvas.config(cursor="cross")
        self._canvas.bind("<ButtonPress-1>", self._on_press)
        self._canvas.bind("<B1-Motion>", self._on_drag)
        self._canvas.bind("<ButtonRelease-1>", self._on_release)
        self._crop_btn.config(state="disabled")
        self._cancel_crop_btn.config(state="normal")
        self._status_label.config(text="拖拽鼠标框选要保留的区域")
        self.win.title("图片查看 - 裁剪中...")

    def _exit_crop_mode(self):
        self._crop_mode = False
        self._canvas.config(cursor="")
        self._canvas.unbind("<ButtonPress-1>")
        self._canvas.unbind("<B1-Motion>")
        self._canvas.unbind("<ButtonRelease-1>")
        if self._crop_rect:
            self._canvas.delete(self._crop_rect)
            self._crop_rect = None
        self._crop_coords = None
        self._crop_start = None
        self._crop_btn.config(state="normal")
        self._cancel_crop_btn.config(state="disabled")
        self._save_btn.config(state="disabled")
        self._status_label.config(text="")
        self.win.title("图片查看")

    def _on_press(self, event):
        self._crop_start = (event.x, event.y)
        if self._crop_rect:
            self._canvas.delete(self._crop_rect)
            self._crop_rect = None

    def _on_drag(self, event):
        if not self._crop_start:
            return
        x1, y1 = self._crop_start
        x2, y2 = event.x, event.y
        if self._crop_rect:
            self._canvas.coords(self._crop_rect, x1, y1, x2, y2)
        else:
            self._crop_rect = self._canvas.create_rectangle(
                x1, y1, x2, y2, outline="red", width=2, dash=(4, 2))

    def _on_release(self, event):
        if not self._crop_start:
            return
        x1, y1 = self._crop_start
        x2, y2 = event.x, event.y
        if abs(x2 - x1) < 10 or abs(y2 - y1) < 10:
            if self._crop_rect:
                self._canvas.delete(self._crop_rect)
                self._crop_rect = None
            self._crop_start = None
            self._save_btn.config(state="disabled")
            self._status_label.config(text="拖拽鼠标框选要保留的区域")
            return
        self._crop_coords = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        self._save_btn.config(state="normal")
        self._status_label.config(text=f"已选区: {abs(x2-x1)}x{abs(y2-y1)}px，点击「保存」确认")
        self._crop_start = None

    def _do_crop(self):
        if not self._crop_coords or not self._pil_image:
            return
        x1, y1, x2, y2 = self._crop_coords
        cropped = self._pil_image.crop((x1, y1, x2, y2))

        if self.on_crop_save:
            self.on_crop_save(cropped, self.src)
            self.win.destroy()
        else:
            from tkinter import filedialog
            path = filedialog.asksaveasfilename(
                parent=self.win,
                title="保存裁剪图片",
                defaultextension=".jpg",
                filetypes=[("JPEG", "*.jpg"), ("PNG", "*.png"), ("所有文件", "*.*")])
            if path:
                cropped.save(path)
            self.win.destroy()

    def _save_metadata(self, author, image_time):
        """Persist author and image_time. Uses case_images table for dbimg://, plus JSON on raw_cases."""
        a = author.strip() or None
        t = image_time.strip() or None
        import database as db
        # Save to case_images table for dbimg:// images
        if self._img_id:
            db.update_case_image(self._img_id, {"author": a, "image_time": t})
        # Always save to raw_cases.image_metadata JSON (works for all image types)
        if self.case_id and self.src:
            db.set_image_metadata(self.case_id, self.src, a, t)
        if self.on_metadata_update:
            self.on_metadata_update()


# ---------------------------------------------------------------------------
#  Tree thumbnail helpers  (shared logic for both tabs)
# ---------------------------------------------------------------------------

def load_tree_thumbnail(tree, item_id, url, case_id, photo_cache):
    """Load a thumbnail for a tree row.  Safe to run in a background thread.

    Parameters
    ----------
    tree : ttk.Treeview
    item_id : str  -- tree item id
    url : str      -- single image URL or local path
    case_id : int
    photo_cache : dict[int, ImageTk.PhotoImage]
    """
    try:
        data, _ = fetch_image_bytes(url, timeout=10)
        if data is None:
            return
        img = open_image_from_bytes(data)
        img.thumbnail((80, 60), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)

        def update_tree():
            try:
                if tree.exists(item_id):
                    photo_cache[case_id] = photo
                    tree.item(item_id, image=photo, text="")
            except Exception:
                pass
        tree.after(0, update_tree)
    except Exception:
        pass

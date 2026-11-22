import tkinter as tk
from tkinter import ttk, messagebox
import database as db
import country_mapping
from utils.import_export import export_to_excel
from ui.spot_check_plot import draw_spot_check_matrix

try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.font_manager as fm
    # Use CJK-supporting fonts available on macOS
    _cjk_fonts = [f.name for f in fm.fontManager.ttflist
                  if any(k in f.name for k in ['PingFang SC', 'PingFang HK', 'Heiti SC', 'STHeiti',
                                               'Arial Unicode MS', 'Songti SC', 'STSong'])]
    if _cjk_fonts:
        matplotlib.rcParams['font.sans-serif'] = [_cjk_fonts[0]] + matplotlib.rcParams['font.sans-serif']
        matplotlib.rcParams['axes.unicode_minus'] = False
    HAS_MPL = True
except ImportError:
    HAS_MPL = False


def _to_float(v):
    """Normalize coordinate strings (special hyphens) to float."""
    if v is None:
        return None
    s = str(v).replace("‑", "-").replace("−", "-").replace("–", "-").replace("—", "-").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _norm_coords(rows):
    """Convert sqlite rows to dicts with float latitude/longitude."""
    out = []
    for r in rows:
        d = dict(r)
        d["latitude"] = _to_float(d.get("latitude"))
        d["longitude"] = _to_float(d.get("longitude"))
        out.append(d)
    return out


class AnalysisTab:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self._loaded = False
        self._world_drawn = False
        self._coastal_drawn = False
        self._continent_drawn = False
        self._sankey_drawn = False
        self._chord_drawn = False
        self._coastal_alt_drawn = False
        self._temporal_drawn = False
        self._demolition_drawn = False
        self._flowchart_drawn = False
        self._gdp_drawn = False
        self._pearson_drawn = False
        self._country_choro_drawn = False
        self._channel_drawn = False
        self._sd_world_drawn = False
        self._spot_check_drawn = False
        self._sa_time_drawn = False
        self._sa_gdp_drawn = False
        self._current_fig = None
        self._tab_figs = {}
        self._coastal_lines = []
        self._coastal_canvas = None
        self._coastal_state = None
        self._load_coastal_lines()
        self._build_upper()
        self._build_chart_area()

    def _lazy_load_chart(self):
        """Draw the first visible figure when the user switches to the Analysis tab."""
        if HAS_MPL and not self._flowchart_drawn:
            self.frame.after(100, lambda: self._draw_flowchart(self._flowchart_frame))

    def _build_upper(self):
        bar = ttk.Frame(self.frame)
        bar.pack(fill="x", padx=5, pady=5)

        ttk.Button(bar, text="导出全部数据到 Excel", command=self._export_all).pack(side="left", padx=3)
        ttk.Button(bar, text="导出图表 PNG", command=self._export_png).pack(side="left", padx=3)

        self.stats_label = ttk.Label(bar, text="")
        self.stats_label.pack(side="right", padx=10)

    def _build_chart_area(self):
        self.map_notebook = ttk.Notebook(self.frame)
        self.map_notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self._world_map_frame = ttk.Frame(self.map_notebook)
        self._coastal_frame = ttk.Frame(self.map_notebook)
        self._continent_frame = ttk.Frame(self.map_notebook)
        self._country_choro_frame = ttk.Frame(self.map_notebook)
        self._sankey_frame = ttk.Frame(self.map_notebook)
        self._chord_frame = ttk.Frame(self.map_notebook)
        self._coastal_alt_frame = ttk.Frame(self.map_notebook)
        self._temporal_frame = ttk.Frame(self.map_notebook)
        self._demolition_frame = ttk.Frame(self.map_notebook)
        self._flowchart_frame = ttk.Frame(self.map_notebook)
        self._gdp_frame = ttk.Frame(self.map_notebook)
        self._pearson_frame = ttk.Frame(self.map_notebook)
        self._sd_world_frame = ttk.Frame(self.map_notebook)
        self._channel_fig_frame = ttk.Frame(self.map_notebook)
        self._spot_check_frame = ttk.Frame(self.map_notebook)
        self._sa_time_frame = ttk.Frame(self.map_notebook)
        self._sa_gdp_frame = ttk.Frame(self.map_notebook)

        self.map_notebook.add(self._flowchart_frame, text="  研究设计  ")
        self.map_notebook.add(self._world_map_frame, text="  世界流布图  ")
        self.map_notebook.add(self._sd_world_frame, text="  SD 世界流布图  ")
        self.map_notebook.add(self._coastal_frame, text="  空间分布海陆分析  ")
        self.map_notebook.add(self._coastal_alt_frame, text="  海陆分析备选  ")
        self.map_notebook.add(self._continent_frame, text="  五大洲分析  ")
        self.map_notebook.add(self._country_choro_frame, text="  案例国家分布图  ")
        self.map_notebook.add(self._sankey_frame, text="  国家组群桑基图  ")
        self.map_notebook.add(self._chord_frame, text="  城市类型弦图  ")
        self.map_notebook.add(self._temporal_frame, text="  时间分布折线  ")
        self.map_notebook.add(self._demolition_frame, text="  损毁年份折线  ")
        self.map_notebook.add(self._gdp_frame, text="  建设与GDP  ")
        self.map_notebook.add(self._pearson_frame, text="  Pearson相关  ")
        self.map_notebook.add(self._channel_fig_frame, text="  七渠道分析图  ")
        self.map_notebook.add(self._spot_check_frame, text="  SD抽检矩阵  ")
        self.map_notebook.add(self._sa_time_frame, text="  SA 建设+损毁  ")
        self.map_notebook.add(self._sa_gdp_frame, text="  SA 建设GDP+Pearson  ")

        self.map_notebook.bind("<<NotebookTabChanged>>", self._on_map_tab_changed)

        if not HAS_MPL:
            ttk.Label(self._world_map_frame, text="Please install matplotlib: pip install matplotlib",
                      font=("", 12)).pack(expand=True)

    def _on_map_tab_changed(self, event):
        if not self.frame.winfo_ismapped():
            return
        tab_idx = self.map_notebook.index("current")
        _old_fig = self._current_fig
        if HAS_MPL:
            if tab_idx == 0 and not self._flowchart_drawn:
                self._draw_flowchart(self._flowchart_frame)
            elif tab_idx == 1 and not self._world_drawn:
                self._draw_world_map(self._world_map_frame)
            elif tab_idx == 2 and not self._sd_world_drawn:
                self._draw_sd_world_map(self._sd_world_frame)
            elif tab_idx == 3 and not self._coastal_drawn:
                self._draw_coastal_map(self._coastal_frame)
            elif tab_idx == 4 and not self._coastal_alt_drawn:
                self._draw_coastal_alt(self._coastal_alt_frame)
            elif tab_idx == 5 and not self._continent_drawn:
                self._draw_continent_map(self._continent_frame)
            elif tab_idx == 6 and not self._country_choro_drawn:
                self._draw_country_choropleth(self._country_choro_frame)
            elif tab_idx == 7 and not self._sankey_drawn:
                self._draw_sankey_map(self._sankey_frame)
            elif tab_idx == 8 and not self._chord_drawn:
                self._draw_chord_map(self._chord_frame)
            elif tab_idx == 9 and not self._temporal_drawn:
                self._draw_temporal_line(self._temporal_frame)
            elif tab_idx == 10 and not self._demolition_drawn:
                self._draw_demolition_line(self._demolition_frame)
            elif tab_idx == 11 and not self._gdp_drawn:
                self._draw_gdp_chart(self._gdp_frame)
            elif tab_idx == 12 and not self._pearson_drawn:
                self._draw_pearson_chart(self._pearson_frame)
            elif tab_idx == 13 and not self._channel_drawn:
                self._draw_channel_figure(self._channel_fig_frame)
            elif tab_idx == 14 and not self._spot_check_drawn:
                self._current_fig = draw_spot_check_matrix(self._spot_check_frame)
                self._spot_check_drawn = True
            elif tab_idx == 15 and not self._sa_time_drawn:
                self._draw_sa_time_figure(self._sa_time_frame)
            elif tab_idx == 16 and not self._sa_gdp_drawn:
                self._draw_sa_gdp_figure(self._sa_gdp_frame)
        # Cache the figure for the current tab so that re-selecting an
        # already-drawn tab still exports the correct figure.
        if self._current_fig is not None and self._current_fig is not _old_fig:
            self._tab_figs[tab_idx] = self._current_fig

    def _export_all(self):
        df = db.export_all_to_dataframe()
        if df.empty:
            messagebox.showwarning("导出", "数据库中没有数据。")
            return
        if export_to_excel(df, "导出全部数据"):
            messagebox.showinfo("导出完成", "数据已成功导出。")

    def _export_png(self):
        from tkinter import filedialog

        tab_idx = self.map_notebook.index("current")
        tab_text = self.map_notebook.tab(tab_idx, "text").strip()
        fig = self._tab_figs.get(tab_idx) or getattr(self, "_current_fig", None)
        if fig is None:
            messagebox.showwarning("导出", "请先生成图表。")
            return

        default_name = (tab_text or "chart").replace(" ", "_")
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG 图片", "*.png")],
            initialfile=f"{default_name}.png",
        )
        if path:
            # 固定长宽比的图（如三面板时间分布图）自选无损导出：TkAgg 会把
            # figure 缩放到画布控件尺寸导致压扁，这里恢复其原始尺寸并按印刷
            # 分辨率导出。其余图保持原行为（去首轴标题后直接 savefig）。
            intended = getattr(fig, "_intended_size_inches", None)
            strip_main = intended is None
            main_ax = fig.axes[0] if fig.axes else None
            saved = main_ax.get_title() if (strip_main and main_ax) else ""
            if strip_main and main_ax:
                main_ax.set_title("")
            try:
                if intended is not None:
                    fig.set_size_inches(*intended)
                    fig.savefig(path, dpi=300, bbox_inches="tight")
                else:
                    fig.savefig(path)
            finally:
                if strip_main and main_ax:
                    main_ax.set_title(saved)
            messagebox.showinfo("导出完成", f"已保存到：{path}")

    @staticmethod
    def _extract_year(year_text):
        """Sorting key: exact founding year (3-4 digits) -> int; century-only
        date -> century midpoint; empty/unknown -> 9999. Matches the temporal
        analysis binning convention (century midpoint)."""
        import re
        ys = (year_text or "").strip()
        m = re.match(r"^(\d{3,4})$", ys)
        if m:
            return int(m.group(1))
        m = re.match(r"^(\d{1,2})(st|nd|rd|th)$", ys)
        if m:
            return (int(m.group(1)) - 1) * 100 + 50
        return 9999

    def _draw_world_map(self, parent_frame):
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3
        import numpy as np
        from collections import defaultdict

        try:
            import cartopy.crs as ccrs
            import cartopy.feature as cfeature
        except ImportError:
            ttk.Label(parent_frame, text="请安装 cartopy：pip install cartopy",
                     font=("", 12)).pack(expand=True)
            return

        try:
            import mplcursors
            HAS_CURSORS = True
        except ImportError:
            HAS_CURSORS = False

        conn = sqlite3.connect(db.DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(
            "SELECT id, name, name_en, country, latitude, longitude, year_built FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL AND name IS NOT NULL AND name != ''"
        )
        rows = cur.fetchall()
        conn.close()
        rows = _norm_coords(rows)

        if not rows:
            ttk.Label(parent_frame, text="暂无经纬度数据", font=("", 12)).pack(expand=True)
            return

        # Load country CN→EN mapping (short-name fuzzy match)
        cn_to_en = {}
        try:
            import pandas as pd
            import os as _os
            xlsx_path = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))),
                                      "..", "所有国家中英文.xlsx")
            df_map = pd.read_excel(xlsx_path)
            for _, row in df_map.iterrows():
                cn_full = str(row["中文"]).strip()
                en = str(row["英文"]).strip()
                if cn_full and en and cn_full != "nan" and en != "nan":
                    cn_to_en[cn_full] = en
                    if len(cn_full) >= 2:
                        key = cn_full[:2]
                        if key not in cn_to_en:
                            cn_to_en[key] = en
            for exact, en in [("俄罗斯", "Russian Federation"), ("韩国", "Republic of Korea"),
                              ("朝鲜", "Democratic People's Republic of Korea"),
                              ("文莱", "Brunei Darussalam"),
                              ("老挝", "Lao People's Democratic Republic"),
                              ("澳大利亚", "Australia")]:
                cn_to_en[exact] = en
        except Exception:
            pass

        # Group by country, find earliest case per country
        country_cases = defaultdict(list)
        for r in rows:
            co = (r["country"] or "Unknown").strip()
            country_cases[co].append(r)

        earliest_ids = set()
        for co, cases in country_cases.items():
            cases_sorted = sorted(cases, key=lambda r: self._extract_year(r["year_built"]))
            earliest = cases_sorted[0]
            earliest_ids.add(earliest["id"])

        # Build point arrays + labels (English where available)
        all_lons, all_lats, all_labels = [], [], []
        earliest_annotations = []

        for r in rows:
            co_cn = (r["country"] or "Unknown").strip()
            co_en = cn_to_en.get(co_cn, co_cn)
            name_en = (r["name_en"] or "").strip()
            name_display = name_en if name_en else (r["name"] or "?")
            all_lons.append(r["longitude"])
            all_lats.append(r["latitude"])
            all_labels.append(f"{co_en}\n{name_display}")
            if r["id"] in earliest_ids:
                earliest_annotations.append((r["longitude"], r["latitude"],
                                             f"{co_en}\n{name_display}", co_en))

        n_countries = len(country_cases)
        self.stats_label.config(text=f"共 {len(rows)} 个案例，{n_countries} 个国家/地区")

        # --- Figure ---
        fig = Figure(figsize=(8, 4.5), dpi=100)
        ax = fig.add_subplot(111, projection=ccrs.Robinson())
        ax.set_global()

        ax.add_feature(cfeature.LAND, facecolor="#d9dde0", edgecolor="none", zorder=0)
        ax.add_feature(cfeature.OCEAN, facecolor="#f5f8fa", zorder=0)
        ax.gridlines(draw_labels=False, linewidth=0.3, color="#aaaaaa", alpha=0.4, linestyle="--")
        ax.spines["geo"].set_visible(False)

        scatters = []
        if all_lons:
            s = ax.scatter(
                all_lons, all_lats,
                s=80,
                c="#3b82f6",
                alpha=0.3,
                edgecolors="none",
                transform=ccrs.PlateCarree(),
                zorder=5,
            )
            scatters.append(s)

        ax.set_title("海外中式建筑世界流布图", fontsize=16, fontweight="bold", pad=12)

        fig.subplots_adjust(bottom=0.01, top=0.99, left=0.01, right=0.99)

        self._current_fig = fig

        # --- Embed in tkinter ---
        # Draw once first so a renderer exists for measuring label boxes.
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.draw()

        # Permanent annotations for earliest case per country (overlap-avoiding placement)
        self._place_country_labels(fig, ax, canvas, ccrs, earliest_annotations,
                                   all_lons, all_lats)

        # Hover tooltips (all cases)
        if HAS_CURSORS and scatters:
            cursor = mplcursors.cursor(scatters, hover=mplcursors.HoverMode.Transient)
            @cursor.connect("add")
            def _on_add(sel):
                idx = sel.index
                if idx < len(all_labels):
                    sel.annotation.set_text(all_labels[idx])
                    sel.annotation.set_fontsize(8)
                    sel.annotation.set_bbox(dict(
                        facecolor="white", alpha=0.9, edgecolor="#ccc",
                        boxstyle="round,pad=0.3"
                    ))

        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._world_drawn = True

    def _draw_sd_world_map(self, parent_frame):
        """SD Data Descriptor world distribution map (English, Robinson projection)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3
        import numpy as np
        from collections import defaultdict

        try:
            import cartopy.crs as ccrs
            import cartopy.feature as cfeature
        except ImportError:
            ttk.Label(parent_frame, text="Please install cartopy: pip install cartopy",
                     font=("", 12)).pack(expand=True)
            return

        try:
            import mplcursors
            HAS_CURSORS = True
        except ImportError:
            HAS_CURSORS = False

        conn = sqlite3.connect(db.DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(
            "SELECT id, name, name_en, country, latitude, longitude, year_built FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL AND name IS NOT NULL AND name != ''"
        )
        rows = cur.fetchall()
        conn.close()
        rows = _norm_coords(rows)

        if not rows:
            ttk.Label(parent_frame, text="No data with coordinates.", font=("", 12)).pack(expand=True)
            return

        # Load country CN→EN mapping
        cn_to_en = {}
        try:
            import pandas as pd
            import os as _os
            xlsx_path = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))),
                                      "..", "所有国家中英文.xlsx")
            df_map = pd.read_excel(xlsx_path)
            for _, row in df_map.iterrows():
                cn_full = str(row["中文"]).strip()
                en = str(row["英文"]).strip()
                if cn_full and en and cn_full != "nan" and en != "nan":
                    cn_to_en[cn_full] = en
                    if len(cn_full) >= 2:
                        key = cn_full[:2]
                        if key not in cn_to_en:
                            cn_to_en[key] = en
            for exact, en in [("俄罗斯", "Russian Federation"), ("韩国", "Republic of Korea"),
                              ("朝鲜", "Democratic People's Republic of Korea"),
                              ("文莱", "Brunei Darussalam"),
                              ("老挝", "Lao People's Democratic Republic"),
                              ("澳大利亚", "Australia")]:
                cn_to_en[exact] = en
        except Exception:
            pass

        # Group by country
        country_cases = defaultdict(list)
        for r in rows:
            co = (r["country"] or "Unknown").strip()
            country_cases[co].append(r)

        # Find earliest case per country for label annotation
        earliest_per_country = {}
        for co, cases in country_cases.items():
            cases_sorted = sorted(cases, key=lambda r: self._extract_year(r["year_built"]))
            earliest_per_country[co] = cases_sorted[0]

        all_lons, all_lats, all_labels = [], [], []
        for r in rows:
            co_cn = (r["country"] or "Unknown").strip()
            co_en = cn_to_en.get(co_cn, co_cn)
            name_en = (r["name_en"] or "").strip()
            name_display = name_en if name_en else (r["name"] or "?")
            all_lons.append(r["longitude"])
            all_lats.append(r["latitude"])
            all_labels.append(f"{co_en}\n{name_display}")

        n_cases = len(rows)
        n_countries = len(country_cases)

        country_counts = {co: len(cases) for co, cases in country_cases.items()}

        # --- Figure ---
        fig = Figure(figsize=(10, 5.5), dpi=100)
        ax = fig.add_subplot(111, projection=ccrs.Robinson())
        ax.set_global()

        ax.add_feature(cfeature.LAND, facecolor="#d9dde0", edgecolor="none", zorder=0)
        ax.add_feature(cfeature.OCEAN, facecolor="#f5f8fa", zorder=0)
        ax.gridlines(draw_labels=False, linewidth=0.3, color="#aaaaaa", alpha=0.4, linestyle="--")
        ax.spines["geo"].set_visible(False)

        scatters = []
        if all_lons:
            s = ax.scatter(
                all_lons, all_lats,
                s=20, c="#f97316", alpha=0.35, edgecolors="none",
                transform=ccrs.PlateCarree(), zorder=5,
            )
            scatters.append(s)

        fig.subplots_adjust(bottom=0.01, top=0.99, left=0.01, right=0.99)

        # --- Create canvas BEFORE label measurement (need renderer) ---
        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.draw()

        # ---- Greedy country label placement ----
        pc = ccrs.PlateCarree()
        ann_trans = pc._as_mpl_transform(ax)

        annotations = []
        for co, case in earliest_per_country.items():
            co_en = cn_to_en.get(co, co)
            cnt = country_counts[co]
            texts = f"{co_en}: {cnt}"
            annotations.append((case["longitude"], case["latitude"], texts, co_en))

        pts = ax.projection.transform_points(pc, np.asarray(all_lons, dtype=float),
                                             np.asarray(all_lats, dtype=float))
        pts_disp = ax.transData.transform(pts[:, :2])
        pts_disp = pts_disp[np.isfinite(pts_disp).all(axis=1)]

        px_per_pt = fig.dpi / 72.0
        point_r = 5.0
        renderer = canvas.get_renderer()

        # Load saved positions
        import json as _json
        pos_file = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))),
                                 "data", "sd_map_label_positions.json")
        try:
            with open(pos_file, "r", encoding="utf-8") as f:
                saved = _json.load(f)
        except (OSError, ValueError):
            saved = {}

        dirs_8 = [
            (1, 0, "left", "center"), (-1, 0, "right", "center"),
            (0, 1, "center", "bottom"), (0, -1, "center", "top"),
            (0.71, 0.71, "left", "bottom"), (-0.71, 0.71, "right", "bottom"),
            (0.71, -0.71, "left", "top"), (-0.71, -0.71, "right", "top"),
        ]
        radii_7 = [9, 15, 22, 31, 42, 55, 70]

        def _rect_overlap(a, b):
            w = min(a[2], b[2]) - max(a[0], b[0])
            h = min(a[3], b[3]) - max(a[1], b[1])
            return w * h if (w > 0 and h > 0) else 0.0

        # Measure label boxes
        items = []
        for lon, lat, text, country in annotations:
            t = ax.text(0.5, 0.5, text, fontsize=5.5, fontweight="bold",
                        transform=ax.transAxes)
            tb = t.get_window_extent(renderer)
            t.remove()
            x, y = ax.projection.transform_point(lon, lat, pc)
            px, py = ax.transData.transform((x, y))
            items.append({"lon": lon, "lat": lat, "text": text, "country": country,
                          "w": tb.width + 8, "h": tb.height + 8,
                          "px": float(px), "py": float(py)})

        def _label_rect(it, ox, oy, ha, va):
            w, h = it["w"], it["h"]
            apx = it["px"] + ox * px_per_pt
            apy = it["py"] + oy * px_per_pt
            x0 = apx if ha == "left" else apx - w if ha == "right" else apx - w / 2
            y0 = apy if va == "bottom" else apy - h if va == "top" else apy - h / 2
            return (x0, y0, x0 + w, y0 + h)

        occupied = []
        for it in items:
            sv = saved.get(it["country"])
            if sv:
                it["pos"] = (float(sv["ox"]), float(sv["oy"]),
                             sv.get("ha", "left"), sv.get("va", "center"))
                occupied.append(_label_rect(it, *it["pos"]))

        for it in items:
            d = np.hypot(pts_disp[:, 0] - it["px"], pts_disp[:, 1] - it["py"])
            it["density"] = int((d < 150).sum())
        items.sort(key=lambda it: -it["density"])

        ax_box = ax.get_window_extent(renderer)

        for it in items:
            if "pos" in it:
                continue
            m = (np.abs(pts_disp[:, 0] - it["px"]) < 400) & \
                (np.abs(pts_disp[:, 1] - it["py"]) < 400)
            near = pts_disp[m]
            own = (np.abs(near[:, 0] - it["px"]) < 1) & (np.abs(near[:, 1] - it["py"]) < 1)
            near = near[~own]

            best = None
            found_clean = False
            for r in radii_7:
                for dx, dy, ha, va in dirs_8:
                    ox, oy = dx * r, dy * r
                    rect = _label_rect(it, ox, oy, ha, va)

                    cost = 0.0
                    if (rect[0] < ax_box.x0 or rect[2] > ax_box.x1
                            or rect[1] < ax_box.y0 or rect[3] > ax_box.y1):
                        cost += 1e6
                    for oc in occupied:
                        cost += 3.0 * _rect_overlap(rect, oc)
                    if len(near):
                        inx = (np.minimum(rect[2], near[:, 0] + point_r)
                               - np.maximum(rect[0], near[:, 0] - point_r))
                        iny = (np.minimum(rect[3], near[:, 1] + point_r)
                               - np.maximum(rect[1], near[:, 1] - point_r))
                        hit = (inx > 0) & (iny > 0)
                        cost += float((inx[hit] * iny[hit]).sum())

                    if cost == 0:
                        best = (r * 1.5, ox, oy, ha, va, rect)
                        found_clean = True
                        break
                    score = cost + r * 1.5
                    if best is None or score < best[0]:
                        best = (score, ox, oy, ha, va, rect)
                if found_clean:
                    break

            _, ox, oy, ha, va, rect = best
            occupied.append(rect)
            it["pos"] = (ox, oy, ha, va)

        made = []
        for it in items:
            ox, oy, ha, va = it["pos"]
            short = (ox * ox + oy * oy) ** 0.5 <= radii_7[0]

            ann = ax.annotate(
                it["text"],
                xy=(it["lon"], it["lat"]),
                xycoords=ann_trans,
                xytext=(ox, oy),
                textcoords="offset points",
                fontsize=5.5,
                fontweight="bold",
                color="#1e3a5f",
                ha=ha,
                va="center",
                bbox=dict(facecolor="white", alpha=0.5, edgecolor="none",
                          boxstyle="round,pad=0.25"),
                arrowprops=dict(
                    arrowstyle="-",
                    color="#888888",
                    lw=0.5,
                    connectionstyle="arc3,rad=0.15",
                ) if not short else dict(
                    arrowstyle="-",
                    color="#888888",
                    lw=0.5,
                ),
                zorder=10,
            )
            ann.draggable()
            made.append((it["country"], ann, ox, oy))

        def _save_sd_labels(event):
            data = {}
            for country, ann, iox, ioy in made:
                ox, oy = ann.xyann
                moved = abs(ox - iox) > 0.1 or abs(oy - ioy) > 0.1
                if moved or country in saved:
                    data[country] = {"ox": round(float(ox), 1), "oy": round(float(oy), 1),
                                     "ha": ann.get_ha(), "va": ann.get_va()}
            if data != saved:
                saved.clear()
                saved.update(data)
                try:
                    _os.makedirs(_os.path.dirname(pos_file), exist_ok=True)
                    with open(pos_file, "w", encoding="utf-8") as f:
                        _json.dump(saved, f, ensure_ascii=False, indent=2)
                except OSError:
                    pass

        fig.canvas.mpl_connect("button_release_event", _save_sd_labels)

        # ---- Save initial label positions ----
        save_data = {}
        for it in items:
            ox, oy, ha, va = it["pos"]
            save_data[it["country"]] = {"ox": round(float(ox), 1), "oy": round(float(oy), 1),
                                         "ha": ha, "va": va}
        try:
            _os.makedirs(_os.path.dirname(pos_file), exist_ok=True)
            with open(pos_file, "w", encoding="utf-8") as f:
                _json.dump(save_data, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

        # ---- Legend (draggable, position saved) ----
        from matplotlib.lines import Line2D
        _leg_file = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))),
                                  "data", "sd_map_legend_pos.json")
        _leg_saved = None
        try:
            with open(_leg_file, "r", encoding="utf-8") as f:
                _leg_saved = _json.load(f)
        except (OSError, ValueError):
            pass

        # Custom legend axes (separate from geo axes, always visible)
        _leg_default = [0.02, 0.02, 0.20, 0.16]  # [left, bottom, width, height] in figure fraction
        _leg_rect = _leg_saved if (_leg_saved and len(_leg_saved) == 4) else _leg_default
        ax_leg = fig.add_axes(_leg_rect, facecolor="white", alpha=0.9, zorder=100)
        ax_leg.set_xlim(0, 1)
        ax_leg.set_ylim(0, 1)
        ax_leg.axis("off")
        for spine in ax_leg.spines.values():
            spine.set_edgecolor("#555555")
            spine.set_linewidth(0.8)
            spine.set_visible(True)

        from matplotlib.patches import Rectangle
        # Case
        ax_leg.scatter([0.125], [0.82], s=40, c="#f97316", alpha=0.5, edgecolors="none", clip_on=False)
        ax_leg.text(0.28, 0.82, "Chinoiserie Architecture case", fontsize=7, va="center", ha="left")
        # Land
        ax_leg.add_patch(Rectangle((0.05, 0.48), 0.15, 0.12, facecolor="#d9dde0", edgecolor="none", clip_on=False))
        ax_leg.text(0.28, 0.54, "Land", fontsize=7, va="center", ha="left")
        # Ocean
        ax_leg.add_patch(Rectangle((0.05, 0.18), 0.15, 0.12, facecolor="#f5f8fa", edgecolor="#aaaaaa", linewidth=0.3, clip_on=False))
        ax_leg.text(0.28, 0.24, "Ocean", fontsize=7, va="center", ha="left")

        # Drag handling for legend axes
        _leg_drag = {"active": False, "dx": 0.0, "dy": 0.0}

        def _leg_fig_xy(event):
            return fig.transFigure.inverted().transform((event.x, event.y))

        def _leg_press(event):
            if event.inaxes is ax_leg:
                fx, fy = _leg_fig_xy(event)
                bb = ax_leg.get_position()
                _leg_drag.update(active=True, dx=fx - bb.x0, dy=fy - bb.y0)

        def _leg_motion(event):
            if not _leg_drag["active"]:
                return
            fx, fy = _leg_fig_xy(event)
            bb = ax_leg.get_position()
            new_x = max(0, min(1 - bb.width, fx - _leg_drag["dx"]))
            new_y = max(0, min(1 - bb.height, fy - _leg_drag["dy"]))
            ax_leg.set_position([new_x, new_y, bb.width, bb.height])
            fig.canvas.draw_idle()

        def _leg_release(event):
            if _leg_drag["active"]:
                _leg_drag["active"] = False
                bb = ax_leg.get_position()
                pos = [round(bb.x0, 4), round(bb.y0, 4), round(bb.width, 4), round(bb.height, 4)]
                try:
                    _os.makedirs(_os.path.dirname(_leg_file), exist_ok=True)
                    with open(_leg_file, "w", encoding="utf-8") as f:
                        _json.dump(pos, f)
                except OSError:
                    pass

        fig.canvas.mpl_connect("button_press_event", _leg_press)
        fig.canvas.mpl_connect("motion_notify_event", _leg_motion)
        fig.canvas.mpl_connect("button_release_event", _leg_release)

        canvas.draw()

        # Hover tooltips
        if HAS_CURSORS and scatters:
            cursor = mplcursors.cursor(scatters, hover=mplcursors.HoverMode.Transient)
            @cursor.connect("add")
            def _on_add(sel):
                idx = sel.index
                if idx < len(all_labels):
                    sel.annotation.set_text(all_labels[idx])
                    sel.annotation.set_fontsize(8)
                    sel.annotation.set_bbox(dict(
                        facecolor="white", alpha=0.9, edgecolor="#ccc",
                        boxstyle="round,pad=0.3"
                    ))

        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        self._sd_world_drawn = True

    def _draw_coastal_map(self, parent_frame):
        """Figure B: individual case points on a coastal-gradient basemap."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3
        import numpy as np

        try:
            import cartopy.crs as ccrs
            import cartopy.feature as cfeature
        except ImportError:
            ttk.Label(parent_frame, text="请安装 cartopy：pip install cartopy",
                     font=("", 12)).pack(expand=True)
            return

        # Landlocked countries for stats
        _landlocked = {
            "阿富汗", "安道尔", "亚美尼亚", "奥地利", "阿塞拜疆", "白俄罗斯",
            "不丹", "玻利维亚", "博茨瓦纳", "布基纳法索", "布隆迪", "中非",
            "乍得", "捷克", "斯威士兰", "埃塞俄比亚", "匈牙利", "哈萨克斯坦",
            "吉尔吉斯斯坦", "老挝", "莱索托", "列支敦士登", "卢森堡", "马拉维",
            "马里", "摩尔多瓦", "蒙古", "尼泊尔", "尼日尔", "北马其顿", "巴拉圭",
            "卢旺达", "圣马力诺", "塞尔维亚", "斯洛伐克", "南苏丹", "瑞士",
            "塔吉克斯坦", "土库曼斯坦", "乌干达", "乌兹别克斯坦", "梵蒂冈",
            "赞比亚", "津巴布韦",
        }

        def _is_landlocked(country_cn):
            for ll in _landlocked:
                if ll in country_cn or country_cn in ll:
                    return True
            return False

        conn = sqlite3.connect(db.DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(
            "SELECT id, country, latitude, longitude FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL "
            "AND name IS NOT NULL AND name != ''"
        )
        rows = cur.fetchall()
        conn.close()
        rows = _norm_coords(rows)

        if not rows:
            ttk.Label(parent_frame, text="暂无数据", font=("", 12)).pack(expand=True)
            return

        # --- Distance to nearest coastline → size classes ---
        # 距离数据直接读论文补充表 S2（数据库更新后运行 论文/regenerate_supplementary_tables.py 同步）
        import os as _osd
        import csv as _csvd
        dists = None
        _dist_csv = _osd.path.join(
            _osd.path.dirname(_osd.path.dirname(_osd.path.abspath(__file__))),
            "论文", "supplementary_table_coastal_distance.csv")
        dist_by_coord = {}
        try:
            with open(_dist_csv, encoding="utf-8-sig") as f:
                for row in _csvd.DictReader(f):
                    key = (round(float(row["Latitude"]), 6), round(float(row["Longitude"]), 6))
                    dist_by_coord[key] = float(row["Distance to nearest coastline (km)"])
        except Exception:
            dist_by_coord = {}
        if dist_by_coord:
            dists = [dist_by_coord.get((round(r["latitude"], 6), round(r["longitude"], 6)))
                     for r in rows]

        # 4 档: ≤50 最大+分离蓝环 / 50–100 / 100–500 / >500 内陆保持原大小
        _bands = [(50, 232), (100, 100), (500, 48), (float("inf"), 28)]

        def _band_of(d):
            if d is None:
                return len(_bands) - 1   # 表中未匹配到（新增案例未再生成表）按内陆样式
            for i, (th, _) in enumerate(_bands):
                if d <= th:
                    return i
            return len(_bands) - 1

        band_idx = None
        if dists is not None:
            band_idx = [_band_of(d) for d in dists]
            pt_sizes = [_bands[i][1] for i in band_idx]
            band_counts = [band_idx.count(i) for i in range(len(_bands))]
        else:
            pt_sizes, band_counts = 28, None

        # Stats
        country_coastal = {}
        for r in rows:
            co = (r["country"] or "Unknown").strip()
            if co not in country_coastal:
                country_coastal[co] = not _is_landlocked(co)

        n_coastal = sum(1 for v in country_coastal.values() if v)
        n_landlocked = sum(1 for v in country_coastal.values() if not v)
        coastal_cases = sum(1 for r in rows if country_coastal.get((r["country"] or "Unknown").strip(), True))
        landlocked_cases = len(rows) - coastal_cases
        total = len(rows)

        coastal_pct = coastal_cases / total * 100 if total > 0 else 0
        landlocked_pct = landlocked_cases / total * 100 if total > 0 else 0

        self.stats_label.config(
            text=f"沿海 {n_coastal} 国 / {coastal_cases} 例 ({coastal_pct:.1f}%)  |  "
                 f"内陆 {n_landlocked} 国 / {landlocked_cases} 例 ({landlocked_pct:.1f}%)"
        )

        all_lons = [r["longitude"] for r in rows]
        all_lats = [r["latitude"] for r in rows]

        # --- Toolbar ---
        toolbar = ttk.Frame(parent_frame)
        toolbar.pack(fill="x", padx=5, pady=(5, 0))
        self._brush_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(toolbar, text="画笔模式", variable=self._brush_var).pack(side="left", padx=5)
        ttk.Button(toolbar, text="撤销 (Ctrl+Z)", command=self._coastal_undo).pack(side="left", padx=3)
        ttk.Button(toolbar, text="清除全部连线", command=self._coastal_clear_all).pack(side="left", padx=3)
        ttk.Label(toolbar, text="开启画笔后拖拽画线  |  Ctrl+Z 撤销  |  关闭画笔时可拖拽图例（自动保存）",
                  foreground="#666").pack(side="right", padx=10)

        # --- Figure ---
        fig = Figure(figsize=(8, 4.5), dpi=100)
        ax = fig.add_subplot(111, projection=ccrs.Robinson())
        ax.set_global()

        ax.add_feature(cfeature.OCEAN, facecolor="#d8eaf5", zorder=0)

        import os as _os2
        _coast_file = _os2.path.join(
            _os2.path.expanduser("~"), ".local", "share", "cartopy",
            "shapefiles", "natural_earth", "physical", "ne_110m_coastline.shp")
        if _os2.path.exists(_coast_file):
            ax.add_feature(cfeature.LAND, facecolor="#ffffff", edgecolor="none", zorder=0)
            ocean_rgb = np.array([216, 234, 245]) / 255.0
            land_rgb = np.array([255, 255, 255]) / 255.0
            mid_rgb = (ocean_rgb + land_rgb) / 2
            for lw, alpha in [(24, 0.12), (14, 0.20), (8, 0.35), (4, 0.55), (1.5, 0.80)]:
                ax.add_feature(cfeature.COASTLINE, linewidth=lw,
                               edgecolor=mid_rgb, alpha=alpha, zorder=2)
        else:
            ax.add_feature(cfeature.LAND, facecolor="#ffffff",
                           edgecolor="#b0b8bf", linewidth=0.3, zorder=0)

        ax.gridlines(draw_labels=False, linewidth=0.3, color="#aaaaaa", alpha=0.3, linestyle="--")
        ax.spines["geo"].set_visible(False)

        # Blue dots sized by coastline-distance class
        _pt_scatter = None
        if all_lons:
            _pt_scatter = ax.scatter(all_lons, all_lats,
                                     s=pt_sizes, c="#4a85c4", alpha=0.65,
                                     edgecolors="none",
                                     transform=ccrs.PlateCarree(), zorder=5)
            # 外环圈数区分档位: ≤50 三圈 / 50–100 两圈 / 100–500 一圈 / >500 无
            if band_idx is not None:
                _ring_sizes = {0: (420, 660, 960), 1: (200, 330), 2: (110,)}
                for b, sizes_b in _ring_sizes.items():
                    blons = [all_lons[i] for i, bi in enumerate(band_idx) if bi == b]
                    blats = [all_lats[i] for i, bi in enumerate(band_idx) if bi == b]
                    if not blons:
                        continue
                    for rs in sizes_b:
                        ax.scatter(blons, blats, s=rs, facecolors="none",
                                   edgecolors="#1f5fa8", linewidths=0.7, alpha=0.8,
                                   transform=ccrs.PlateCarree(), zorder=6)

        # Highlight scatters
        hover_hl = ax.scatter([], [], s=180, facecolors="none",
                              edgecolors="#ff4444", linewidths=2.5,
                              transform=ccrs.PlateCarree(), zorder=100)
        select_hl = ax.scatter([], [], s=220, facecolors="none",
                               edgecolors="#cc0000", linewidths=3,
                               transform=ccrs.PlateCarree(), zorder=101)

        # --- Canvas ---
        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # --- State ---
        state = {
            "rows": rows,
            "all_lons": all_lons,
            "all_lats": all_lats,
            "selected_idx": None,
            "hovered_idx": None,
            "hover_hl": hover_hl,
            "select_hl": select_hl,
            "line_artists": [],
            "canvas": canvas,
            "ax": ax,
            "fig": fig,
            "ccrs": ccrs,
            "pt_scatter": _pt_scatter,
        }
        self._coastal_state = state
        self._coastal_canvas = canvas

        # Draw saved lines
        self._redraw_coastal_lines(ax, ccrs, state)

        # --- Hover: red circle highlight ---
        try:
            import mplcursors
            if _pt_scatter is not None:
                _cursor = mplcursors.cursor([_pt_scatter], hover=mplcursors.HoverMode.Transient)

                @_cursor.connect("add")
                def _on_add(sel):
                    idx = sel.index
                    sel.annotation.set_visible(False)
                    hover_hl.set_offsets([[all_lons[idx], all_lats[idx]]])
                    hover_hl.set_visible(True)
                    canvas.draw_idle()

                @_cursor.connect("remove")
                def _on_remove(sel):
                    hover_hl.set_visible(False)
                    canvas.draw_idle()
        except ImportError:
            pass

        # --- Freehand drawing ---
        _drawing = {"active": False, "points": [], "line": None}

        def _to_lonlat(x, y):
            """Convert data coords (Robinson) to lon/lat."""
            import cartopy.crs as _ccrs
            return _ccrs.PlateCarree().transform_point(x, y, ax.projection)

        def _on_press(event):
            if not self._brush_var.get():
                return
            if event.inaxes != ax:
                return
            if event.button != 1:
                return
            _drawing["active"] = True
            lon, lat = _to_lonlat(event.xdata, event.ydata)
            _drawing["points"] = [[lon, lat]]

        def _on_motion(event):
            if not _drawing["active"]:
                return
            if event.inaxes != ax or event.xdata is None:
                return
            lon, lat = _to_lonlat(event.xdata, event.ydata)
            _drawing["points"].append([lon, lat])

        def _on_release(event):
            if not _drawing["active"]:
                return
            _drawing["active"] = False
            pts = _drawing["points"]
            if len(pts) < 2:
                _drawing["points"] = []
                return

            # Smooth and convert to double line
            smoothed = self._smooth_polyline(pts)
            left, right = self._make_double_line(smoothed)

            # Draw two parallel curves
            for curve in [left, right]:
                clons = [p[0] for p in curve]
                clats = [p[1] for p in curve]
                line = ax.plot(clons, clats, color="#4a85c4", linewidth=5, alpha=0.30,
                               transform=ccrs.PlateCarree(), zorder=3)
                state["line_artists"].append(line[0])

            # Store stroke
            self._coastal_lines.append(pts)
            self._save_coastal_lines()
            _drawing["points"] = []
            canvas.draw_idle()

        canvas.mpl_connect("button_press_event", _on_press)
        canvas.mpl_connect("motion_notify_event", _on_motion)
        canvas.mpl_connect("button_release_event", _on_release)

        # Ctrl+Z / Cmd+Z binding
        def _on_ctrl_z(event):
            self._coastal_undo()

        canvas.get_tk_widget().bind("<Control-z>", _on_ctrl_z)
        canvas.get_tk_widget().bind("<Command-z>", _on_ctrl_z)

        # --- Legend: 自绘面板（无背景无边框，可拖拽，位置记忆） ---
        import json as _jsonl
        import os as _osl
        from matplotlib.patches import Rectangle

        _leg_pos_file = _osl.path.join(_osl.path.dirname(db.DB_PATH),
                                       "coastal_legend_pos.json")
        _leg_pos = [0.015, 0.05, 0.20, 0.42]
        try:
            with open(_leg_pos_file) as f:
                p = _jsonl.load(f)
                if isinstance(p, list) and len(p) == 4:
                    _leg_pos = p
        except Exception:
            pass

        ax_leg = fig.add_axes(_leg_pos)
        ax_leg.set_xlim(0, 1)
        ax_leg.set_ylim(0, 1)
        ax_leg.axis("off")
        ax_leg.patch.set_alpha(0)

        _icon_x, _text_x = 0.13, 0.26
        _fs = 7

        def _leg_row_patch(y, fc, ec, lw, label):
            ax_leg.add_patch(Rectangle((_icon_x - 0.09, y - 0.028), 0.18, 0.056,
                                       facecolor=fc, edgecolor=ec, linewidth=lw,
                                       zorder=3, clip_on=False))
            ax_leg.text(_text_x, y, label, fontsize=_fs, va="center", ha="left")

        _leg_row_patch(0.97, "#ffffff", "#b0b8bf", 0.3, "陆地")
        _leg_row_patch(0.89, "#d8eaf5", "none", 0, "海洋")

        if band_counts:
            _band_names = ["≤50 km", "50–100 km", "100–500 km", ">500 km"]
            _leg_ring_sizes = {0: (420, 660, 960), 1: (200, 330), 2: (110,), 3: ()}
            _row_y = [0.74, 0.56, 0.43, 0.34]
            for i, name in enumerate(_band_names):
                y = _row_y[i]
                ax_leg.scatter([_icon_x], [y], s=_bands[i][1], c="#4a85c4",
                               alpha=0.65, edgecolors="none", zorder=4, clip_on=False)
                for rs in _leg_ring_sizes[i]:
                    ax_leg.scatter([_icon_x], [y], s=rs, facecolors="none",
                                   edgecolors="#1f5fa8", linewidths=0.7, alpha=0.8,
                                   zorder=4, clip_on=False)
                ax_leg.text(_text_x, y, f"距海岸线 {name}（{band_counts[i]}例）",
                            fontsize=_fs, va="center", ha="left")
        else:
            ax_leg.scatter([_icon_x], [0.74], s=28, c="#4a85c4", alpha=0.65,
                           edgecolors="none", zorder=4, clip_on=False)
            ax_leg.text(_text_x, 0.74, "案例点", fontsize=_fs, va="center", ha="left")

        ax_leg.plot([_icon_x - 0.09, _icon_x + 0.09], [0.22, 0.22],
                    color="#4a85c4", linewidth=4, alpha=0.30, clip_on=False)
        ax_leg.plot([_icon_x - 0.09, _icon_x + 0.09], [0.19, 0.19],
                    color="#4a85c4", linewidth=4, alpha=0.30, clip_on=False)
        ax_leg.text(_text_x, 0.205, "沿海线", fontsize=_fs, va="center", ha="left")

        # 拖拽图例（画笔模式关闭时生效）
        _leg_drag = {"active": False, "dx": 0.0, "dy": 0.0}

        def _leg_fig_xy(event):
            return fig.transFigure.inverted().transform((event.x, event.y))

        def _leg_press(event):
            if self._brush_var.get():
                return
            if event.inaxes is ax_leg:
                fx, fy = _leg_fig_xy(event)
                bb = ax_leg.get_position()
                _leg_drag.update(active=True, dx=fx - bb.x0, dy=fy - bb.y0)

        def _leg_motion(event):
            if _leg_drag["active"]:
                fx, fy = _leg_fig_xy(event)
                bb = ax_leg.get_position()
                ax_leg.set_position([fx - _leg_drag["dx"], fy - _leg_drag["dy"],
                                     bb.width, bb.height])
                canvas.draw_idle()

        def _leg_release(event):
            if _leg_drag["active"]:
                _leg_drag["active"] = False
                bb = ax_leg.get_position()
                try:
                    with open(_leg_pos_file, "w") as f:
                        _jsonl.dump([bb.x0, bb.y0, bb.width, bb.height], f)
                except OSError:
                    pass

        canvas.mpl_connect("button_press_event", _leg_press)
        canvas.mpl_connect("motion_notify_event", _leg_motion)
        canvas.mpl_connect("button_release_event", _leg_release)

        # --- Final ---
        ax.set_title("海外中式建筑空间分布：海陆格局", fontsize=14, fontweight="bold", pad=10)
        fig.subplots_adjust(bottom=0.01, top=0.97, left=0.01, right=0.99)
        canvas.draw()
        self._coastal_drawn = True

    def _draw_continent_map(self, parent_frame):
        """五大洲分析: continent choropleth (orange) + draggable Asia-subregion inset (beige)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3
        import re as _re
        import os as _os
        import json as _json
        import numpy as np
        from collections import defaultdict
        from matplotlib.colors import LinearSegmentedColormap

        try:
            import cartopy.crs as ccrs
            import cartopy.feature as cfeature
            from cartopy.io import shapereader
        except ImportError:
            ttk.Label(parent_frame, text="请安装 cartopy：pip install cartopy",
                      font=("", 12)).pack(expand=True)
            return

        shp = _os.path.expanduser(
            "~/.local/share/cartopy/shapefiles/natural_earth/cultural/"
            "ne_110m_admin_0_countries.shp")
        if not _os.path.exists(shp):
            ttk.Label(parent_frame, text="缺少国家边界数据 ne_110m_admin_0_countries.shp",
                      font=("", 12)).pack(expand=True)
            return

        conn = sqlite3.connect(db.DB_PATH)
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT country FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL "
            "AND name IS NOT NULL AND name != ''").fetchall()
        conn.close()
        if not rows:
            ttk.Label(parent_frame, text="暂无数据", font=("", 12)).pack(expand=True)
            return

        # --- 案例数按 M49 五大区域 + 亚洲细分统计（伊朗归中亚、俄罗斯归东欧，论文口径） ---
        _zw = _re.compile(r'[​‌‍‎‏﻿]+')
        # DB country 字段为英文 M49 短名，直接查英文映射；
        # 亚区拼写统一为 Natural Earth 的 "South-Eastern Asia"（与下方 subregion_of 一致）
        region_of = {
            en: (reg, sub.replace("South-eastern", "South-Eastern"))
            for en, (reg, sub) in country_mapping.M49_REGION.items()
        }

        reg_n, sub_n = defaultdict(int), defaultdict(int)
        for r in rows:
            reg, sub = region_of.get(_zw.sub('', r["country"] or ''), ("", ""))
            if reg:
                reg_n[reg] += 1
            if sub:
                sub_n[sub] += 1
        total = len(rows)

        _cn = {"Asia": "亚洲", "Europe": "欧洲", "Americas": "美洲",
               "Oceania": "大洋洲", "Africa": "非洲"}
        _order = ["Asia", "Europe", "Americas", "Oceania", "Africa"]
        self.stats_label.config(text="  |  ".join(
            f"{_cn[k]} {reg_n[k]} 例 ({reg_n[k]/total*100:.1f}%)" for k in _order))

        # --- Toolbar ---
        toolbar = ttk.Frame(parent_frame)
        toolbar.pack(fill="x", padx=5, pady=(5, 0))
        ttk.Label(toolbar, text="拖拽亚洲小图 / 文字标注可调整位置（自动保存）",
                  foreground="#666").pack(side="right", padx=10)

        # --- Colors ---
        oranges = LinearSegmentedColormap.from_list("or", ["#fde5c8", "#f5a75b", "#d95f0e"])
        beiges = LinearSegmentedColormap.from_list("be", ["#fdf9ec", "#e6d49a", "#c2a04a"])
        reg_max = max(reg_n.values()) if reg_n else 1
        sub_max = max(sub_n.values()) if sub_n else 1

        def shade(cmap, val, vmax):
            return cmap(np.sqrt(val / vmax)) if vmax else "#f2f2f2"

        recs = list(shapereader.Reader(shp).records())

        def is_china(rec):
            # 研究范围为中国以外，中国本土（含台湾地区）不填色
            return rec.attributes.get("ADMIN") in ("China", "Taiwan")

        def continent_of(rec):
            c = rec.attributes.get("CONTINENT")
            return "Americas" if c in ("North America", "South America") else c

        def subregion_of(rec):
            if rec.attributes.get("ADMIN") == "Iran":
                return "Central Asia"   # 论文口径
            return rec.attributes.get("SUBREGION")

        # --- Main map ---
        fig = Figure(figsize=(8, 4.5), dpi=100)
        ax = fig.add_subplot(111, projection=ccrs.Robinson())
        ax.set_global()
        ax.add_feature(cfeature.OCEAN, facecolor="#ffffff", zorder=0)
        ax.spines["geo"].set_visible(False)

        for rec in recs:
            cont = continent_of(rec)
            if is_china(rec):
                ax.add_geometries([rec.geometry], ccrs.PlateCarree(),
                                  facecolor="#ffffff", edgecolor="#c9c9c9",
                                  linewidth=0.4, zorder=2)
                continue
            fc = shade(oranges, reg_n[cont], reg_max) if cont in reg_n else "#f2f2f2"
            ax.add_geometries([rec.geometry], ccrs.PlateCarree(),
                              facecolor=fc, edgecolor="white", linewidth=0.3, zorder=2)

        # --- Labels (draggable, positions persisted) ---
        lbl_file = _os.path.join(_os.path.dirname(db.DB_PATH), "continent_label_pos.json")
        saved_lbl = {}
        try:
            with open(lbl_file) as f:
                saved_lbl = _json.load(f)
        except Exception:
            pass

        label_artists = {}
        _label_pos = {"Asia": (95, 48), "Europe": (12, 54), "Americas": (-100, 42),
                      "Oceania": (134, -24), "Africa": (17, 8)}
        for k in _order:
            lon, lat = saved_lbl.get(k, _label_pos[k])
            t = ax.text(lon, lat, f"{k}\n{reg_n[k]} ({reg_n[k]/total*100:.1f}%)",
                        transform=ccrs.PlateCarree(), fontsize=6.5, fontweight="bold",
                        ha="center", va="center", color="#333333", zorder=6)
            label_artists[k] = (t, "main")

        # --- Legend: colorbars + patch legend ---
        from matplotlib.cm import ScalarMappable
        from matplotlib.colors import PowerNorm
        from matplotlib.patches import Patch
        _ticks1 = sorted({0, 50, 150, 300, reg_max} - {t for t in (50, 150, 300) if t > reg_max})
        cax1 = fig.add_axes([0.71, 0.15, 0.19, 0.022])
        cb1 = fig.colorbar(ScalarMappable(norm=PowerNorm(0.5, vmin=0, vmax=reg_max),
                                          cmap=oranges),
                           cax=cax1, orientation="horizontal")
        cb1.set_label("Cases per continent", fontsize=5.5)
        cb1.set_ticks(_ticks1)
        cb1.ax.tick_params(labelsize=5, length=2)
        cb1.outline.set_linewidth(0.4)
        _ticks2 = sorted({0, 25, 100, 200, sub_max} - {t for t in (25, 100, 200) if t > sub_max})
        cax2 = fig.add_axes([0.71, 0.06, 0.19, 0.022])
        cb2 = fig.colorbar(ScalarMappable(norm=PowerNorm(0.5, vmin=0, vmax=sub_max),
                                          cmap=beiges),
                           cax=cax2, orientation="horizontal")
        cb2.set_label("Cases per Asian subregion", fontsize=5.5)
        cb2.set_ticks(_ticks2)
        cb2.ax.tick_params(labelsize=5, length=2)
        cb2.outline.set_linewidth(0.4)
        ax.legend(handles=[Patch(facecolor="#f2f2f2", edgecolor="#dddddd",
                                 label="No documented cases"),
                           Patch(facecolor="#ffffff", edgecolor="#c9c9c9",
                                 label="China (outside study scope)")],
                  loc="upper left", fontsize=5.5, frameon=False)

        # --- Asia inset (draggable, position persisted) ---
        pos_file = _os.path.join(_os.path.dirname(db.DB_PATH), "continent_inset_pos.json")
        inset_pos = [0.015, 0.06, 0.30, 0.42]
        try:
            with open(pos_file) as f:
                p = _json.load(f)
                if isinstance(p, list) and len(p) == 4:
                    inset_pos = p
        except Exception:
            pass

        axi = fig.add_axes(inset_pos, projection=ccrs.PlateCarree())
        axi.set_extent([25, 150, -12, 58], crs=ccrs.PlateCarree())
        axi.add_feature(cfeature.OCEAN, facecolor="#ffffff", zorder=0)
        for rec in recs:
            if rec.attributes.get("CONTINENT") != "Asia":
                axi.add_geometries([rec.geometry], ccrs.PlateCarree(),
                                   facecolor="#f2f2f2", edgecolor="white",
                                   linewidth=0.2, zorder=1)
                continue
            if is_china(rec):
                axi.add_geometries([rec.geometry], ccrs.PlateCarree(),
                                   facecolor="#ffffff", edgecolor="#c9c9c9",
                                   linewidth=0.3, zorder=2)
                continue
            sub = subregion_of(rec)
            axi.add_geometries([rec.geometry], ccrs.PlateCarree(),
                               facecolor=shade(beiges, sub_n.get(sub, 0), sub_max),
                               edgecolor="white", linewidth=0.25, zorder=2)
        _sub_labels = [("South-Eastern Asia", "SE Asia", 118, -6),
                       ("Eastern Asia", "E Asia", 122, 42),
                       ("Southern Asia", "S Asia", 76, 14),
                       ("Central Asia", "C Asia", 64, 44),
                       ("Western Asia", "W Asia", 40, 32)]
        for key, short, lon, lat in _sub_labels:
            n = sub_n.get(key, 0)
            lkey = f"sub:{short}"
            plon, plat = saved_lbl.get(lkey, (lon, lat))
            t = axi.text(plon, plat, f"{short} {n} ({n/total*100:.1f}%)",
                         transform=ccrs.PlateCarree(), fontsize=5, fontweight="bold",
                         ha="center", va="center", color="#444444", zorder=6)
            label_artists[lkey] = (t, "inset")
        axi.set_title("Asia by subregion", fontsize=6.5)
        for s in axi.spines.values():
            s.set_visible(False)

        ax.set_title("海外中式建筑五大洲分布", fontsize=13, fontweight="bold", pad=8)
        fig.subplots_adjust(bottom=0.01, top=0.95, left=0.01, right=0.99)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # --- Drag handling: text labels + inset ---
        drag = {"mode": None, "key": None, "dx": 0.0, "dy": 0.0}

        def _fig_xy(event):
            return fig.transFigure.inverted().transform((event.x, event.y))

        def _on_press(event):
            for key, (txt, _which) in label_artists.items():
                hit, _ = txt.contains(event)
                if hit:
                    drag.update(mode="text", key=key)
                    return
            if event.inaxes is axi:
                fx, fy = _fig_xy(event)
                bb = axi.get_position()
                drag.update(mode="inset", dx=fx - bb.x0, dy=fy - bb.y0)

        def _on_motion(event):
            if drag["mode"] == "text":
                txt, which = label_artists[drag["key"]]
                tgt = ax if which == "main" else axi
                try:
                    x, y = tgt.transData.inverted().transform((event.x, event.y))
                    lon, lat = ccrs.PlateCarree().transform_point(x, y, tgt.projection)
                    if np.isfinite(lon) and np.isfinite(lat):
                        txt.set_position((lon, lat))
                        canvas.draw_idle()
                except Exception:
                    pass
            elif drag["mode"] == "inset":
                fx, fy = _fig_xy(event)
                bb = axi.get_position()
                axi.set_position([fx - drag["dx"], fy - drag["dy"], bb.width, bb.height])
                canvas.draw_idle()

        def _on_release(event):
            if drag["mode"] == "inset":
                bb = axi.get_position()
                try:
                    with open(pos_file, "w") as f:
                        _json.dump([bb.x0, bb.y0, bb.width, bb.height], f)
                except OSError:
                    pass
            elif drag["mode"] == "text":
                try:
                    with open(lbl_file, "w") as f:
                        _json.dump({k: list(t.get_position())
                                    for k, (t, _) in label_artists.items()}, f)
                except OSError:
                    pass
            drag["mode"] = None

        canvas.mpl_connect("button_press_event", _on_press)
        canvas.mpl_connect("motion_notify_event", _on_motion)
        canvas.mpl_connect("button_release_event", _on_release)

        canvas.draw()
        self._continent_drawn = True

    def _draw_country_choropleth(self, parent_frame):
        """案例国家分布图: country-level choropleth shaded by case count."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3
        import os as _os
        import numpy as np
        from collections import defaultdict
        from matplotlib.colors import LinearSegmentedColormap

        try:
            import cartopy.crs as ccrs
            import cartopy.feature as cfeature
            from cartopy.io import shapereader
        except ImportError:
            ttk.Label(parent_frame, text="请安装 cartopy：pip install cartopy",
                      font=("", 12)).pack(expand=True)
            return

        shp = _os.path.expanduser(
            "~/.local/share/cartopy/shapefiles/natural_earth/cultural/"
            "ne_110m_admin_0_countries.shp")
        if not _os.path.exists(shp):
            ttk.Label(parent_frame, text="缺少国家边界数据 ne_110m_admin_0_countries.shp",
                      font=("", 12)).pack(expand=True)
            return

        conn = sqlite3.connect(db.DB_PATH)
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT country FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL "
            "AND name IS NOT NULL AND name != ''").fetchall()
        conn.close()

        # Hardcoded CN→EN country mapping (68 countries in database)
        CN_TO_EN = {
            "泰国": "Thailand", "日本": "Japan", "马来西亚": "Malaysia",
            "越南": "Vietnam", "韩国": "South Korea", "美国": "United States of America",
            "新加坡": "Singapore", "印度尼西亚": "Indonesia", "德国": "Germany",
            "澳大利亚": "Australia", "俄罗斯": "Russia", "法国": "France",
            "加拿大": "Canada", "英国": "United Kingdom", "朝鲜": "North Korea",
            "蒙古": "Mongolia", "菲律宾": "Philippines", "荷兰": "Netherlands",
            "瑞典": "Sweden", "意大利": "Italy", "柬埔寨": "Cambodia",
            "丹麦": "Denmark", "新西兰": "New Zealand", "缅甸": "Myanmar",
            "印度": "India", "乌克兰": "Ukraine", "波兰": "Poland",
            "老挝": "Laos", "毛里求斯": "Mauritius", "不丹": "Bhutan",
            "古巴": "Cuba", "吉尔吉斯斯坦": "Kyrgyzstan", "奥地利": "Austria",
            "比利时": "Belgium", "西班牙": "Spain", "白俄罗斯": "Belarus",
            "伊朗": "Iran", "匈牙利": "Hungary", "南非": "South Africa",
            "哈萨克斯坦": "Kazakhstan", "墨西哥": "Mexico", "尼日利亚": "Nigeria",
            "巴西": "Brazil", "捷克": "Czech Republic", "津巴布韦": "Zimbabwe",
            "赞比亚": "Zambia", "瑞士": "Switzerland", "东帝汶": "Timor-Leste",
            "博茨瓦纳": "Botswana", "厄瓜多尔": "Ecuador", "土耳其": "Turkey",
            "圭亚那": "Guyana", "基里巴斯": "Kiribati", "塞舌尔": "Seychelles",
            "密克罗尼西亚联邦": "Micronesia", "尼泊尔": "Nepal", "巴拿马": "Panama",
            "拉脱维亚": "Latvia", "摩尔多瓦共和国": "Moldova", "文莱": "Brunei",
            "智利": "Chile", "爱尔兰": "Ireland", "秘鲁": "Peru",
            "纳米比亚": "Namibia", "葡萄牙": "Portugal", "贝宁": "Benin",
            "阿根廷": "Argentina", "马拉维": "Malawi", "马达加斯加": "Madagascar",
        }

        # Count cases per country (CN→EN)
        country_count = defaultdict(int)
        for r in rows:
            cn = (r["country"] or "").strip()
            en = CN_TO_EN.get(cn, cn)
            country_count[en] += 1

        # Shapefile name variants for matching
        SHAPEFILE_VARIANTS = {
            "South Korea": ["Republic of Korea", "Korea, Republic of", "South Korea"],
            "North Korea": ["Democratic People's Republic of Korea",
                           "Dem. Rep. Korea", "North Korea"],
            "Russia": ["Russian Federation", "Russia"],
            "United States of America": ["United States", "United States of America"],
            "Vietnam": ["Viet Nam", "Vietnam"],
            "Laos": ["Lao People's Democratic Republic", "Lao PDR", "Laos"],
            "Brunei": ["Brunei Darussalam", "Brunei"],
            "Myanmar": ["Myanmar", "Burma"],
            "Timor-Leste": ["Timor-Leste", "East Timor"],
            "Micronesia": ["Micronesia", "Federated States of Micronesia"],
            "Moldova": ["Moldova", "Republic of Moldova"],
            "Czech Republic": ["Czech Republic", "Czechia"],
            "Turkey": ["Turkey", "Türkiye"],
            "United Kingdom": ["United Kingdom"],
            "Kyrgyzstan": ["Kyrgyzstan", "Kyrgyz Republic"],
            "Iran": ["Iran", "Islamic Republic of Iran"],
        }

        reader = shapereader.Reader(shp)
        name_index = {}
        for rec in reader.records():
            for f in ["NAME", "NAME_LONG", "ADMIN", "SOVEREIGNT", "NAME_SORT"]:
                v = rec.attributes.get(f, "")
                if v:
                    name_index[v] = rec

        def _match(en_name):
            if en_name in name_index:
                return en_name
            candidates = SHAPEFILE_VARIANTS.get(en_name, [en_name])
            for c in candidates:
                if c in name_index:
                    return c
            for k in name_index:
                if en_name.lower() == k.lower():
                    return k
            return None

        # Bin definitions (10 levels)
        bins = [0, 1, 2, 3, 5, 8, 13, 20, 35, 60, 9999]
        bin_labels = ["1", "2", "3-4", "5-7", "8-12", "13-19", "20-34", "35-59", "60-79", "80+"]
        n_bins = len(bins) - 1
        blues = LinearSegmentedColormap.from_list("bl", ["#d4e6f1", "#2471a3", "#1a5276"])

        fig = Figure(figsize=(14, 7), dpi=120)
        ax = fig.add_subplot(111, projection=ccrs.Robinson())
        ax.set_global()
        ax.set_frame_on(False)
        ax.add_feature(cfeature.OCEAN, facecolor="#f0f4f8", zorder=0)

        matched_cns = set()
        labels_data = []
        for rec in reader.records():
            geom = rec.geometry
            name = rec.attributes.get("NAME", "")
            count = 0
            for en_name, cnt in country_count.items():
                k = _match(en_name)
                if k == name:
                    count = cnt
                    matched_cns.add(en_name)
                    break

            if count == 0:
                ax.add_geometries([geom], ccrs.PlateCarree(), facecolor="#e8e8e8",
                                  edgecolor="#cccccc", linewidth=0.3, zorder=1)
            else:
                bin_idx = 0
                for j in range(n_bins):
                    if bins[j] < count <= bins[j+1]:
                        bin_idx = j
                        break
                t = bin_idx / max(n_bins - 1, 1)
                color = blues(t)
                ax.add_geometries([geom], ccrs.PlateCarree(), facecolor=color,
                                  edgecolor="#ffffff", linewidth=0.4, zorder=2)
                pt = geom.representative_point()
                labels_data.append((pt.x, pt.y, name, count))

        # Unmatched diagnostics
        unmatched = set(country_count.keys()) - matched_cns
        if unmatched:
            print(f"[Choropleth] Unmatched countries: {unmatched}")

        # Draggable labels with leader lines — load saved positions
        import matplotlib.patheffects as pe
        import json as _json

        _label_pos_file = _os.path.join(_os.path.dirname(db.DB_PATH),
                                        "_choro_label_positions.json")
        _saved_positions = {}
        if _os.path.exists(_label_pos_file):
            try:
                with open(_label_pos_file, "r", encoding="utf-8") as _f:
                    _saved_positions = _json.load(_f)
            except Exception:
                pass

        label_artists = []  # (text, line, anchor_lon, anchor_lat)

        for lon, lat, name, cnt in labels_data:
            slon, slat = _saved_positions.get(name, (lon, lat))
            t = ax.text(slon, slat, name, transform=ccrs.PlateCarree(),
                        fontsize=6, ha="center", va="center", color="#1a1a2e",
                        path_effects=[pe.withStroke(linewidth=1.5, foreground="white")],
                        zorder=4, picker=True)
            line = ax.plot([lon, slon], [lat, slat], transform=ccrs.PlateCarree(),
                           color="#888888", linewidth=0.4, zorder=3)[0]
            label_artists.append((t, line, lon, lat))

        # Drag handling — convert mouse (Robinson metres) → PlateCarree lon/lat
        _drag_state = {"artist": None}

        def _on_press(event):
            if event.inaxes != ax:
                return
            for t, line, clon, clat in label_artists:
                contains, _ = t.contains(event)
                if contains:
                    _drag_state["artist"] = (t, line, clon, clat)
                    _drag_state["text_start"] = t.get_position()
                    _drag_state["mouse_start"] = ccrs.PlateCarree().transform_point(
                        event.xdata, event.ydata, ax.projection)
                    return

        def _on_motion(event):
            if _drag_state["artist"] is None or event.inaxes != ax:
                return
            t, line, clon, clat = _drag_state["artist"]
            cur_ll = ccrs.PlateCarree().transform_point(
                event.xdata, event.ydata, ax.projection)
            dlon = cur_ll[0] - _drag_state["mouse_start"][0]
            dlat = cur_ll[1] - _drag_state["mouse_start"][1]
            new_lon = _drag_state["text_start"][0] + dlon
            new_lat = _drag_state["text_start"][1] + dlat
            t.set_position((new_lon, new_lat))
            line.set_data([clon, new_lon], [clat, new_lat])
            canvas.draw_idle()

        def _on_release(event):
            if _drag_state["artist"] is not None:
                # Save all positions
                pos = {}
                for tt, _, _, _ in label_artists:
                    p = tt.get_position()
                    pos[tt.get_text()] = [float(p[0]), float(p[1])]
                try:
                    with open(_label_pos_file, "w", encoding="utf-8") as _f:
                        _json.dump(pos, _f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
            _drag_state["artist"] = None

        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.mpl_connect("button_press_event", _on_press)
        canvas.mpl_connect("motion_notify_event", _on_motion)
        canvas.mpl_connect("button_release_event", _on_release)

        # Draggable legend
        from matplotlib.patches import Patch
        legend_patches = []
        for j in range(n_bins):
            t_val = j / max(n_bins - 1, 1)
            legend_patches.append(Patch(facecolor=blues(t_val), edgecolor="#ffffff",
                                         label=bin_labels[j]))
        legend = ax.legend(handles=legend_patches, loc="upper right", fontsize=7,
                           title="Cases per country", title_fontsize=8,
                           frameon=False, ncol=1, bbox_to_anchor=(0.98, 0.96))
        legend.set_draggable(True)

        # ---- Inset: Long-tail bar chart (ranked countries) ----
        _inset_pos_file = _os.path.join(_os.path.dirname(db.DB_PATH),
                                        "_choro_inset_pos.json")
        _saved_inset = None
        if _os.path.exists(_inset_pos_file):
            try:
                with open(_inset_pos_file, "r", encoding="utf-8") as _f:
                    _saved_inset = _json.load(_f)
            except Exception:
                pass
        inset_rect = _saved_inset if _saved_inset else [0.04, 0.06, 0.38, 0.20]
        ax_inset = fig.add_axes(inset_rect, facecolor="none",
                                zorder=10)
        ax_inset.set_zorder(10)

        # Plot ranked bars — sort by (count desc, name asc) for stable unique labels
        sorted_countries = sorted(country_count.items(), key=lambda x: (-x[1], x[0]))
        sorted_counts = [c for _, c in sorted_countries]
        ranks = list(range(1, len(sorted_counts) + 1))
        bar_colors = [blues(min(c / max(sorted_counts), 1.0) * 0.8 + 0.1)
                      for c in sorted_counts]
        ax_inset.bar(ranks, sorted_counts, width=0.85, color=bar_colors,
                     edgecolor="white", linewidth=0.2)
        ax_inset.set_xlabel("Country rank", fontsize=7, color="#444444")
        ax_inset.set_ylabel("Cases", fontsize=7, color="#444444")
        ax_inset.tick_params(labelsize=6, colors="#666666")
        ax_inset.spines["top"].set_visible(False)
        ax_inset.spines["right"].set_visible(False)
        ax_inset.spines["left"].set_color("#cccccc")
        ax_inset.spines["bottom"].set_color("#cccccc")

        # Axis labels
        _xlabel_ann = ax_inset.annotate("Country rank",
                                        xy=(0.5, -0.08), xycoords="axes fraction",
                                        fontsize=7, color="#444444",
                                        ha="center", va="top",
                                        annotation_clip=False)
        _ylabel_ann = ax_inset.annotate("Cases",
                                        xy=(-0.04, 0.5), xycoords="axes fraction",
                                        fontsize=7, color="#444444", rotation=90,
                                        ha="center", va="bottom",
                                        annotation_clip=False)

        # Country labels as draggable annotations
        inset_anns = []
        used_positions = {}
        for i in range(len(sorted_countries)):
            name, c = sorted_countries[i]
            if len(name) > 14:
                name = name[:13] + "."
            label = f"{name} ({c})"
            # Offset to avoid overlap at same Y
            x_jitter = 0
            y_offset = 12  # points above bar
            key = (ranks[i], c)
            if key in used_positions:
                used_positions[key] += 1
                x_jitter = used_positions[key] * 16
            else:
                used_positions[key] = 0
            ann = ax_inset.annotate(label,
                                    xy=(ranks[i], c),
                                    xytext=(ranks[i] + x_jitter * 0.04,
                                            c + y_offset * 0.05),
                                    fontsize=4.5, color="#1a5276",
                                    rotation=90, ha="center", va="bottom",
                                    annotation_clip=False)
            ann.draggable()
            inset_anns.append(ann)

        # Remove default axis labels (we use draggable annotations instead)
        ax_inset.set_xlabel("")
        ax_inset.set_ylabel("")

        # Load saved positions
        _text_pos_file = _os.path.join(_os.path.dirname(db.DB_PATH),
                                       "_choro_inset_text_pos.json")
        _saved_text_pos = {}
        if _os.path.exists(_text_pos_file):
            try:
                _saved_text_pos = _json.load(open(_text_pos_file, "r", encoding="utf-8"))
            except Exception:
                pass
        # Restore positions for country labels
        _all_draggables = []
        for i, ann in enumerate(inset_anns):
            _all_draggables.append((f"__country_{i}__", ann))
        for key, ann in _all_draggables:
            if key in _saved_text_pos:
                xy, xytext = _saved_text_pos[key]
                ann.xy = tuple(xy)
                try:
                    ann.set_position(tuple(xytext))
                except Exception:
                    pass

        # Save positions on any mouse release (catches drag end)
        def _save_ann_positions():
            pos = {}
            for key, ann in _all_draggables:
                try:
                    pos[key] = [list(ann.xy), list(ann.get_position())]
                except Exception:
                    pass
            try:
                _json.dump(pos, open(_text_pos_file, "w", encoding="utf-8"),
                          ensure_ascii=False, indent=2)
            except Exception:
                pass

        # Hook into motion_notify to detect drag end
        _saved_positions_hook = {"counter": 0}

        def _ann_release_hook(event):
            _saved_positions_hook["counter"] += 1
            if _saved_positions_hook["counter"] % 3 == 0:
                _save_ann_positions()

        # --- Inset move + resize ---
        _inset_move_drag = {"active": False, "mode": None}
        _resize_margin = 16

        def _hit_inset(event):
            if event.x is None or event.y is None:
                return False
            try:
                fx, fy = fig.transFigure.inverted().transform((event.x, event.y))
            except Exception:
                return False
            bbox = ax_inset.get_position()
            return bbox.x0 <= fx <= bbox.x0 + bbox.width and bbox.y0 <= fy <= bbox.y0 + bbox.height

        def _inset_move_press(event):
            if event.button != 1:
                return
            if not _hit_inset(event):
                return
            # Don't start move/resize if user clicked on a bar
            for patch in ax_inset.patches:
                if patch.contains(event)[0]:
                    return
            _inset_move_drag["active"] = True
            _inset_move_drag["ref_x"] = event.x
            _inset_move_drag["ref_y"] = event.y
            _inset_move_drag["ref_bbox"] = ax_inset.get_position().bounds
            bbox_px = ax_inset.get_position().transformed(
                fig.transFigure).get_points()
            right_edge_px = bbox_px[1][0]
            if abs(event.x - right_edge_px) < _resize_margin:
                _inset_move_drag["mode"] = "resize"
            else:
                _inset_move_drag["mode"] = "move"

        def _inset_move_motion(event):
            if not _inset_move_drag["active"] or event.x is None:
                return
            bx, by, bw, bh = _inset_move_drag["ref_bbox"]
            if _inset_move_drag["mode"] == "resize":
                dx = (event.x - _inset_move_drag["ref_x"]) / (fig.bbox.width or 1)
                new_bw = max(0.12, min(0.82, bw + dx))
                ax_inset.set_position([bx, by, new_bw, bh])
            else:
                dx = (event.x - _inset_move_drag["ref_x"]) / (fig.bbox.width or 1)
                dy = (event.y - _inset_move_drag["ref_y"]) / (fig.bbox.height or 1)
                new_bx = max(0, min(1 - bw, bx + dx))
                new_by = max(0, min(1 - bh, by + dy))
                ax_inset.set_position([new_bx, new_by, bw, bh])
            canvas.draw_idle()

        def _inset_move_release(event):
            if _inset_move_drag["active"]:
                bx, by, bw, bh = ax_inset.get_position().bounds
                try:
                    with open(_inset_pos_file, "w", encoding="utf-8") as _f:
                        _json.dump([bx, by, bw, bh], _f)
                except Exception:
                    pass
                _save_ann_positions()
            _inset_move_drag["active"] = False
            _inset_move_drag["mode"] = None

        canvas.mpl_connect("button_press_event", _inset_move_press)
        canvas.mpl_connect("motion_notify_event", _inset_move_motion)
        canvas.mpl_connect("button_release_event", _inset_move_release)
        # Save annotation positions on any release
        canvas.mpl_connect("button_release_event", _ann_release_hook)

        ax.set_title("Chinoiserie Architecture Cases by Country",
                     fontsize=13, fontweight="bold", pad=10)
        fig.subplots_adjust(bottom=0.02, top=0.97, left=0.02, right=0.98)

        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._current_fig = fig
        self._country_choro_drawn = True


    def _draw_sankey_map(self, parent_frame):
        """国家组群桑基图: 国家 → 组群(可重复成员) → 案例数量档."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3
        import re as _re
        import os as _os
        from matplotlib.path import Path as MplPath
        from matplotlib.patches import PathPatch, Rectangle
        import matplotlib.patheffects as pe

        conn = sqlite3.connect(db.DB_PATH)
        rows = conn.execute(
            "SELECT country, COUNT(*) FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL "
            "AND name IS NOT NULL AND name != '' GROUP BY country").fetchall()
        conn.close()
        _zw = _re.compile(r'[​‌‍‎‏﻿\s]+')
        cnt = {_zw.sub('', c): n for c, n in rows if c}
        if not cnt:
            ttk.Label(parent_frame, text="暂无数据", font=("", 12)).pack(expand=True)
            return

        # 国家英文名
        en = {}
        try:
            import pandas as pd
            xlsx = _os.path.join(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))),
                                 "..", "所有国家中英文.xlsx")
            for _, r in pd.read_excel(xlsx).iterrows():
                cn, e = _zw.sub('', str(r["中文"])), str(r["英文"]).strip()
                if cn and cn != "nan" and e and e != "nan":
                    en[cn] = e
        except Exception:
            pass
        en.update({"俄罗斯": "Russia", "韩国": "Rep. of Korea", "朝鲜": "DPR Korea",
                   "文莱": "Brunei", "老挝": "Laos", "澳大利亚": "Australia",
                   "英国": "United Kingdom", "伊朗": "Iran", "美国": "United States",
                   "密克罗尼西亚联邦": "Micronesia", "摩尔多瓦": "Moldova"})

        groups = [
            ("ASEAN", "#e8894a",
             ["Thailand", "Malaysia", "Singapore", "Indonesia", "Viet Nam",
              "Philippines", "Cambodia", "Myanmar",
              "Lao People's Democratic Republic", "Brunei Darussalam"]),
            ("RCEP", "#4a85c4",
             ["Japan", "Republic of Korea", "Australia", "New Zealand",
              "Thailand", "Malaysia", "Singapore", "Indonesia", "Viet Nam",
              "Philippines", "Cambodia", "Myanmar",
              "Lao People's Democratic Republic", "Brunei Darussalam"]),
            ("Sinographic sphere", "#8f6db8",
             ["Japan", "Republic of Korea", "Democratic People's Republic of Korea", "Viet Nam"]),
            ("G7", "#5aa469",
             ["United States of America", "Japan", "Germany", "France",
              "United Kingdom of Great Britain and Northern Ireland", "Italy", "Canada"]),
            ("BRICS", "#c25b5b",
             ["Russian Federation", "India", "Brazil", "South Africa"]),
        ]
        gmap = {}
        for g, _col, lst in groups:
            for c in lst:
                gmap.setdefault(c, []).append(g)
        none_g = "Non-member"
        gcolor = {g: col for g, col, _ in groups}
        gcolor[none_g] = "#9a9a9a"
        g_order = [g for g, _, _ in groups] + [none_g]
        cls_order = [">50 cases", "11–50 cases", "3–10 cases", "1–2 cases"]

        def cls(n):
            if n > 50:
                return cls_order[0]
            if n >= 11:
                return cls_order[1]
            if n >= 3:
                return cls_order[2]
            return cls_order[3]

        lm, mr = {}, {}
        for c, n in cnt.items():
            for g in gmap.get(c, [none_g]):
                lm[(c, g)] = n
                mr[(g, cls(n))] = mr.get((g, cls(n)), 0) + n
        ltot = {c: sum(v for (cc, _g), v in lm.items() if cc == c) for c in cnt}
        mtot = {g: sum(v for (_c, gg), v in lm.items() if gg == g) for g in g_order}
        rtot = {k: sum(v for (_g, kk), v in mr.items() if kk == k) for k in cls_order}
        ncls = {k: sum(1 for c, n in cnt.items() if cls(n) == k) for k in cls_order}

        self.stats_label.config(text="  |  ".join(
            f"{g} {mtot[g]}" for g in g_order if mtot[g]))

        def lkey(c):
            gi = min(g_order.index(g) for g in gmap.get(c, [none_g]))
            return (gi, -cnt[c], c)

        l_order = sorted(cnt, key=lkey)

        fig = Figure(figsize=(9.5, 7.5), dpi=100)
        ax = fig.add_subplot(111)
        ax.axis("off")
        xcol = {"L": 0.0, "M": 0.50, "R": 0.96}
        nw = 0.018
        gap = {"L": 14, "M": 14, "R": 2.2}

        def layout(order, tot, col):
            pos, y = {}, 0.0
            for k in order:
                pos[k] = y
                y += tot[k] + gap[col]
            return pos, y - gap[col]

        lpos, lh = layout(cls_order, rtot, "L")   # 左列: 案例数量档
        mpos, mh = layout(g_order, mtot, "M")     # 中列: 组群
        rpos, rh = layout(l_order, ltot, "R")     # 右列: 国家
        h_all = max(lh, mh, rh)
        for pos, h in ((lpos, lh), (mpos, mh), (rpos, rh)):
            off = (h_all - h) / 2
            for k in pos:
                pos[k] += off

        def ribbon(x0, y0, x1, y1, w, color, alpha=0.38):
            x0 += nw
            dx = (x1 - x0) * 0.5
            verts = [(x0, y0), (x0 + dx, y0), (x1 - dx, y1), (x1, y1),
                     (x1, y1 + w), (x1 - dx, y1 + w), (x0 + dx, y0 + w),
                     (x0, y0 + w), (x0, y0)]
            codes = [MplPath.MOVETO, MplPath.CURVE4, MplPath.CURVE4, MplPath.CURVE4,
                     MplPath.LINETO, MplPath.CURVE4, MplPath.CURVE4, MplPath.CURVE4,
                     MplPath.CLOSEPOLY]
            ax.add_patch(PathPatch(MplPath(verts, codes), facecolor=color,
                                   edgecolor="none", alpha=alpha, zorder=1))

        # 第一段: 数量档(左) → 组群(中)
        tout = {k: 0.0 for k in cls_order}
        min_ = {g: 0.0 for g in g_order}
        for g in g_order:
            for k in cls_order:
                v = mr.get((g, k))
                if not v:
                    continue
                ribbon(xcol["L"], lpos[k] + tout[k], xcol["M"], mpos[g] + min_[g],
                       v, gcolor[g])
                tout[k] += v
                min_[g] += v
        # 第二段: 组群(中) → 国家(右)
        mout = {g: 0.0 for g in g_order}
        cin = {c: 0.0 for c in cnt}
        for c in l_order:
            for g in g_order:
                v = lm.get((c, g))
                if not v:
                    continue
                ribbon(xcol["M"], mpos[g] + mout[g], xcol["R"], rpos[c] + cin[c],
                       v, gcolor[g])
                mout[g] += v
                cin[c] += v

        # 左列: 案例数量档
        for k in cls_order:
            ax.add_patch(Rectangle((xcol["L"], lpos[k]), nw, rtot[k],
                                   facecolor="#666666", edgecolor="none", zorder=3))
            ax.text(xcol["L"] - 0.008, lpos[k] + rtot[k] / 2,
                    f"{k}\n({ncls[k]} countries)",
                    ha="right", va="center", fontsize=6, zorder=10)

        # 右列: 大国直接标, 小国最右侧罗列+引线
        direct_min = 25
        fan = [c for c in l_order if ltot[c] < direct_min]
        for c in l_order:
            ax.add_patch(Rectangle((xcol["R"], rpos[c]), nw, ltot[c],
                                   facecolor="#666666", edgecolor="none", zorder=3))
            if ltot[c] >= direct_min:
                ax.text(xcol["R"] + nw + 0.008, rpos[c] + ltot[c] / 2,
                        f"{en.get(c, c)} ({cnt[c]})",
                        ha="left", va="center", fontsize=5.5, zorder=10)
        fx_text = 1.30
        if fan:
            ys = [h_all - (i + 0.5) * h_all / len(fan) for i in range(len(fan))]
            for c, ty in zip(fan, ys):
                ny = rpos[c] + ltot[c] / 2
                ax.text(fx_text, ty, f"{en.get(c, c)} ({cnt[c]})", ha="left",
                        va="center", fontsize=4.6, color="#333333", zorder=10)
                ax.plot([xcol["R"] + nw + 0.003, fx_text - 0.004], [ny, ty],
                        color="#bbbbbb", linewidth=0.35, zorder=2)

        # 中列: 组群名放色块中心, 最上层, 白描边
        for g in g_order:
            ax.add_patch(Rectangle((xcol["M"], mpos[g]), nw, mtot[g],
                                   facecolor=gcolor[g], edgecolor="none", zorder=3))
            ax.text(xcol["M"] + nw / 2, mpos[g] + mtot[g] / 2, f"{g}\n({mtot[g]})",
                    ha="center", va="center", fontsize=6.5, fontweight="bold",
                    color="#222222", zorder=20,
                    path_effects=[pe.withStroke(linewidth=2.2, foreground="white")])

        ax.set_xlim(-0.17, 1.47)
        ax.set_ylim(-10, h_all + 15)
        ax.set_title("海外中式建筑国家组群桑基图", fontsize=13, fontweight="bold", pad=8)
        fig.subplots_adjust(bottom=0.02, top=0.95, left=0.01, right=0.99)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()
        self._sankey_drawn = True

    def _draw_chord_map(self, parent_frame):
        """城市类型弦图: 12类区域类型共现 (数据源: 论文/city_classification.csv + cases 表)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3
        import csv as _csv
        import os as _os
        import importlib.util
        import numpy as np
        import re as _re

        try:
            from mpl_chord_diagram import chord_diagram
        except ImportError:
            ttk.Label(parent_frame, text="请安装弦图库：pip install mpl-chord-diagram",
                      font=("", 12)).pack(expand=True)
            return

        # 复用论文生成脚本里的城市名归并与类型定义
        base = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
        spec = importlib.util.spec_from_file_location(
            "rst", _os.path.join(base, "论文", "regenerate_supplementary_tables.py"))
        rst = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(rst)

        cls = {}
        try:
            with open(rst.CITY_CLASS_CSV, encoding="utf-8-sig") as f:
                for r in _csv.DictReader(f):
                    cls[(r["Country"], r["City"])] = [l for l in r["Labels"].split("/")
                                                      if l.strip()]
        except OSError:
            ttk.Label(parent_frame, text="缺少 论文/city_classification.csv",
                      font=("", 12)).pack(expand=True)
            return

        conn = sqlite3.connect(db.DB_PATH)
        rows = conn.execute(
            "SELECT country, city FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL "
            "AND name IS NOT NULL AND name != ''").fetchall()
        conn.close()
        if not rows:
            ttk.Label(parent_frame, text="暂无数据", font=("", 12)).pack(expand=True)
            return

        types = rst.CITY_TYPE_ORDER
        idx = {t: i for i, t in enumerate(types)}
        M = np.zeros((len(types), len(types)))
        n_unclassified = 0
        for country, city in rows:
            c = _re.sub(r'[​‌‍‎‏﻿]+', '', country or '').strip()
            labs = cls.get((c, rst.norm_city(city)))
            if not labs:
                n_unclassified += 1
                continue
            k = len(labs)
            if k == 1:
                M[idx[labs[0]], idx[labs[0]]] += 1
            else:
                wt = 1.0 / (k - 1)   # 分数加权: 弧长 = 该类型真实案例数
                for i in range(k):
                    for j in range(i + 1, k):
                        a, b = idx[labs[i]], idx[labs[j]]
                        M[a, b] += wt
                        M[b, a] += wt

        counts = {t: int(round(M[idx[t]].sum())) for t in types}
        self.stats_label.config(text="  |  ".join(
            f"{t} {counts[t]}" for t in types if counts[t]) +
            (f"  |  未分类 {n_unclassified}" if n_unclassified else ""))

        names = [f"{rst.CITY_TYPE_EN[t]} ({counts[t]})" for t in types]
        colors = ["#d95f02", "#1b9e77", "#7570b3", "#e7298a", "#66a61e", "#e6ab02",
                  "#a6761d", "#666666", "#386cb0", "#f0027f", "#bf5b17", "#fdc086"]

        fig = Figure(figsize=(8, 7.5), dpi=100)
        ax = fig.add_subplot(111)
        chord_diagram(M, names=names, ax=ax, colors=colors,
                      use_gradient=True, sort="size", fontsize=8,
                      pad=2.0, chordwidth=0.7, width=0.08, rotate_names=True)
        # 外扩坐标范围, 给旋转标签留白, 避免被画布裁掉
        ax.set_xlim(-1.65, 1.65)
        ax.set_ylim(-1.55, 1.55)
        ax.set_aspect("equal")
        ax.set_title("海外中式建筑城市类型弦图", fontsize=12, fontweight="bold", pad=6)
        fig.subplots_adjust(bottom=0.0, top=0.94, left=0.0, right=1.0)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()
        self._chord_drawn = True

    def _draw_coastal_alt(self, parent_frame):
        """海陆分析备选图: 距海岸线距离 Violin (全球 + 沿海国/内陆国, 不分洲) (数据源: 附表S2)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import csv as _csv
        import os as _os
        import numpy as np

        base = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
        s2 = _os.path.join(base, "论文", "supplementary_table_coastal_distance.csv")
        try:
            with open(s2, encoding="utf-8-sig") as f:
                recs = [(r["Country"], float(r["Distance to nearest coastline (km)"]))
                        for r in _csv.DictReader(f)]
        except OSError:
            ttk.Label(parent_frame, text="缺少附表 S2, 请先运行 论文/regenerate_supplementary_tables.py",
                      font=("", 12)).pack(expand=True)
            return
        if not recs:
            ttk.Label(parent_frame, text="附表 S2 为空", font=("", 12)).pack(expand=True)
            return

        _landlocked = {
            "阿富汗", "安道尔", "亚美尼亚", "奥地利", "阿塞拜疆", "白俄罗斯",
            "不丹", "玻利维亚", "博茨瓦纳", "布基纳法索", "布隆迪", "中非",
            "乍得", "捷克", "斯威士兰", "埃塞俄比亚", "匈牙利", "哈萨克斯坦",
            "吉尔吉斯斯坦", "老挝", "莱索托", "列支敦士登", "卢森堡", "马拉维",
            "马里", "摩尔多瓦", "蒙古", "尼泊尔", "尼日尔", "北马其顿", "巴拉圭",
            "卢旺达", "圣马力诺", "塞尔维亚", "斯洛伐克", "南苏丹", "瑞士",
            "塔吉克斯坦", "土库曼斯坦", "乌干达", "乌兹别克斯坦", "梵蒂冈",
            "赞比亚", "津巴布韦",
        }

        def _is_landlocked(country_cn):
            for ll in _landlocked:
                if ll in country_cn or country_cn in ll:
                    return True
            return False

        dists = np.array([d for _, d in recs])
        d_coastal = np.array([d for c, d in recs if not _is_landlocked((c or "").strip())])
        d_landlocked = np.array([d for c, d in recs if _is_landlocked((c or "").strip())])

        med = float(np.median(dists))
        self.stats_label.config(
            text=f"{len(dists)} 例 | 中位 {med:.1f} km | ≤50km {(dists<=50).mean()*100:.1f}% | "
                 f"≤100km {(dists<=100).mean()*100:.1f}% | 沿海国 {len(d_coastal)} 例 | "
                 f"内陆国 {len(d_landlocked)} 例")

        # 蓝色渐变: 下(全部)最深 → 上(内陆国)最浅
        groups = [("All cases", dists, "#1f5fa8"),
                  ("Coastal-country\ncases", d_coastal, "#5b9bd5"),
                  ("Landlocked-country\ncases", d_landlocked, "#b7d4ec")]
        logs = [np.log10(np.maximum(g, 0.5)) for _, g, _ in groups]

        fig = Figure(figsize=(8, 5.2), dpi=100)
        ax = fig.add_subplot(111)
        # 横向小提琴: 下→上 = 全部 / 沿海国 / 内陆国
        vp = ax.violinplot(logs, positions=[1, 2, 3], orientation="horizontal",
                           showmedians=True, showextrema=False, widths=0.75)
        for body, (_, _, col) in zip(vp["bodies"], groups):
            body.set_facecolor(col)
            body.set_alpha(0.75)
            body.set_edgecolor("none")
        vp["cmedians"].set_color("#333333")
        vp["cmedians"].set_linewidth(1.2)

        # 每例叠加抖动散点 (内陆国样本小, 散点让分布可见)
        rng = np.random.default_rng(42)
        for i, lg in enumerate(logs, start=1):
            jitter = rng.uniform(-0.26, 0.26, len(lg))
            ax.scatter(lg, i + jitter, s=6, color=groups[i - 1][2],
                       alpha=0.4, edgecolors="none", zorder=4)

        # 文字标注（全部可拖拽, 位置存 data/coastal_alt_label_pos.json, 论文版沿用）
        import json as _json
        _lbl_file = _os.path.join(base, "data", "coastal_alt_label_pos.json")
        try:
            with open(_lbl_file, encoding="utf-8") as f:
                _saved = _json.load(f)
        except (OSError, ValueError):
            _saved = {}

        _made = []  # (key, annotation, default_ox, default_oy)

        def _ann(key, txt, xy, dox, doy, **kw):
            sv = _saved.get(key)
            ox, oy = (sv["ox"], sv["oy"]) if sv else (dox, doy)
            a = ax.annotate(txt, xy=xy, xytext=(ox, oy),
                            textcoords="offset points", fontsize=7.5, **kw)
            a.draggable()
            _made.append((key, a, dox, doy))

        # 竖向参考线: 全球中位 与 50 km
        for key, d0, txt in [("ref_median", med, f"Global median {med:.1f} km"),
                             ("ref_50km", 50, "50 km")]:
            ax.axvline(np.log10(d0), color="#c25b5b", linewidth=0.8, linestyle="--", zorder=2)
            _ann(key, txt, (np.log10(d0), 3.55), 3, 0, color="#993333")

        # 每组标注中位数
        for i, (name, g, _) in enumerate(groups, start=1):
            m = float(np.median(g))
            _ann(f"med_{i}", f"Median {m:.0f} km", (np.log10(max(m, 0.5)), i),
                 0, 30, color="#333333", ha="center")

        def _save_lbl(event):
            data = {}
            for key, a, dox, doy in _made:
                ox, oy = a.xyann
                if abs(ox - dox) > 0.1 or abs(oy - doy) > 0.1 or key in _saved:
                    data[key] = {"ox": round(float(ox), 1), "oy": round(float(oy), 1)}
            if data != _saved:
                _saved.clear()
                _saved.update(data)
                try:
                    with open(_lbl_file, "w", encoding="utf-8") as f:
                        _json.dump(data, f, indent=1)
                except OSError:
                    pass

        ax.set_yticks([1, 2, 3])
        ax.set_yticklabels([f"{name}\n(n = {len(g)})" for name, g, _ in groups], fontsize=8.5)
        ax.set_ylim(0.4, 3.75)
        ax.set_xticks([0, 1, 2, 3])
        ax.set_xticklabels(["1", "10", "100", "1000"], fontsize=8)
        ax.set_xlabel("Distance to nearest coastline (km, log scale)", fontsize=9)
        ax.grid(True, axis="x", linewidth=0.3, alpha=0.35)
        ax.set_title("海外中式建筑距海岸线距离分布（备选：Violin）",
                     fontsize=13, fontweight="bold", pad=10)
        fig.subplots_adjust(bottom=0.10, top=0.92, left=0.12, right=0.97)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.mpl_connect("button_release_event", _save_lbl)
        canvas.draw()
        self._coastal_alt_drawn = True

    def _draw_temporal_line(self, parent_frame):
        """时间分布三面板组图: 建设/损毁/修复 年份分布, 共享 x 轴 (25年分箱)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3 as _sqlite3
        import os as _os
        import re as _re
        from collections import Counter as _Counter

        base = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
        dbfile = _os.path.join(base, "data", "architecture.db")
        conn = _sqlite3.connect(dbfile)
        rows = conn.execute(
            "SELECT year_built, demolition_year, restoration_year FROM cases "
            "WHERE latitude IS NOT NULL AND name != ''").fetchall()
        conn.close()

        def _parse_year(s):
            """Return (year, kind); kind is exact/circa/century/None."""
            if s is None:
                return None, None
            s = str(s).strip()
            if s.lower() in ("", "unknown"):
                return None, None
            m = _re.match(r"^(\d{1,2})(st|nd|rd|th)$", s)
            if m:
                return (int(m.group(1)) - 1) * 100 + 50, "century"  # 世纪中点
            m = _re.match(r"^c\.\s*(\d{4})\s*[-–]\s*(\d{2,4})$", s, _re.I)
            if m:
                a = int(m.group(1))
                b = m.group(2)
                b = int(str(a)[:2] + b) if len(b) == 2 else int(b)
                return (a + b) // 2, "circa"  # circa 范围 -> 中点
            m = _re.match(r"^(\d{3,4})$", s)
            if m:
                return int(m.group(1)), "exact"
            m = _re.search(r"(\d{4})", s)  # 逗号多年份等兜底
            if m:
                return int(m.group(1)), "exact"
            return None, None

        built, demolished, restored = [], [], []
        built_exact, dem_exact, res_exact = [], [], []
        for yb, yd, yr in rows:
            v, k = _parse_year(yb)
            if v:
                built.append(v)
                if k == "exact":
                    built_exact.append(v)
            v, k = _parse_year(yd)
            if v:
                demolished.append(v)
                if k == "exact":
                    dem_exact.append(v)
            v, k = _parse_year(yr)
            if v:
                restored.append(v)
                if k == "exact":
                    res_exact.append(v)

        qb = _Counter((y // 25) * 25 for y in built)
        qd = _Counter((y // 25) * 25 for y in demolished)
        qr = _Counter((y // 25) * 25 for y in restored)
        # 十年放大图仅用精确年份（与论文图6口径一致）
        dec_b = _Counter((y // 10) * 10 for y in built_exact)
        dec_d = _Counter((y // 10) * 10 for y in dem_exact)
        dec_r = _Counter((y // 10) * 10 for y in res_exact)

        bins = list(range(175, 2050, 25))
        xs = [b + 12.5 for b in bins]

        fig = Figure(figsize=(11, 12), dpi=100)
        # 记录原始尺寸，导出时用于恢复（TkAgg 会把 figure 缩放到画布导致压扁）
        fig._intended_size_inches = (11.0, 12.0)
        ax1 = fig.add_subplot(311)
        ax2 = fig.add_subplot(312, sharex=ax1)
        ax3 = fig.add_subplot(313, sharex=ax1)

        # --- 建设面板 ---
        ax1.plot(xs, [qb.get(b, 0) for b in bins],
                 color="#1f5fa8", lw=1.7, marker="o", ms=3.2, zorder=3)
        ax1.set_ylabel("Founded per quarter-century", fontsize=10)
        ax1.set_xlim(100, 2075)
        ax1.set_ylim(0, max([qb.get(b, 0) for b in bins] + [1]) * 1.15)
        ax1.grid(axis="y", color="0.9", linewidth=0.7, zorder=0)
        for s in ("top", "right"):
            ax1.spines[s].set_visible(False)
        ax1.tick_params(labelsize=9)
        ax1.set_title("Founding years", fontsize=12, fontweight="bold")

        # 三个波峰放大 inset（位置: [左, 下, 宽, 高], 相对本面板坐标）
        waves = [
            ("Wave 1  1600-1799", 1600, 1800, [0.06, 0.52, 0.24, 0.32]),
            ("Wave 2  1800-1949", 1800, 1950, [0.35, 0.52, 0.24, 0.32]),
            ("Wave 3  1980-2025", 1980, 2030, [0.64, 0.52, 0.24, 0.32]),
        ]
        for label, lo, hi, pos in waves:
            axi = ax1.inset_axes(pos)
            b2 = list(range(lo, hi, 10))
            axi.plot([b + 5 for b in b2], [dec_b.get(b, 0) for b in b2],
                     color="#5b9bd5", lw=1.2, marker="o", ms=2.2, zorder=3)
            axi.set_xlim(lo - 6, hi + 6 if hi < 2030 else 2031)
            ymax = max(dec_b.get(b, 0) for b in b2)
            axi.set_ylim(0, ymax * 1.2)
            axi.set_xticks(range(lo, hi + 1, 50))
            axi.tick_params(labelsize=7)
            axi.grid(axis="y", color="0.92", linewidth=0.5, zorder=0)
            axi.set_title(label, fontsize=8.5)
            for s in ("top", "right"):
                axi.spines[s].set_visible(False)

        # --- 损毁面板 ---
        ax2.plot(xs, [qd.get(b, 0) for b in bins],
                 color="#c25b5b", lw=1.7, marker="o", ms=3.2, zorder=3)
        ax2.set_ylabel("Destroyed per quarter-century", fontsize=10)
        ax2.set_ylim(0, max([qd.get(b, 0) for b in bins] + [1]) * 1.2)
        ax2.grid(axis="y", color="0.9", linewidth=0.7, zorder=0)
        for s in ("top", "right"):
            ax2.spines[s].set_visible(False)
        ax2.tick_params(labelsize=9)
        ax2.set_title("Destruction years", fontsize=12, fontweight="bold")

        # 三个关键损毁时期 decade 放大
        dem_waves = [
            ("Imjin War 1575-99", 1570, 1600, [0.06, 0.52, 0.24, 0.32]),
            ("Anti-religious & WWII 1925-49", 1920, 1950, [0.35, 0.52, 0.24, 0.32]),
            ("Redevelopment 2000-", 2000, 2025, [0.64, 0.52, 0.24, 0.32]),
        ]
        for label, lo, hi, pos in dem_waves:
            axi = ax2.inset_axes(pos)
            b2 = list(range(lo, hi, 10))
            axi.plot([b + 5 for b in b2], [dec_d.get(b, 0) for b in b2],
                     color="#d99a9a", lw=1.2, marker="o", ms=2.2, zorder=3)
            axi.set_xlim(lo - 6, hi + 6)
            axi.set_ylim(0, max([dec_d.get(b, 0) for b in b2] + [1]) * 1.2)
            axi.set_xticks(range(lo, hi + 1, 10))
            axi.tick_params(labelsize=7)
            axi.grid(axis="y", color="0.92", linewidth=0.5, zorder=0)
            axi.set_title(label, fontsize=8.5)
            for s in ("top", "right"):
                axi.spines[s].set_visible(False)

        # --- 修复面板 ---
        ax3.plot(xs, [qr.get(b, 0) for b in bins],
                 color="#4a9b6a", lw=1.7, marker="o", ms=3.2, zorder=3)
        ax3.set_ylabel("Restored per quarter-century", fontsize=10)
        ax3.set_xlabel("Year", fontsize=10)
        ax3.set_ylim(0, max([qr.get(b, 0) for b in bins] + [1]) * 1.2)
        ax3.set_xticks(range(200, 2200, 200))
        ax3.grid(axis="y", color="0.9", linewidth=0.7, zorder=0)
        for s in ("top", "right"):
            ax3.spines[s].set_visible(False)
        ax3.tick_params(labelsize=9)
        ax3.set_title("Restoration years", fontsize=12, fontweight="bold")

        # 两个修复阶段 decade 放大
        res_waves = [
            ("Restorations 1950-79", 1950, 1980, [0.06, 0.52, 0.24, 0.32]),
            ("Restorations 1980-2025", 1980, 2025, [0.35, 0.52, 0.24, 0.32]),
        ]
        for label, lo, hi, pos in res_waves:
            axi = ax3.inset_axes(pos)
            b2 = list(range(lo, hi, 10))
            axi.plot([b + 5 for b in b2], [dec_r.get(b, 0) for b in b2],
                     color="#7bb894", lw=1.2, marker="o", ms=2.2, zorder=3)
            axi.set_xlim(lo - 6, hi + 6)
            axi.set_ylim(0, max([dec_r.get(b, 0) for b in b2] + [1]) * 1.2)
            axi.set_xticks(range(lo, hi + 1, 20))
            axi.tick_params(labelsize=7)
            axi.grid(axis="y", color="0.92", linewidth=0.5, zorder=0)
            axi.set_title(label, fontsize=8.5)
            for s in ("top", "right"):
                axi.spines[s].set_visible(False)

        self.stats_label.config(
            text=f"Founded {len(built)} | Destroyed {len(demolished)} | Restored {len(restored)} | "
                 f"25-year bins; founding: 3 wave insets, destruction: 3 period insets, restoration: 2 phase insets")

        fig.subplots_adjust(left=0.07, right=0.97, bottom=0.05, top=0.97, hspace=0.55)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()
        self._temporal_drawn = True

    def _draw_demolition_line(self, parent_frame):
        """损毁年份折线图: 每25年有记录损毁事件数, 与时间分布折线同轴口径 (数据源: 附表S7口径)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3 as _sqlite3
        import os as _os
        import re as _re
        from collections import Counter as _Counter

        base = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
        dbfile = _os.path.join(base, "data", "architecture.db")
        conn = _sqlite3.connect(dbfile)
        rows = conn.execute(
            "SELECT demolition_year, restoration_year FROM cases "
            "WHERE latitude IS NOT NULL AND name != '' "
            "AND demolition_year IS NOT NULL AND demolition_year != ''").fetchall()
        conn.close()

        years, rebuilt = [], 0
        for d, r in rows:
            ds = (d or "").strip()
            m = _re.match(r"^(\d{3,4})$", ds)
            if m:
                years.append(int(m.group(1)))
            else:
                m = _re.match(r"^(\d{1,2})(st|nd|rd|th)$", ds)
                if m:
                    years.append((int(m.group(1)) - 1) * 100 + 50)  # 世纪中点
                else:
                    continue
            if (r or "").strip():
                rebuilt += 1

        if not years:
            ttk.Label(parent_frame, text="无损毁记录", font=("", 12)).pack(expand=True)
            return

        q = _Counter((y // 25) * 25 for y in years)
        bins = list(range(175, 2050, 25))
        xs = [b + 12.5 for b in bins]
        counts = [q.get(b, 0) for b in bins]

        self.stats_label.config(
            text=f"有记录损毁事件 {len(years)} 例 | 其中有重建记录 {rebuilt} 例 "
                 f"({rebuilt/len(years)*100:.0f}%) | 25年分箱, 与时间分布折线同轴")

        fig = Figure(figsize=(11, 4.2), dpi=100)
        ax = fig.add_subplot(111)
        ax.plot(xs, counts, color="#c25b5b", lw=1.7, marker="o", ms=3.2, zorder=3)
        ax.set_xlabel("Year of destruction", fontsize=10)
        ax.set_ylabel("Cases destroyed per quarter-century", fontsize=10)
        ax.set_xlim(100, 2075)
        ax.set_ylim(0, max(counts) * 1.2)
        ax.set_xticks(range(200, 2200, 200))
        ax.grid(axis="y", color="0.9", linewidth=0.7, zorder=0)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        ax.tick_params(labelsize=9)
        ax.set_title("海外中式建筑损毁年份分布（有记录损毁事件）",
                     fontsize=13, fontweight="bold", pad=10)
        fig.subplots_adjust(bottom=0.12, top=0.90, left=0.07, right=0.97)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()
        self._demolition_drawn = True

    @staticmethod
    def _smooth_polyline(points):
        """Smooth a polyline using B-spline interpolation."""
        import numpy as np
        from scipy.interpolate import splprep, splev

        if len(points) < 2:
            return points

        x = np.array([p[0] for p in points])
        y = np.array([p[1] for p in points])

        # Remove near-duplicate consecutive points
        keep = [True]
        for i in range(1, len(x)):
            d = ((x[i] - x[i-1]) ** 2 + (y[i] - y[i-1]) ** 2) ** 0.5
            keep.append(d > 1e-6)
        x, y = x[keep], y[keep]

        if len(x) < 2:
            return points
        if len(x) <= 3:
            k = len(x) - 1
        else:
            k = 3  # cubic

        s = len(x) * 0.02  # light smoothing
        tck, u = splprep([x, y], s=s, k=k)
        n_out = max(200, len(x) * 20)
        u_new = np.linspace(0, 1, n_out)
        x_new, y_new = splev(u_new, tck)
        return [[float(x_new[i]), float(y_new[i])] for i in range(len(x_new))]

    @staticmethod
    def _make_double_line(points, offset=1.2):
        """Offset a polyline on both sides to create parallel double curves."""
        if len(points) < 2:
            return points, points

        left, right = [], []
        for i in range(len(points)):
            if i == 0:
                dx = points[1][0] - points[0][0]
                dy = points[1][1] - points[0][1]
            elif i == len(points) - 1:
                dx = points[-1][0] - points[-2][0]
                dy = points[-1][1] - points[-2][1]
            else:
                dx = points[i + 1][0] - points[i - 1][0]
                dy = points[i + 1][1] - points[i - 1][1]

            length = (dx * dx + dy * dy) ** 0.5
            if length < 1e-10:
                left.append(points[i][:])
                right.append(points[i][:])
                continue
            nx = -dy / length * offset
            ny = dx / length * offset

            left.append([points[i][0] + nx, points[i][1] + ny])
            right.append([points[i][0] - nx, points[i][1] - ny])

        return left, right

    def _load_coastal_lines(self):
        import json
        import os as _os3
        path = _os3.path.join(_os3.path.dirname(_os3.path.dirname(_os3.path.abspath(__file__))),
                              "data", "coastal_lines.json")
        try:
            with open(path, "r", encoding="utf-8") as f:
                self._coastal_lines = json.load(f)
        except (OSError, ValueError):
            self._coastal_lines = []

    def _save_coastal_lines(self):
        import json
        import os as _os4
        path = _os4.path.join(_os4.path.dirname(_os4.path.dirname(_os4.path.abspath(__file__))),
                              "data", "coastal_lines.json")
        try:
            _os4.makedirs(_os4.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._coastal_lines, f, ensure_ascii=False, indent=1)
        except OSError:
            pass

    def _redraw_coastal_lines(self, ax, ccrs, state):
        for artist in state["line_artists"]:
            artist.remove()
        state["line_artists"].clear()
        for stroke in self._coastal_lines:
            if not stroke:
                continue
            # Check format: [[lon,lat],...] (polyline) or [lon1,lat1,lon2,lat2] (segment)
            if isinstance(stroke[0], list):
                # Polyline from freehand drawing → double curve
                smoothed = self._smooth_polyline(stroke)
                left, right = self._make_double_line(smoothed)
                for curve in [left, right]:
                    clons = [p[0] for p in curve]
                    clats = [p[1] for p in curve]
                    line = ax.plot(clons, clats, color="#4a85c4", linewidth=5, alpha=0.30,
                                   transform=ccrs.PlateCarree(), zorder=3)
                    state["line_artists"].append(line[0])
            else:
                # Legacy segment format [lon1, lat1, lon2, lat2]
                lon1, lat1, lon2, lat2 = stroke
                elon1, elat1, elon2, elat2 = self._extend_segment(lon1, lat1, lon2, lat2)
                line = ax.plot([elon1, elon2], [elat1, elat2],
                               color="#4a85c4", linewidth=5, alpha=0.30,
                               transform=ccrs.PlateCarree(), zorder=3)
                state["line_artists"].append(line[0])

    @staticmethod
    def _extend_segment(lon1, lat1, lon2, lat2, fraction=0.12):
        dlon = lon2 - lon1
        dlat = lat2 - lat1
        return (lon1 - dlon * fraction, lat1 - dlat * fraction,
                lon2 + dlon * fraction, lat2 + dlat * fraction)

    def _coastal_undo(self):
        if not self._coastal_lines:
            return
        self._coastal_lines.pop()
        self._save_coastal_lines()
        if self._coastal_state is not None:
            s = self._coastal_state
            s["selected_idx"] = None
            s["select_hl"].set_visible(False)
            self._redraw_coastal_lines(s["ax"], s["ccrs"], s)
            s["canvas"].draw_idle()

    def _coastal_clear_all(self):
        if not self._coastal_lines:
            return
        if not messagebox.askyesno("确认", "确定要清除所有连线吗？"):
            return
        self._coastal_lines.clear()
        self._save_coastal_lines()
        if self._coastal_state is not None:
            s = self._coastal_state
            s["selected_idx"] = None
            s["select_hl"].set_visible(False)
            self._redraw_coastal_lines(s["ax"], s["ccrs"], s)
            s["canvas"].draw_idle()

    @staticmethod
    def _place_country_labels(fig, ax, canvas, ccrs, annotations, all_lons, all_lats):
        """Greedy label placement: try short offsets first, expanding outward in 8
        directions until a spot is found that covers neither case points nor other
        labels. Uncrowded labels therefore get short (or no) leader lines."""
        import json
        import os
        import numpy as np

        pos_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                "data", "map_label_positions.json")
        try:
            with open(pos_file, "r", encoding="utf-8") as f:
                saved = json.load(f)
        except (OSError, ValueError):
            saved = {}

        renderer = canvas.get_renderer()
        pc = ccrs.PlateCarree()
        ann_trans = pc._as_mpl_transform(ax)

        # Display (pixel) coords of every case point
        pts = ax.projection.transform_points(pc, np.asarray(all_lons, dtype=float),
                                             np.asarray(all_lats, dtype=float))
        pts_disp = ax.transData.transform(pts[:, :2])
        pts_disp = pts_disp[np.isfinite(pts_disp).all(axis=1)]

        px_per_pt = fig.dpi / 72.0
        point_r = 7.0  # approx. scatter marker radius in px (s=80)
        ax_box = ax.get_window_extent(renderer)

        dirs = [
            (1, 0, "left", "center"), (-1, 0, "right", "center"),
            (0, 1, "center", "bottom"), (0, -1, "center", "top"),
            (0.71, 0.71, "left", "bottom"), (-0.71, 0.71, "right", "bottom"),
            (0.71, -0.71, "left", "top"), (-0.71, -0.71, "right", "top"),
        ]
        radii = [9, 15, 22, 31, 42, 55, 70]  # candidate offsets in points, shortest first

        def rect_overlap(a, b):
            w = min(a[2], b[2]) - max(a[0], b[0])
            h = min(a[3], b[3]) - max(a[1], b[1])
            return w * h if (w > 0 and h > 0) else 0.0

        # Measure each label's box size once
        items = []
        for lon, lat, text, country in annotations:
            t = ax.text(0.5, 0.5, text, fontsize=7, fontweight="bold",
                        transform=ax.transAxes)
            tb = t.get_window_extent(renderer)
            t.remove()
            x, y = ax.projection.transform_point(lon, lat, pc)
            px, py = ax.transData.transform((x, y))
            items.append({"lon": lon, "lat": lat, "text": text, "country": country,
                          "w": tb.width + 8, "h": tb.height + 8,
                          "px": float(px), "py": float(py)})

        def label_rect(it, ox, oy, ha, va):
            w, h = it["w"], it["h"]
            apx = it["px"] + ox * px_per_pt
            apy = it["py"] + oy * px_per_pt
            x0 = apx if ha == "left" else apx - w if ha == "right" else apx - w / 2
            y0 = apy if va == "bottom" else apy - h if va == "top" else apy - h / 2
            return (x0, y0, x0 + w, y0 + h)

        # Manually saved positions win; reserve their space before auto-placement
        occupied = []
        for it in items:
            sv = saved.get(it["country"])
            if sv:
                it["pos"] = (float(sv["ox"]), float(sv["oy"]),
                             sv.get("ha", "left"), sv.get("va", "center"))
                occupied.append(label_rect(it, *it["pos"]))

        # Place labels in crowded areas first so they get the good spots
        for it in items:
            d = np.hypot(pts_disp[:, 0] - it["px"], pts_disp[:, 1] - it["py"])
            it["density"] = int((d < 150).sum())
        items.sort(key=lambda it: -it["density"])

        for it in items:
            if "pos" in it:
                continue
            m = (np.abs(pts_disp[:, 0] - it["px"]) < 400) & \
                (np.abs(pts_disp[:, 1] - it["py"]) < 400)
            near = pts_disp[m]
            own = (np.abs(near[:, 0] - it["px"]) < 1) & (np.abs(near[:, 1] - it["py"]) < 1)
            near = near[~own]

            best = None  # (score, ox, oy, ha, va, rect)
            found_clean = False
            for r in radii:
                for dx, dy, ha, va in dirs:
                    ox, oy = dx * r, dy * r
                    rect = label_rect(it, ox, oy, ha, va)

                    cost = 0.0
                    if (rect[0] < ax_box.x0 or rect[2] > ax_box.x1
                            or rect[1] < ax_box.y0 or rect[3] > ax_box.y1):
                        cost += 1e6
                    for oc in occupied:
                        cost += 3.0 * rect_overlap(rect, oc)
                    if len(near):
                        inx = (np.minimum(rect[2], near[:, 0] + point_r)
                               - np.maximum(rect[0], near[:, 0] - point_r))
                        iny = (np.minimum(rect[3], near[:, 1] + point_r)
                               - np.maximum(rect[1], near[:, 1] - point_r))
                        hit = (inx > 0) & (iny > 0)
                        cost += float((inx[hit] * iny[hit]).sum())

                    score = cost + r * 1.5  # prefer shorter leader lines
                    if cost == 0:
                        best = (score, ox, oy, ha, va, rect)
                        found_clean = True
                        break
                    if best is None or score < best[0]:
                        best = (score, ox, oy, ha, va, rect)
                if found_clean:
                    break

            _, ox, oy, ha, va, rect = best
            occupied.append(rect)
            it["pos"] = (ox, oy, ha, va)

        made = []    # (country_key, ann, iox, ioy)
        for it in items:
            ox, oy, ha, va = it["pos"]
            short = (ox * ox + oy * oy) ** 0.5 <= radii[0]

            ann = ax.annotate(
                it["text"],  # "Country\nBuilding Name"
                xy=(it["lon"], it["lat"]),
                xycoords=ann_trans,
                xytext=(ox, oy),
                textcoords="offset points",
                fontsize=7,
                fontweight="bold",
                color="#1e3a5f",
                ha=ha,
                va="center",
                bbox=dict(
                    facecolor="white", alpha=0.5, edgecolor="none",
                    boxstyle="round,pad=0.25"
                ),
                arrowprops=dict(
                    arrowstyle="-",
                    color="#888888",
                    lw=0.5,
                    connectionstyle="arc3,rad=0.15",
                ) if not short else dict(
                    arrowstyle="-",
                    color="#888888",
                    lw=0.5,
                ),
                zorder=10,
            )
            ann.draggable()
            made.append((it["country"], ann, ox, oy))

        def _save_label_positions(event):
            data = {}
            for country, ann, iox, ioy in made:
                ox, oy = ann.xyann
                moved = abs(ox - iox) > 0.1 or abs(oy - ioy) > 0.1
                if moved or country in saved:
                    data[country] = {"ox": round(float(ox), 1), "oy": round(float(oy), 1),
                                     "ha": ann.get_ha(), "va": ann.get_va()}
            if data != saved:
                saved.clear()
                saved.update(data)
                try:
                    os.makedirs(os.path.dirname(pos_file), exist_ok=True)
                    with open(pos_file, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=1)
                except OSError:
                    pass

        canvas.mpl_connect("button_release_event", _save_label_positions)

    def _draw_flowchart(self, parent_frame):
        """Research design flowchart: five-stage workflow (Section 5.2)."""
        import matplotlib.patches as mpatches
        import json, os

        fig = Figure(figsize=(15, 3.8), dpi=100)
        fig.patch.set_facecolor("white")
        ax = fig.add_subplot(111)
        ax.set_xlim(0, 15)
        ax.set_ylim(0, 3.8)
        ax.axis("off")

        STAGE_COLORS = ["#1a5276", "#117a65", "#b9770e", "#943126", "#6c3483"]
        ARROW_COLOR = "#95a5a6"

        stages = [
            {"num": "1", "title": "Multi-lingual\nSystematic Search",
             "detail": "Grey lit. (Google + Baidu)\nScholarly lit. (Google Scholar)\n194 countries and areas"},
            {"num": "2", "title": "LLM-Assisted\nScreening",
             "detail": "DeepSeek-chat (grey)\nGemini-2.5-flash (scholarly)\nRelevance + justification"},
            {"num": "3", "title": "Case\nExtraction",
             "detail": "Full-text retrieval\nLLM extraction (DeepSeek-chat)\nStructured JSON output"},
            {"num": "4", "title": "Deduplication &\nVerification",
             "detail": "Multi-source verification\nImage-based review\nAHP evaluation (8 criteria)"},
            {"num": "5", "title": "Spatial & Temporal\nAnalysis",
             "detail": "World distribution · Coastal\nContinental · Country · City\nFounding · Destruction · Restoration"},
        ]

        n = len(stages)
        box_w = 2.2
        box_h = 2.2
        gap_x = 0.8
        total_w = n * box_w + (n - 1) * gap_x
        start_x = (15 - total_w) / 2
        box_y = 1.2

        # Load saved positions
        pos_file = os.path.join(os.path.dirname(__file__), "..", "论文", "flowchart_positions.json")
        _saved = {}
        try:
            with open(pos_file, "r", encoding="utf-8") as f:
                _saved = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            pass

        _made = []

        def _ann(key, txt, xy, dox, doy, **kw):
            sv = _saved.get(key)
            ox, oy = (sv["ox"], sv["oy"]) if sv else (dox, doy)
            a = ax.annotate(txt, xy=xy, xytext=(ox, oy),
                            textcoords="offset points", **kw)
            a.draggable()
            _made.append((key, a, dox, doy))

        for i, s in enumerate(stages):
            x = start_x + i * (box_w + gap_x)
            cx = x + box_w / 2

            box = mpatches.FancyBboxPatch(
                (x, box_y), box_w, box_h,
                boxstyle="round,pad=0.15", facecolor="white",
                edgecolor=STAGE_COLORS[i], linewidth=2.2, zorder=2
            )
            ax.add_patch(box)

            # Number circle - use Circle for perfect roundness
            circle = mpatches.Circle(
                (cx, box_y + box_h - 0.45), 0.38,
                facecolor=STAGE_COLORS[i], edgecolor="none", zorder=3,
                transform=ax.transData
            )
            ax.add_patch(circle)
            ax.text(cx, box_y + box_h - 0.45, s["num"],
                    ha="center", va="center", fontsize=14, fontweight="bold",
                    color="white", zorder=4)

            # Title (draggable)
            _ann(f"title_{i}", s["title"],
                 (cx, box_y + box_h - 1.0), 0, 0,
                 fontsize=9, fontweight="bold", color=STAGE_COLORS[i],
                 ha="center", va="top", linespacing=1.25,
                 bbox=dict(boxstyle="round,pad=0.08", facecolor="white",
                           edgecolor="none", alpha=0.85))

            # Detail (draggable)
            _ann(f"detail_{i}", s["detail"],
                 (cx, box_y + 0.4), 0, 0,
                 fontsize=7.2, color="#566573", ha="center", va="center",
                 linespacing=1.35,
                 bbox=dict(boxstyle="round,pad=0.08", facecolor="white",
                           edgecolor="none", alpha=0.85))

            if i < n - 1:
                nx = start_x + (i + 1) * (box_w + gap_x)
                ax.annotate("", xy=(nx, box_y + box_h / 2),
                            xytext=(x + box_w, box_y + box_h / 2),
                            arrowprops=dict(arrowstyle="->", color=ARROW_COLOR,
                                            lw=2, shrinkA=7, shrinkB=7),
                            zorder=1)

        _ann("title_main", "Research Design: Five-Stage Workflow",
             (7.5, 3.5), 0, 0,
             fontsize=13, fontweight="bold", color="#2c3e50",
             ha="center", va="center")

        def _save_positions(event):
            data = {}
            for key, a, dox, doy in _made:
                ox, oy = a.xyann
                if abs(ox - dox) > 0.1 or abs(oy - doy) > 0.1 or key in _saved:
                    data[key] = {"ox": round(float(ox), 1), "oy": round(float(oy), 1)}
            if data != _saved:
                _saved.clear()
                _saved.update(data)
                try:
                    os.makedirs(os.path.dirname(pos_file), exist_ok=True)
                    with open(pos_file, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=1)
                except OSError:
                    pass

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(expand=True)
        canvas.draw()
        canvas.mpl_connect("button_release_event", _save_positions)
        self._flowchart_drawn = True

    def _draw_gdp_chart(self, parent_frame):
        """Three-line comparison: cases founded vs China GDP total vs GDP per capita."""
        import csv, os, re, sqlite3
        import numpy as np
        from matplotlib.ticker import FuncFormatter

        HERE = os.path.dirname(os.path.abspath(__file__))
        DB = os.path.join(HERE, "..", "data", "architecture.db")

        # --- Founding data (25-year bins) ---
        def parse_year(y):
            if y is None:
                return None, None
            y = str(y).strip()
            if not y or y.lower() == "unknown":
                return None, None
            m = re.fullmatch(r"(\d{1,2})(st|nd|rd|th)", y)
            if m:
                return None, int(m.group(1))
            m = re.fullmatch(r"c\.\s*(\d{4})\s*[-–]\s*(\d{2,4})", y, re.I)
            if m:
                a = int(m.group(1)); b = m.group(2)
                b = int(str(a)[:2] + b) if len(b) == 2 else int(b)
                return (a + b) // 2, None
            if re.search(r"[，,]", y):
                m = re.search(r"(\d{4})", y)
                if m:
                    return int(m.group(1)), None
            if re.fullmatch(r"\d{3,4}", y):
                return int(y), None
            return None, None

        con = sqlite3.connect(DB)
        rows = con.execute(
            "SELECT year_built FROM cases WHERE latitude IS NOT NULL "
            "AND longitude IS NOT NULL AND name != ''").fetchall()
        con.close()
        years_list = []
        for (yb,) in rows:
            y, c = parse_year(yb)
            if y is not None: years_list.append(y)
            elif c is not None: years_list.append((c - 1) * 100 + 50)
        bins = {}
        for y in years_list:
            b = (y // 25) * 25
            bins[b] = bins.get(b, 0) + 1

        # --- GDP data from 25-year grid CSV ---
        grid_path = os.path.join(HERE, "..", "论文", "maddison_china_25yr_grid.csv")
        gdp_years, gdp_total, gdp_pc = [], [], []
        with open(grid_path) as f:
            for row in csv.DictReader(f):
                gdp_years.append(float(row["quarter_century_midpoint"]))
                gdp_total.append(float(row["gdp_total_trillion_2011intl"]))
                gdp_pc.append(float(row["gdp_per_capita_2011intl"]))

        fig = Figure(figsize=(12, 6), dpi=100)
        ax = fig.add_subplot(111)

        C_FOUND = "#1a3a5c"
        C_TOTAL = "#c0392b"
        C_PC = "#e67e22"

        xs = sorted(bins)
        ax.plot([x + 12.5 for x in xs], [bins[x] for x in xs],
                color=C_FOUND, lw=1.8, zorder=5,
                label="Cases founded per quarter-century")
        ax.set_xlim(0, 2025)
        ax.set_xticks(range(0, 2025, 250))
        ax.set_ylim(0, max(bins.values()) * 1.15)
        ax.set_xlabel("Year (CE)")
        ax.set_ylabel("Cases founded per quarter-century", color=C_FOUND)
        ax.tick_params(axis="y", colors=C_FOUND)
        ax.spines["top"].set_visible(False)

        # GDP total (right axis)
        ax2 = ax.twinx()
        ax2.plot(gdp_years, gdp_total, color=C_TOTAL, lw=1.4, zorder=4,
                 label="China GDP (trillion 2011 int.-$)")
        ax2.set_ylabel("China GDP (trillion 2011 int.-$)", color=C_TOTAL)
        ax2.tick_params(axis="y", colors=C_TOTAL)
        ax2.spines["top"].set_visible(False)

        # GDP per capita (far right axis)
        ax3 = ax.twinx()
        ax3.spines["right"].set_position(("outward", 42))
        ax3.plot(gdp_years, gdp_pc, color=C_PC, lw=1.4, zorder=4,
                 label="China GDP per capita (2011 int.-$)")
        ax3.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v/1000:g}k"))
        ax3.set_ylabel("China GDP per capita (2011 int.-$)", color=C_PC)
        ax3.tick_params(axis="y", colors=C_PC)
        ax3.spines["top"].set_visible(False)

        # Legend
        lines, labels = [], []
        for a in (ax, ax2, ax3):
            l, lb = a.get_legend_handles_labels()
            lines += l; labels += lb
        ax.legend(lines, labels, loc="upper left", frameon=False, fontsize=8)

        fig.tight_layout()
        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()
        self._gdp_drawn = True

    def _draw_pearson_chart(self, parent_frame):
        """Bar chart: Pearson r (cases founded vs China GDP total) by period with 95% CI."""
        import csv as _csv, os as _os, sqlite3
        import numpy as np
        from scipy import stats as _stats
        import re as _re

        HERE = _os.path.dirname(_os.path.abspath(__file__))
        DB = _os.path.join(HERE, "..", "data", "architecture.db")
        GRID = _os.path.join(HERE, "..", "论文", "maddison_china_25yr_grid.csv")

        def _parse_y(y):
            if y is None:
                return None
            y = str(y).strip()
            if not y or y.lower() == "unknown":
                return None
            m = _re.fullmatch(r"(\d{1,2})(st|nd|rd|th)", y)
            if m: return (int(m.group(1)) - 1) * 100 + 50
            m = _re.fullmatch(r"c\.\s*(\d{4})\s*[-–]\s*(\d{2,4})", y, _re.I)
            if m:
                a = int(m.group(1)); b = m.group(2)
                b = int(str(a)[:2] + b) if len(b) == 2 else int(b)
                return (a + b) // 2
            if _re.search(r"[，,]", y):
                m = _re.search(r"(\d{4})", y)
                if m: return int(m.group(1))
            if _re.fullmatch(r"\d{3,4}", y):
                return int(y)
            return None

        con = sqlite3.connect(DB)
        rows = con.execute(
            "SELECT year_built FROM cases WHERE latitude IS NOT NULL "
            "AND name != ''").fetchall()
        con.close()
        years = []
        for (yb,) in rows:
            v = _parse_y(yb)
            if v is not None:
                years.append(v)

        founding_bins = {}
        for y in years:
            b = (y // 25) * 25
            founding_bins[b] = founding_bins.get(b, 0) + 1

        gdp_years, gdp_total = [], []
        with open(GRID) as f:
            for row in _csv.DictReader(f):
                gdp_years.append(float(row["quarter_century_midpoint"]))
                gdp_total.append(float(row["gdp_total_trillion_2011intl"]))

        x_founding, x_gdp = [], []
        for gy, gv in zip(gdp_years, gdp_total):
            b = int((gy // 25) * 25)
            x_founding.append(founding_bins.get(b, 0))
            x_gdp.append(gv)

        def _pearson(x, y):
            if len(x) < 3:
                return None, None, None, None, None
            r, p = _stats.pearsonr(x, y)
            n = len(x)
            z = np.arctanh(r)
            se = 1 / np.sqrt(n - 3)
            z_lo = z - 1.96 * se
            z_hi = z + 1.96 * se
            return r, np.tanh(z_lo), np.tanh(z_hi), p, n

        def _sig(p):
            if p is None: return "n/a"
            if p < 0.001: return "***"
            if p < 0.01: return "**"
            if p < 0.05: return "*"
            return "n.s."

        periods = []
        for label, lo, hi in [
            ("1000–2024\n(Full)", 1000, 2024),
            ("1000–1599\n(Pre-Wave)", 1000, 1600),
            ("1600–1799\n(Wave I)", 1600, 1800),
            ("1800–1949\n(Wave II + Trough)", 1800, 1950),
        ]:
            indices = [i for i, gy in enumerate(gdp_years) if lo <= gy <= hi]
            xf = [x_founding[i] for i in indices]
            xg = [x_gdp[i] for i in indices]
            r, r_lo, r_hi, p, n = _pearson(xf, xg)
            if r is None:
                continue
            periods.append((label, round(r, 4), round(r_lo, 4), round(r_hi, 4),
                          _sig(p), n, round(p, 5)))

        fig = Figure(figsize=(9, 5.5), dpi=100)
        ax = fig.add_subplot(111)

        labels = [p[0] for p in periods]
        r_vals = [p[1] for p in periods]
        ci_lo = [p[1] - p[2] for p in periods]  # distance below r
        ci_hi = [p[3] - p[1] for p in periods]  # distance above r
        sigs = [p[4] for p in periods]

        x = range(len(periods))
        colors = ["#7BA7BC", "#9DC0A8", "#E89F7E", "#BFB5A8"]
        bars = ax.bar(x, r_vals, color=colors, width=0.5, zorder=3,
                      yerr=[ci_lo, ci_hi], capsize=6, error_kw={"lw": 1.2, "zorder": 4})

        # significance stars
        for i, (r, sig) in enumerate(zip(r_vals, sigs)):
            y_pos = r + ci_hi[i] + 0.06 if r >= 0 else r + ci_hi[i] + 0.10
            ax.text(i, y_pos, sig, ha="center", va="bottom", fontsize=11, fontweight="bold")

        ax.axhline(y=0, color="black", lw=0.6, zorder=1)
        ax.axhline(y=1, color="#cccccc", lw=0.5, ls="--", zorder=1)
        ax.set_xticks(list(x))
        ax.set_xticklabels(labels, fontsize=9)
        ax.set_ylabel("Pearson r (cases founded vs GDP total)", fontsize=10)
        ax.set_ylim(-0.85, 1.15)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="y", color="#e0e0e0", lw=0.4, zorder=0)

        # annotation
        ax.text(0.5, -0.18,
                "r = Σ(x−x̄)(y−ȳ) / √[Σ(x−x̄)² Σ(y−ȳ)²];  95% CI via Fisher z-transformation;  "
                "GDP data from Maddison Project Database 2023 (25-year grid)",
                transform=ax.transAxes, ha="center", fontsize=7.5,
                fontstyle="italic")
        ax.text(0.5, -0.26,
                "*** p < 0.001    ** p < 0.01    * p < 0.05    n.s. p ≥ 0.05",
                transform=ax.transAxes, ha="center", fontsize=7.5, color="#555555")

        fig.tight_layout()
        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()
        self._pearson_drawn = True

    def _draw_sa_time_figure(self, parent_frame):
        """SA 合并图1: 建设时间 + 损毁时间, 共享时间横轴, 双Y轴 (左建设/右损毁)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import sqlite3 as _sqlite3, os as _os, re as _re, json as _json
        from collections import Counter as _Counter

        base = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
        dbfile = _os.path.join(base, "data", "architecture.db")
        conn = _sqlite3.connect(dbfile)
        found_rows = conn.execute(
            "SELECT year_built FROM cases "
            "WHERE latitude IS NOT NULL AND name != ''").fetchall()
        demo_rows = conn.execute(
            "SELECT demolition_year FROM cases "
            "WHERE latitude IS NOT NULL AND name != '' "
            "AND demolition_year IS NOT NULL AND demolition_year != ''").fetchall()
        conn.close()

        def _parse(ys):
            ys = (ys or "").strip()
            m = _re.match(r"^(\d{3,4})$", ys)
            if m:
                return int(m.group(1)), True
            m = _re.match(r"^(\d{1,2})(st|nd|rd|th)$", ys)
            if m:
                return (int(m.group(1)) - 1) * 100 + 50, False
            return None, False

        found_years, found_exact = [], []
        for (y,) in found_rows:
            v, exact = _parse(y)
            if v is not None:
                found_years.append(v)
                if exact:
                    found_exact.append(v)
        demo_years = []
        for (y,) in demo_rows:
            v, exact = _parse(y)
            if v is not None:
                demo_years.append(v)

        qf = _Counter((y // 25) * 25 for y in found_years)
        qd = _Counter((y // 25) * 25 for y in demo_years)
        dec = _Counter((y // 10) * 10 for y in found_exact)
        bins = list(range(175, 2050, 25))
        xs = [b + 12.5 for b in bins]
        f_counts = [qf.get(b, 0) for b in bins]
        d_counts = [qd.get(b, 0) for b in bins]

        self.stats_label.config(
            text=f"建设 {len(found_years)} 例(精确 {len(found_exact)}) | 损毁 {len(demo_years)} 例 | "
                 f"25年分箱, 共享时间横轴, 双Y轴 + 三波峰10年放大")

        fig = Figure(figsize=(11, 5.5), dpi=100)
        ax = fig.add_subplot(111)
        ax.plot(xs, f_counts, color="#1f5fa8", lw=1.7, marker="o", ms=3.2,
                zorder=3, label="Founded")
        ax.set_xlabel("Year (CE)", fontsize=10)
        ax.set_ylabel("Cases founded per quarter-century",
                      color="#1f5fa8", fontsize=10)
        ax.set_xlim(100, 2075)
        ax.set_ylim(0, max(f_counts) * 1.15)
        ax.set_xticks(range(200, 2200, 200))
        ax.grid(axis="y", color="0.9", linewidth=0.7, zorder=0)
        ax.tick_params(axis="y", colors="#1f5fa8", labelsize=9)
        ax.tick_params(axis="x", labelsize=9)
        ax.spines["top"].set_visible(False)
        ax.spines["left"].set_color("#1f5fa8")

        ax2 = ax.twinx()
        ax2.plot(xs, d_counts, color="#c25b5b", lw=1.7, marker="s", ms=3.2,
                 zorder=3, label="Destroyed")
        ax2.set_ylabel("Cases destroyed per quarter-century",
                       color="#c25b5b", fontsize=10)
        ax2.set_ylim(0, max(d_counts) * 1.4 if max(d_counts) else 1)
        ax2.tick_params(axis="y", colors="#c25b5b", labelsize=9)
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_color("#c25b5b")

        insets_pos_file = _os.path.join(base, "data", "sa_time_insets_pos.json")
        try:
            with open(insets_pos_file) as f:
                saved_insets = _json.load(f)
            if not isinstance(saved_insets, dict):
                saved_insets = {}
        except (OSError, ValueError):
            saved_insets = {}

        wave_defs = [
            ("Wave 1  1600-1799", 1600, 1800, [0.16, 0.60, 0.18, 0.24]),
            ("Wave 2  1800-1949", 1800, 1950, [0.41, 0.60, 0.18, 0.24]),
            ("Wave 3  1950-2025", 1950, 2030, [0.66, 0.60, 0.18, 0.24]),
        ]
        wave_axes = []
        for label, lo, hi, dpos in wave_defs:
            pos = saved_insets.get(label, dpos)
            axi = fig.add_axes(pos, facecolor="white")
            b2 = list(range(lo, hi, 10))
            axi.plot([b + 5 for b in b2], [dec.get(b, 0) for b in b2],
                     color="#5b9bd5", lw=1.2, marker="o", ms=2.2, zorder=3)
            axi.set_xlim(lo - 6, hi + 6 if hi < 2030 else 2031)
            ymax = max(dec.get(b, 0) for b in b2)
            axi.set_ylim(0, ymax * 1.2)
            axi.set_xticks(range(lo, hi + 1, 50))
            axi.tick_params(labelsize=7)
            axi.grid(axis="y", color="0.92", linewidth=0.5, zorder=0)
            axi.set_title(label, fontsize=8.5)
            for s in ("top", "right"):
                axi.spines[s].set_visible(False)
            wave_axes.append((label, axi))

        from matplotlib.lines import Line2D as _Line2D
        handles = [
            _Line2D([], [], color="#1f5fa8", lw=1.7, marker="o", ms=3.2,
                    label="Founded"),
            _Line2D([], [], color="#c25b5b", lw=1.7, marker="s", ms=3.2,
                    label="Destroyed"),
            _Line2D([], [], color="#5b9bd5", lw=1.2, marker="o", ms=2.2,
                    label="Wave insets: exactly dated cases, per decade"),
        ]
        ax.legend(handles=handles, loc="upper left", fontsize=9, frameon=False)

        fig.subplots_adjust(bottom=0.10, top=0.92, left=0.08, right=0.92)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)

        drag = {"mode": None, "label": None, "dx": 0.0, "dy": 0.0}

        def _fig_xy(event):
            return fig.transFigure.inverted().transform((event.x, event.y))

        def _on_press(event):
            for label, axi in wave_axes:
                if event.inaxes is axi:
                    fx, fy = _fig_xy(event)
                    bb = axi.get_position()
                    drag.update(mode="inset", label=label,
                                dx=fx - bb.x0, dy=fy - bb.y0)
                    return

        def _on_motion(event):
            if drag["mode"] == "inset":
                fx, fy = _fig_xy(event)
                for label, axi in wave_axes:
                    if label == drag["label"]:
                        bb = axi.get_position()
                        axi.set_position(
                            [fx - drag["dx"], fy - drag["dy"], bb.width, bb.height])
                        canvas.draw_idle()
                        break

        def _on_release(event):
            if drag["mode"] == "inset":
                for label, axi in wave_axes:
                    if label == drag["label"]:
                        bb = axi.get_position()
                        saved_insets[label] = [bb.x0, bb.y0, bb.width, bb.height]
                        break
                try:
                    with open(insets_pos_file, "w") as f:
                        _json.dump(saved_insets, f)
                except OSError:
                    pass
            drag["mode"] = None

        canvas.mpl_connect("button_press_event", _on_press)
        canvas.mpl_connect("motion_notify_event", _on_motion)
        canvas.mpl_connect("button_release_event", _on_release)

        canvas.draw()
        self._sa_time_drawn = True

    def _draw_sa_gdp_figure(self, parent_frame):
        """SA 合并图2: 建设GDP三线图 + Pearson相关柱状图(draggable inset)."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import csv as _csv, os as _os, re as _re, sqlite3, json as _json
        import numpy as np
        from scipy import stats as _stats
        from matplotlib.ticker import FuncFormatter as _FuncFormatter

        HERE = _os.path.dirname(_os.path.abspath(__file__))
        DB = _os.path.join(HERE, "..", "data", "architecture.db")
        GRID = _os.path.join(HERE, "..", "论文", "maddison_china_25yr_grid.csv")

        def _parse(y):
            y = (y or "").strip()
            m = _re.match(r"^(\d{3,4})(\s*CE)?$", y)
            if m:
                return int(m.group(1)), None
            m = _re.match(r"^(\d{1,2})(st|nd|rd|th)$", y)
            if m:
                return None, int(m.group(1))
            return None, None

        con = sqlite3.connect(DB)
        rows = con.execute(
            "SELECT year_built FROM cases WHERE latitude IS NOT NULL "
            "AND longitude IS NOT NULL AND name != ''").fetchall()
        con.close()
        years_list = []
        for (yb,) in rows:
            y, c = _parse(yb)
            if y is not None:
                years_list.append(y)
            elif c is not None:
                years_list.append((c - 1) * 100 + 50)
        bins = {}
        for y in years_list:
            b = (y // 25) * 25
            bins[b] = bins.get(b, 0) + 1

        gdp_years, gdp_total, gdp_pc = [], [], []
        with open(GRID) as f:
            for row in _csv.DictReader(f):
                gdp_years.append(float(row["quarter_century_midpoint"]))
                gdp_total.append(float(row["gdp_total_trillion_2011intl"]))
                gdp_pc.append(float(row["gdp_per_capita_2011intl"]))

        C_FOUND = "#1a3a5c"
        C_TOTAL = "#c0392b"
        C_PC = "#e67e22"

        fig = Figure(figsize=(12, 7), dpi=100)
        ax = fig.add_subplot(111)
        xs = sorted(bins)
        ax.plot([x + 12.5 for x in xs], [bins[x] for x in xs],
                color=C_FOUND, lw=1.8, zorder=5,
                label="Cases founded per quarter-century")
        ax.set_xlim(0, 2025)
        ax.set_xticks(range(0, 2025, 250))
        ax.set_ylim(0, max(bins.values()) * 1.15)
        ax.set_xlabel("Year (CE)")
        ax.set_ylabel("Cases founded per quarter-century", color=C_FOUND, fontsize=10)
        ax.tick_params(axis="y", colors=C_FOUND)
        ax.spines["top"].set_visible(False)

        ax2 = ax.twinx()
        ax2.plot(gdp_years, gdp_total, color=C_TOTAL, lw=1.4, zorder=4,
                 label="China GDP (trillion 2011 int.-$)")
        ax2.set_ylabel("China GDP (trillion 2011 int.-$)", color=C_TOTAL, fontsize=10)
        ax2.tick_params(axis="y", colors=C_TOTAL)
        ax2.spines["top"].set_visible(False)

        ax3 = ax.twinx()
        ax3.spines["right"].set_position(("outward", 42))
        ax3.plot(gdp_years, gdp_pc, color=C_PC, lw=1.4, zorder=4,
                 label="China GDP per capita (2011 int.-$)")
        ax3.yaxis.set_major_formatter(_FuncFormatter(lambda v, _: f"{v/1000:g}k"))
        ax3.set_ylabel("China GDP per capita (2011 int.-$)", color=C_PC, fontsize=10)
        ax3.tick_params(axis="y", colors=C_PC)
        ax3.spines["top"].set_visible(False)

        lines, labels = [], []
        for a in (ax, ax2, ax3):
            l, lb = a.get_legend_handles_labels()
            lines += l
            labels += lb
        ax.legend(lines, labels, loc="upper left", frameon=False, fontsize=8)

        # ---- Pearson inset (draggable, position persisted) ----
        founding_bins = {}
        for y in years_list:
            b = (y // 25) * 25
            founding_bins[b] = founding_bins.get(b, 0) + 1
        x_founding, x_gdp = [], []
        for gy, gv in zip(gdp_years, gdp_total):
            b = int((gy // 25) * 25)
            x_founding.append(founding_bins.get(b, 0))
            x_gdp.append(gv)

        def _pearson(x, y):
            if len(x) < 3:
                return None, None, None, None, None
            r, p = _stats.pearsonr(x, y)
            n = len(x)
            z = np.arctanh(r)
            se = 1 / np.sqrt(n - 3)
            return r, np.tanh(z - 1.96 * se), np.tanh(z + 1.96 * se), p, n

        def _sig(p):
            if p is None:
                return "n/a"
            if p < 0.001:
                return "***"
            if p < 0.01:
                return "**"
            if p < 0.05:
                return "*"
            return "n.s."

        periods = []
        for label, lo, hi in [
            ("1000–2024\n(Full)", 1000, 2024),
            ("1000–1599\n(Pre-Wave)", 1000, 1600),
            ("1600–1799\n(Wave I)", 1600, 1800),
            ("1800–1949\n(Wave II + Trough)", 1800, 1950),
        ]:
            indices = [i for i, gy in enumerate(gdp_years) if lo <= gy <= hi]
            xf = [x_founding[i] for i in indices]
            xg = [x_gdp[i] for i in indices]
            r, r_lo, r_hi, p, n = _pearson(xf, xg)
            if r is None:
                continue
            periods.append((label, round(r, 4), round(r_lo, 4), round(r_hi, 4), _sig(p)))

        pos_file = _os.path.join(_os.path.dirname(DB), "sa_gdp_pearson_pos.json")
        pearson_pos = [0.58, 0.55, 0.36, 0.34]
        try:
            with open(pos_file) as f:
                p = _json.load(f)
                if isinstance(p, list) and len(p) == 4:
                    pearson_pos = p
        except (OSError, ValueError):
            pass

        axi = fig.add_axes(pearson_pos, facecolor="white")
        labels_p = [p[0] for p in periods]
        r_vals = [p[1] for p in periods]
        ci_lo = [p[1] - p[2] for p in periods]
        ci_hi = [p[3] - p[1] for p in periods]
        sigs = [p[4] for p in periods]
        x = range(len(periods))
        colors = ["#7BA7BC", "#9DC0A8", "#E89F7E", "#BFB5A8"]
        axi.bar(x, r_vals, color=colors, width=0.5, zorder=3,
                yerr=[ci_lo, ci_hi], capsize=5, error_kw={"lw": 1.0, "zorder": 4})
        for i, (r, sig) in enumerate(zip(r_vals, sigs)):
            y_pos = r + ci_hi[i] + 0.05 if r >= 0 else r + ci_hi[i] + 0.10
            axi.text(i, y_pos, sig, ha="center", va="bottom",
                     fontsize=9, fontweight="bold")
        axi.axhline(y=0, color="black", lw=0.6, zorder=1)
        axi.set_xticks(list(x))
        axi.set_xticklabels(labels_p, fontsize=7)
        axi.set_ylabel("Pearson r", fontsize=8)
        axi.set_ylim(-0.85, 1.15)
        axi.spines["top"].set_visible(False)
        axi.spines["right"].set_visible(False)
        axi.grid(axis="y", color="#e0e0e0", lw=0.4, zorder=0)
        axi.set_title("Pearson r: founded vs GDP total", fontsize=8.5)

        # ---- Drag handling: pearson inset ----
        drag = {"active": False, "dx": 0.0, "dy": 0.0}

        def _fig_xy(event):
            return fig.transFigure.inverted().transform((event.x, event.y))

        def _on_press(event):
            if event.inaxes is axi:
                fx, fy = _fig_xy(event)
                bb = axi.get_position()
                drag.update(active=True, dx=fx - bb.x0, dy=fy - bb.y0)

        def _on_motion(event):
            if drag["active"]:
                fx, fy = _fig_xy(event)
                bb = axi.get_position()
                axi.set_position([fx - drag["dx"], fy - drag["dy"],
                                  bb.width, bb.height])
                canvas.draw_idle()

        def _on_release(event):
            if drag["active"]:
                drag["active"] = False
                bb = axi.get_position()
                try:
                    with open(pos_file, "w") as f:
                        _json.dump([bb.x0, bb.y0, bb.width, bb.height], f)
                except OSError:
                    pass

        fig.subplots_adjust(bottom=0.10, top=0.93, left=0.07, right=0.90)

        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, parent_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.mpl_connect("button_press_event", _on_press)
        canvas.mpl_connect("motion_notify_event", _on_motion)
        canvas.mpl_connect("button_release_event", _on_release)
        canvas.draw()
        self._sa_gdp_drawn = True

    # ------------------------------------------------------------------
    #  七渠道分析图 — river plot of the seven diffusion channels (Table 1)
    # ------------------------------------------------------------------

    _CHANNEL_NAME = {
        1: "Imperial administration", 2: "Marriage alliance", 3: "Tributary diplomacy",
        4: "Buddhist transmission", 5: "Migrant & commercial", 6: "Court Chinoiserie",
        7: "Modern cultural diplomacy",
    }
    _CHANNEL_CAPTION = {
        1: "Chùa Dâu · 187 CE · Vietnam", 2: "Por-Bazhyn · 777 · Russia",
        3: "Gyeongbokgung · 1395 · Korea", 4: "Tenryū-ji · 1339 · Japan",
        5: "Cheah Kongsi · 1873 · Malaysia", 6: "Kew Pagoda · 1762 · UK",
        7: "Sun Yat-Sen Garden · 1986 · Canada",
    }
    _CHANNEL_IMG_ID = {1: 980, 2: 820, 3: 1093, 4: 1172, 5: 320, 6: 1320, 7: 345}
    _CHANNEL_COLORS = ["#5B9BD5", "#7FB069", "#F2C14E", "#E8804E", "#D06B9A", "#9B72C2", "#46B3A6"]
    _CHANNEL_DESC = {
        1: ("Pre-1600; Jiaozhi, Lelang, far overland · Buddhist temple, administrative/ritual",
            "Han official: timber-frame halls on rammed-earth platforms, dougong bracket sets, axial courtyard planning"),
        2: ("Han–Tang; steppe empires, great distance · Palace/fortress, religious structures",
            "Tang military-palace: rammed-earth walls, strict orthogonal grid, Chinese roof tiles"),
        3: ("Tang–19th c.; Korea, Ryukyu, Viet Nam · Palace/royal, religious shrine, garden, monument/memorial",
            "Chinese axial plan + local roof curvature and materials"),
        4: ("6th c.–present; E Asia, Mongolia, broadest arc · Buddhist temple",
            "Timber pagoda, Mahavira Hall, mountain gate, south-facing courtyard plan; regional variants: Wayo (Japan), datsan (Mongolia), Korean curved-ridge temple halls"),
        5: ("15th c.–present; SE Asian ports–global diaspora hubs; peaks 1850–1899 and 1980+ · Taoist/ancestral temple, folk shrine, shophouse, commercial, cultural/community centre",
            "Swallowtail ridge, jiannian ceramic decoration, narrow-frontage shophouse with air well, Chinatown paifang gate; Straits shophouse hybrid type"),
        6: ("1600–1799 peak, then paused; European royal seats, aristocratic towns · Garden/park, pavilion/folly, palace/royal",
            "European masonry pagoda (exaggerated stacked eaves), octagonal/hexagonal pavilion; Chinese outline, European materials"),
        7: ("1980–present; global, most evenly spread; capitals, sister cities · Buddhist temple, garden/park, Chinatown gate, cultural centre, museum, commercial",
            "Suzhou-style scholar garden miniature: white walls, grey tiles, Taihu rockeries, moon gate, zigzag bridge; expo pavilion: steel frame + abstracted Chinese roof curve"),
    }

    def _draw_channel_figure(self, parent_frame):
        """七渠道传播河流图: 7 条传播途径的时间密度河流 + 代表性案例图片."""
        for w in parent_frame.winfo_children():
            w.destroy()

        import re as _re
        import textwrap as _textwrap
        import io as _io
        import sqlite3 as _sqlite3
        import numpy as np
        from collections import defaultdict
        from PIL import Image as _PILImage

        # ---- channel classification (faithful to Table 1) ----
        _COUNTRY_MAP = {
            "泰国": "thailand", "日本": "japan", "马来西亚": "malaysia", "越南": "viet nam",
            "韩国": "republic of korea", "美国": "united states", "新加坡": "singapore",
            "印度尼西亚": "indonesia", "澳大利亚": "australia", "法国": "france", "德国": "germany",
            "英国": "united kingdom of great britain and northern ireland",
            "俄罗斯": "russian federation", "加拿大": "canada",
            "朝鲜": "democratic people's republic of korea", "蒙古": "mongolia",
            "菲律宾": "philippines", "荷兰": "netherlands", "瑞典": "sweden", "意大利": "italy",
            "柬埔寨": "cambodia", "新西兰": "new zealand", "丹麦": "denmark", "缅甸": "myanmar",
            "波兰": "poland", "乌克兰": "ukraine", "老挝": "lao people's democratic republic",
            "西班牙": "spain", "毛里求斯": "mauritius", "比利时": "belgium", "奥地利": "austria",
            "吉尔吉斯斯坦": "kyrgyzstan", "古巴": "cuba", "印度": "india", "不丹": "bhutan",
            "赞比亚": "zambia", "白俄罗斯": "belarus", "瑞士": "switzerland",
            "津巴布韦": "zimbabwe", "捷克": "czechia", "巴西": "brazil", "尼日利亚": "nigeria",
            "墨西哥": "mexico", "哈萨克斯坦": "kazakhstan", "南非": "south africa",
            "匈牙利": "hungary", "伊朗": "iran (islamic republic of)",
            "马达加斯加": "madagascar", "马拉维": "malawi", "阿根廷": "argentina",
            "贝宁": "benin", "葡萄牙": "portugal", "纳米比亚": "namibia", "秘鲁": "peru",
            "爱尔兰": "ireland", "智利": "chile", "文莱": "brunei darussalam",
            "摩尔多瓦共和国": "republic of moldova", "拉脱维亚": "latvia", "巴拿马": "panama",
            "尼泊尔": "nepal", "密克罗尼西亚联邦": "micronesia", "塞舌尔": "seychelles",
            "基里巴斯": "kiribati", "圭亚那": "guyana", "土耳其": "turkiye",
            "厄瓜多尔": "ecuador", "博茨瓦纳": "botswana", "东帝汶": "timor-leste",
            "united kingdom": "united kingdom of great britain and northern ireland",
            "south korea": "republic of korea", "north korea": "democratic people's republic of korea",
            "russia": "russian federation", "usa": "united states",
            "laos": "lao people's democratic republic", "brunei": "brunei darussalam",
            "iran": "iran (islamic republic of)", "kyrgyz republic": "kyrgyzstan",
            "moldova": "republic of moldova", "turkey": "turkiye", "east timor": "timor-leste",
            "czech republic": "czechia",
        }

        def _norm_country(c):
            c = (c or "").strip().lower()
            return _COUNTRY_MAP.get(c, c)

        def _parse_year(y):
            y = (y or "").strip()
            m = _re.match(r"^(\d{3,4})(\s*CE)?$", y)
            if m:
                return int(m.group(1))
            m = _re.match(r"^(\d{1,2})(st|nd|rd|th)$", y)
            if m:
                return (int(m.group(1)) - 1) * 100 + 50
            return None

        _ASIA_SE = ["泰国", "马来西亚", "越南", "新加坡", "印度尼西亚", "菲律宾", "柬埔寨", "缅甸",
                    "老挝", "文莱", "东帝汶", "thailand", "malaysia", "viet nam", "singapore",
                    "indonesia", "philippines", "cambodia", "myanmar", "laos", "brunei", "timor-leste"]
        _ASIA_E = ["日本", "韩国", "朝鲜", "蒙古", "japan", "south korea", "north korea",
                   "republic of korea", "democratic people's republic of korea", "mongolia"]

        def _get_region(country):
            c = (country or "").strip()
            nc = _norm_country(c)
            if c in _ASIA_SE or nc in [_norm_country(x) for x in _ASIA_SE]:
                return "SE Asia"
            if c in _ASIA_E or nc in [_norm_country(x) for x in _ASIA_E]:
                return "E Asia"
            if nc in ["india", "bhutan", "nepal", "iran (islamic republic of)",
                      "kazakhstan", "kyrgyzstan", "turkiye"]:
                return "S/C/W Asia"
            if any(k in nc for k in ["united kingdom", "france", "germany", "netherlands",
                                     "sweden", "italy", "denmark", "poland", "ukraine",
                                     "austria", "belgium", "spain", "hungary", "czechia",
                                     "switzerland", "belarus", "latvia", "moldova", "ireland",
                                     "portugal", "russian"]):
                return "Europe"
            if any(k in nc for k in ["united states", "canada", "cuba", "mexico", "brazil",
                                     "ecuador", "guyana", "panama", "chile", "peru", "argentina"]):
                return "Americas"
            if any(k in nc for k in ["australia", "new zealand", "kiribati", "micronesia"]):
                return "Oceania"
            return "Africa"

        def _categorize_bt(bt):
            bt = (bt or "").strip().lower()
            if not bt:
                return "other"
            if any(k in bt for k in ["buddhist", "佛教", "佛寺", "佛殿", "僧院", "寺院", "禅宗", "禅修",
                                     "修道院", "datsan", "khurul", "tendai", "obaku", "zen",
                                     "monastery", "pagoda", "vihara", "stupa", "wat", "喇嘛", "藏传",
                                     "nunnery"]):
                return "Buddhist temple"
            if any(k in bt for k in ["taoist", "道教", "妈祖", "关帝", "关羽", "天后", "祠", "祖先",
                                     "宗祠", "宗庙", "宗乡", "tri dharma", "pilgrimage"]):
                return "Taoist/ancestral"
            if any(k in bt for k in ["temple", "寺庙", "庙宇", "庙", "shrine", "礼拜", "儒家", "confuci",
                                     "孔子", "shinto", "church", "宗教"]):
                return "Other religion"
            if any(k in bt for k in ["mosque", "清真寺"]):
                return "Mosque"
            if any(k in bt for k in ["牌坊", "牌楼", "入口", "gateway", "paifang", "archway",
                                     "chinatown entrance", "ceremonial gate", "commemorative"]):
                return "Chinatown gate"
            if any(k in bt for k in ["园林", "花园", "garden", "botanical", "stroll garden",
                                     "landscape garden", "公共园林", "公园", "park", "public park"]):
                return "Garden/park"
            if any(k in bt for k in ["亭", "pavilion", "folly", "观赏", "观景", "景观", "ornamental",
                                     "decorative", "pleasure", "tea house"]):
                return "Pavilion/folly"
            if any(k in bt for k in ["宫殿", "皇宫", "王宫", "皇家", "royal", "palace", "imperial",
                                     "贵族", "noble", "aristocratic", "宅邸", "府邸"]):
                return "Palace/royal"
            if any(k in bt for k in ["博物馆", "museum", "展览", "exhibition", "gallery", "展示"]):
                return "Museum/exhibition"
            if any(k in bt for k in ["文化中心", "文化交流", "文化展示", "community center",
                                     "community centre", "cultural center", "cultural exchange",
                                     "assembly hall", "会馆", "clan association", "宗亲"]):
                return "Cultural/community"
            if any(k in bt for k in ["商业", "市场", "market", "shop", "restaurant", "hotel", "酒店",
                                     "餐饮", "mixed-use", "commercial", "retail", "shophouse"]):
                return "Commercial"
            if any(k in bt for k in ["住宅", "居住", "residential", "house", "dormitory", "guesthouse"]):
                return "Residential"
            if any(k in bt for k in ["纪念碑", "纪念", "monument", "memorial", "landmark", "地标",
                                     "commemoration", "cemetery", "墓地"]):
                return "Monument/memorial"
            if any(k in bt for k in ["学校", "教育", "school", "university", "college", "academy",
                                     "学院", "书院", "孤儿院"]):
                return "Education"
            return "other"

        _MIGRANT_KEYS = ['会馆', 'assembly hall', 'guild', 'kongsi', 'clan', '宗亲', '乡会', '同乡会',
                         '會館', '关帝', '关羽', '妈祖', '天后', '天妃', 'mazu', 'tianhou', 'guan yu',
                         'guandi', 'kwan tai', 'kwan tee', 'kwan ti', 'kwan kung', 'kwan yin',
                         '宗祠', '祠堂', 'ancestral hall', 'joss house', '天后宫', '唐人街',
                         'chinatown', '中华街', '南京町', '華埠', '华埠', 'china town', '牌坊', '牌楼',
                         'paifang', 'pai fang', 'chinatown gate', 'ceremonial gate', '店屋',
                         'shophouse', 'dungan', '东干', '回民', 'hoi quan', 'hội quán', 'teahouse',
                         '茶楼']
        _SHINTO_KEYS = ['shinto', '神社', 'jinja', 'taisha', '大社', '神宫']
        _SEASIA_CN = {'thailand', 'malaysia', 'singapore', 'indonesia', 'philippines', 'cambodia',
                      'myanmar', 'lao people\'s democratic republic', 'brunei darussalam',
                      'timor-leste'}

        def _refine(name, name_en, country, bt, base_func):
            s = ((bt or '') + ' ' + (name or '') + ' ' + (name_en or '')).lower()
            nc = _norm_country(country)
            if any(k in s for k in _SHINTO_KEYS):
                return 'Shinto'
            if any(k in s for k in _MIGRANT_KEYS):
                return 'Migrant'
            nm = (name or '').strip()
            ne = (name_en or '').strip()
            if _re.match(r'^wat\b', ne, _re.I) or _re.match(r'^wat\b', nm, _re.I):
                return 'Buddhist temple'
            if base_func in ('Other religion', 'other'):
                if _re.search(r'(寺|庵|刹|chùa|chua)$', nm, _re.I) or _re.search(r'-(ji|dera|in|sa|an)\b', ne, _re.I):
                    return 'Buddhist temple'
            if nc == 'mongolia' and base_func in ('Other religion', 'other'):
                return 'Buddhist temple'
            if nc in _SEASIA_CN and base_func in ('Other religion', 'Taoist/ancestral',
                                                  'Cultural/community', 'other'):
                return 'Migrant'
            return base_func

        def _classify(name, name_en, country, city, year, bt):
            nc = _norm_country(country)
            region = _get_region(country)
            func = _refine(name, name_en, country, bt, _categorize_bt(bt))
            s = ((bt or '') + ' ' + (name or '')).lower()

            if year is not None and year < 1000 and nc in ['russian federation', 'mongolia',
                                                           'kazakhstan', 'kyrgyzstan']:
                if 'palace' in s or 'fortress' in s or func == 'Palace/royal':
                    return 2

            if nc == 'iran (islamic republic of)':
                return 3

            if func == 'Buddhist temple':
                if nc == 'viet nam' and year is not None and year < 1600:
                    return 1
                return 4

            if nc == 'mongolia':
                return 4

            if func == 'Migrant':
                return 5

            if nc == 'viet nam' and year is not None and year < 1600:
                return 1

            if year is not None and year >= 1980:
                return 7

            if region == 'SE Asia' and nc != 'viet nam' and func == 'Palace/royal':
                return 3

            if region == 'SE Asia' and nc != 'viet nam' and func != 'Palace/royal':
                return 5

            if nc in ['republic of korea', "democratic people's republic of korea"]:
                return 3
            if nc == 'japan':
                return 3
            if nc == 'viet nam':
                return 3

            if region == 'Europe' and func in ['Garden/park', 'Pavilion/folly', 'Palace/royal']:
                return 6

            if region in ['Americas', 'Oceania', 'Africa']:
                return 5
            if region == 'Europe':
                return 6
            if region == 'S/C/W Asia':
                return 4
            return 7

        # ---- load cases & assign channels ----
        conn = _sqlite3.connect(db.DB_PATH)
        rows = conn.execute(
            "SELECT name, name_en, country, city, year_built, building_type FROM cases "
            "WHERE latitude IS NOT NULL AND longitude IS NOT NULL AND name != ''").fetchall()
        conn.close()

        years_by_ch = defaultdict(list)
        n_by_ch = defaultdict(int)
        for name, name_en, country, city, yb, bt in rows:
            y = _parse_year(yb)
            ch = _classify(name, name_en, country, city, y, bt)
            n_by_ch[ch] += 1
            if y is not None:
                years_by_ch[ch].append(y)

        # ---- density smoothing ----
        X = np.linspace(100, 2024, 800)
        SIGMA = 45.0

        def _smooth(years):
            d = np.zeros_like(X)
            for y in years:
                d += np.exp(-0.5 * ((X - y) / SIGMA) ** 2)
            return d

        density = {}
        for ch in range(1, 8):
            ys = years_by_ch.get(ch, [])
            d = _smooth(ys) if ys else np.zeros_like(X)
            dmax = d.max()
            density[ch] = d / dmax if dmax > 0 else d

        # ---- figure ----
        fig = Figure(figsize=(24, 16), dpi=100)
        fig._intended_size_inches = (24.0, 16.0)
        fig_w, fig_h = fig.get_size_inches()

        AX_LEFT, AX_BOT, AX_W, AX_H = 0.05, 0.12, 0.72, 0.84
        ax = fig.add_axes([AX_LEFT, AX_BOT, AX_W, AX_H])
        GAP = 6.0
        YLO, YHI = -1.5, 6 * GAP + 2.2
        SPAN = YHI - YLO

        def _yfrac(y):
            return AX_BOT + (y - YLO) / SPAN * AX_H

        for ch in range(1, 8):
            ycenter = (7 - ch) * GAP
            d = density[ch]
            half = 0.12 + 0.85 * d
            ax.fill_between(X, ycenter - half, ycenter + half,
                            color=self._CHANNEL_COLORS[ch - 1], alpha=0.85, linewidth=0)
            ax.plot(X, np.full_like(X, ycenter), color=self._CHANNEL_COLORS[ch - 1],
                    lw=0.5, alpha=0.45)
            ax.text(108, ycenter + 1.70, f"{ch}  {self._CHANNEL_NAME[ch]}",
                    fontsize=8.5, fontweight="bold", color="#111111",
                    va="bottom", ha="left")
            sig, form = self._CHANNEL_DESC[ch]
            wrapped = _textwrap.fill(sig, width=95) + "\n" + _textwrap.fill(form, width=95)
            ax.text(108, ycenter + 1.55, wrapped, fontsize=7.0, color="#555",
                    va="top", ha="left")

        ax.set_xlim(100, 2024)
        ax.set_ylim(YLO, YHI)
        ax.set_yticks([])
        for s in ("top", "right", "left"):
            ax.spines[s].set_visible(False)
        ax.xaxis.set_ticks_position("bottom")
        ax.tick_params(axis="x", labelsize=8.5)
        ax.set_xticks([187, 500, 1000, 1500, 1800, 2024])
        ax.set_xticklabels(["187", "500", "1000", "1500", "1800", "2024"])
        ax.set_xlabel("Year (CE)", fontsize=11, labelpad=6)

        # ---- representative image column ----
        IMG_CX = 0.86
        IMG_W_IN = 1.5
        IMG_W = IMG_W_IN / fig_w
        for ch in range(1, 8):
            ycenter = (7 - ch) * GAP
            yc = _yfrac(ycenter)
            arr = None
            data, _ = self._load_image_data(f"dbimg://{self._CHANNEL_IMG_ID[ch]}")
            if data:
                try:
                    im = _PILImage.open(_io.BytesIO(data)).convert("RGB")
                    if ch == 3:
                        w, h = im.size
                        im = im.crop((0, 0, w, int(h * 3 / 4)))
                    arr = np.asarray(im)
                except Exception:
                    arr = None
            if arr is not None:
                h_px, w_px = arr.shape[:2]
                aspect = w_px / h_px
                H_in = IMG_W_IN / aspect
            else:
                H_in = IMG_W_IN
            H = H_in / fig_h
            iax = fig.add_axes([IMG_CX - IMG_W / 2, yc - H / 2, IMG_W, H])
            iax.axis("off")
            if arr is not None:
                iax.imshow(arr)
            else:
                iax.text(0.5, 0.5, "[no image]", ha="center", va="center",
                         fontsize=7, color="#999", transform=iax.transAxes)
            for s in iax.spines.values():
                s.set_visible(True)
                s.set_color("#cccccc")
                s.set_linewidth(0.5)
            fig.text(IMG_CX, yc - H / 2 - 0.003, self._CHANNEL_CAPTION[ch],
                     fontsize=7.0, color="#444", ha="center", va="top", style="italic")

        self._current_fig = fig

        # ---- scrollable full-size embedding (figure is print-sized) ----
        container = ttk.Frame(parent_frame)
        container.pack(fill="both", expand=True)

        host = tk.Canvas(container, highlightthickness=0, bg="#fafafa")
        hbar = ttk.Scrollbar(container, orient="horizontal", command=host.xview)
        vbar = ttk.Scrollbar(container, orient="vertical", command=host.yview)
        host.configure(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        hbar.pack(side="bottom", fill="x")
        vbar.pack(side="right", fill="y")
        host.pack(side="left", fill="both", expand=True)

        mpl_canvas = FigureCanvasTkAgg(fig, master=host)
        mpl_widget = mpl_canvas.get_tk_widget()
        mpl_widget.config(width=int(fig_w * fig.get_dpi()), height=int(fig_h * fig.get_dpi()))
        host.create_window((0, 0), window=mpl_widget, anchor="nw")
        mpl_widget.bind("<Configure>",
                        lambda e: host.configure(scrollregion=host.bbox("all")))
        mpl_canvas.draw()

        def _on_wheel(event):
            if event.state & 0x1:  # Shift -> horizontal scroll
                host.xview_scroll(int(-event.delta / 120), "units")
            else:
                host.yview_scroll(int(-event.delta / 120), "units")

        host.bind("<Enter>", lambda e: host.bind_all("<MouseWheel>", _on_wheel))
        host.bind("<Leave>", lambda e: host.unbind_all("<MouseWheel>"))

        self._channel_drawn = True

    def _load_image_data(self, src):
        """Re-use image loading from image_components."""
        import urllib.request
        import ssl
        import os

        if src.startswith("dbimg://"):
            try:
                img_id = int(src.replace("dbimg://", ""))
                record = db.get_case_image(img_id)
                if record and record.get("image_data"):
                    return record["image_data"], True
            except Exception:
                pass
            return None, False

        # local file path
        if os.path.exists(src):
            with open(src, "rb") as f:
                return f.read(), True
        return None, False

    def _get_case_name(self, db_id):
        """Look up case name from database."""
        try:
            row = db.get_case_by_id(db_id)
            if row:
                return row.get("name_en", "") or row.get("name", "") or ""
        except Exception:
            pass
        return ""

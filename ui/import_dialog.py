import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd

import database as db


def _get_all_fields():
    """Return all fields including dynamic ones: [(field_name, label), ...]"""
    fields = [(fn, fl) for (fn, fl, _) in db.RAW_FIELDS]
    for fn in db.get_raw_case_extra_columns():
        fields.append((fn, fn))  # use field name as label for dynamic fields
    return fields


def _fuzzy_match(excel_col):
    """Try to match an Excel column name to a DB field.
    Returns (field_name, label, confidence) or (None, None, 0).
    """
    col_clean = excel_col.strip().lower().replace(" ", "").replace("（", "(").replace("）", ")")
    all_fields = _get_all_fields()

    # 1) Exact match on label
    for (fname, flabel) in all_fields:
        if flabel.strip() == excel_col.strip():
            return (fname, flabel, 1.0)
    # 2) Exact match on field name (case-insensitive)
    for (fname, flabel) in all_fields:
        if fname.lower() == col_clean:
            return (fname, flabel, 1.0)

    # 3) Field name contained in Excel col
    for (fname, flabel) in all_fields:
        if fname.lower() in col_clean:
            return (fname, flabel, 0.9)

    # 4) Excel col contained in label or label contained in Excel col
    # Only for meaningful column names (3+ chars) to avoid spurious matches
    if len(col_clean) >= 3:
        for (fname, flabel) in all_fields:
            lbl = flabel.strip().lower().replace(" ", "")
            if len(col_clean) <= len(lbl) and col_clean in lbl:
                return (fname, flabel, 0.8)
            if len(lbl) < len(col_clean) and lbl in col_clean:
                return (fname, flabel, 0.8)

    # 5) Image column detection (before keyword map, since "图片" is special)
    if re.search(r"图片\d*|图像\d*|image\d*|照片\d*|photo\d*", col_clean, re.IGNORECASE):
        return ("__IMAGE__", "图片(特殊处理)", 1.0)

    # 6) Keyword heuristic matching
    keyword_map = {
        "国家": "country",
        "名称": "name",
        "建造": "construction_year",
        "拆除": "demolition_year",
        "修复": "restoration_year",
        "数量": "quantity",
        "功能": "function",
        "设计": "designer",
        "位置": "city",
        "检索时间": "search_time",
        "检索词": "search_keywords",
        "备注1": "authenticity_notes",
        "备注": "authenticity_notes",
        "传播": "spread_method",
        "连接": "source_urls",
        "链接": "source_urls",
        "名字": "name_official",
        "城市": "city",
        "时间": "search_time",
        "图片链接": "image_urls",
        "建筑名称": "name",
        "英文": "name_en",
        "官方": "name_official",
        "经度": "longitude",
        "纬度": "latitude",
        "来源": "source_urls",
    }
    for kw, fname in keyword_map.items():
        if kw in excel_col:
            for (fn, fl) in all_fields:
                if fn == fname:
                    return (fn, fl, 0.7)

    return (None, None, 0)


def _extract_images_from_excel(filepath):
    """Extract embedded images from an Excel file.
    Handles both standard Excel images (via openpyxl) and WPS DISPIMG images.
    Returns dict: {row_index: [image_bytes, ...]}
    """
    import zipfile
    import xml.etree.ElementTree as ET

    images_by_row = {}

    # --- Method 1: Standard openpyxl extraction ---
    try:
        from openpyxl import load_workbook
        wb = load_workbook(filepath, data_only=True)
        for ws in wb.worksheets:
            if not hasattr(ws, "_images") or not ws._images:
                continue
            for img in ws._images:
                try:
                    row_idx = img.anchor._from.row
                    img_data = img._data()
                    if row_idx not in images_by_row:
                        images_by_row[row_idx] = []
                    images_by_row[row_idx].append(img_data)
                except Exception:
                    continue
        wb.close()
    except Exception:
        pass

    # --- Method 2: WPS DISPIMG extraction from zip ---
    try:
        with zipfile.ZipFile(filepath, "r") as z:
            # Check for WPS cellimages
            if "xl/cellimages.xml" not in z.namelist():
                return images_by_row

            # Parse rels: rId -> media path
            rels_tree = ET.parse(z.open("xl/_rels/cellimages.xml.rels"))
            ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
            rId_to_target = {}
            for rel in rels_tree.findall(".//r:Relationship", ns):
                rid = rel.get("Id")
                target = rel.get("Target")
                if rid and target:
                    rId_to_target[rid] = target

            # Parse cellimages.xml: image ID -> rId
            cellimages_tree = ET.parse(z.open("xl/cellimages.xml"))
            ns2 = {
                "etc": "http://www.wps.cn/officeDocument/2017/etCustomData",
                "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            }
            img_id_to_path = {}
            for cell_img in cellimages_tree.findall(".//etc:cellImage", ns2):
                nv_pr = cell_img.find(".//xdr:cNvPr", ns2)
                blip = cell_img.find(".//a:blip", ns2)
                if nv_pr is not None and blip is not None:
                    img_name = nv_pr.get("name", "")
                    r_embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", "")
                    if img_name and r_embed and r_embed in rId_to_target:
                        img_id_to_path[img_name] = rId_to_target[r_embed]

            if not img_id_to_path:
                return images_by_row

            # Read the sheet data to find DISPIMG formulas and their row positions
            try:
                import pandas as pd
                df_raw = pd.read_excel(filepath, engine="openpyxl", header=None)
            except Exception:
                return images_by_row

            for row_idx, row in df_raw.iterrows():
                for col_val in row:
                    if isinstance(col_val, str) and "DISPIMG" in col_val:
                        match = re.search(r'DISPIMG\("([^"]+)"', col_val)
                        if match:
                            img_id = match.group(1)
                            if img_id in img_id_to_path:
                                media_path = img_id_to_path[img_id]
                                # media_path is like "media/image1.jpeg"
                                full_path = "xl/" + media_path
                                try:
                                    img_bytes = z.read(full_path)
                                    if row_idx not in images_by_row:
                                        images_by_row[row_idx] = []
                                    images_by_row[row_idx].append(img_bytes)
                                except Exception:
                                    continue
    except Exception:
        pass

    return images_by_row


def _detect_image_columns(df):
    """Detect columns that contain DISPIMG formulas or image references.
    Returns list of column names.
    """
    image_cols = []
    for col in df.columns:
        sample = df[col].dropna().head(10)
        if sample.empty:
            continue
        dispimg_count = sum(1 for v in sample if isinstance(v, str) and "DISPIMG" in v)
        if dispimg_count > 0:
            image_cols.append(col)
    return image_cols


class ImportDialog:
    """Interactive Excel import dialog for raw_cases."""

    def __init__(self, parent, on_imported=None):
        self.parent = parent
        self.on_imported = on_imported
        self.df = None
        self.filepath = None
        self.image_data = {}  # row_idx -> [image_bytes, ...]
        self.image_col_names = []  # Excel columns marked as image columns

        self.win = tk.Toplevel(parent)
        self.win.title("导入数据到核验队列")
        self.win.geometry("900x750")
        self.win.transient(parent)

        self._build_ui()

    def _build_ui(self):
        # Step 1: File selection
        file_frame = ttk.LabelFrame(self.win, text=" 选择文件 ", padding=5)
        file_frame.pack(fill="x", padx=10, pady=(10, 5))

        row1 = ttk.Frame(file_frame)
        row1.pack(fill="x")
        ttk.Button(row1, text="选择 Excel 文件", command=self._select_file).pack(side="left", padx=2)
        self._file_label = ttk.Label(row1, text="未选择文件", foreground="#999")
        self._file_label.pack(side="left", padx=10)

        self._preview_label = ttk.Label(row1, text="", foreground="#666")
        self._preview_label.pack(side="right", padx=5)

        # Step 2: Field mapping (scrollable)
        map_frame = ttk.LabelFrame(self.win, text=" 字段映射 ", padding=5)
        map_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self._map_canvas = tk.Canvas(map_frame, highlightthickness=0)
        map_scroll = ttk.Scrollbar(map_frame, orient="vertical", command=self._map_canvas.yview)
        self._map_inner = ttk.Frame(self._map_canvas)

        self._map_inner.bind("<Configure>",
            lambda e: self._map_canvas.configure(scrollregion=self._map_canvas.bbox("all")))
        canvas_window = self._map_canvas.create_window((0, 0), window=self._map_inner, anchor="nw")

        def _on_canvas_configure(event):
            self._map_canvas.itemconfig(canvas_window, width=event.width)
        self._map_canvas.bind("<Configure>", _on_canvas_configure)
        self._map_canvas.configure(yscrollcommand=map_scroll.set)

        self._map_canvas.pack(side="left", fill="both", expand=True)
        map_scroll.pack(side="right", fill="y")

        # Mapping header
        hdr = ttk.Frame(self._map_inner)
        hdr.pack(fill="x", pady=(0, 3))
        ttk.Label(hdr, text="Excel列名", width=20, font=("", 9, "bold")).pack(side="left", padx=2)
        ttk.Label(hdr, text="示例值", width=28, font=("", 9, "bold")).pack(side="left", padx=2)
        ttk.Label(hdr, text="→ 映射到DB字段", width=22, font=("", 9, "bold")).pack(side="left", padx=2)
        ttk.Label(hdr, text="状态", width=8, font=("", 9, "bold")).pack(side="left", padx=2)

        self._mapping_rows_container = ttk.Frame(self._map_inner)
        self._mapping_rows_container.pack(fill="x")

        # Step 3: Options
        opt_frame = ttk.LabelFrame(self.win, text=" 导入选项 ", padding=5)
        opt_frame.pack(fill="x", padx=10, pady=5)

        self._update_existing = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text="更新已有案例（按名称匹配，不创建重复）",
                        variable=self._update_existing).pack(side="left", padx=5)

        # Step 4: Import button
        btn_frame = ttk.Frame(self.win)
        btn_frame.pack(fill="x", padx=10, pady=10)

        self._import_btn = ttk.Button(btn_frame, text="开始导入", command=self._do_import, state="disabled")
        self._import_btn.pack(side="right", padx=3)
        ttk.Button(btn_frame, text="取消", command=self.win.destroy).pack(side="right", padx=3)

        self._status_label = ttk.Label(btn_frame, text="", foreground="#666")
        self._status_label.pack(side="left", padx=5)

    def _select_file(self):
        filepath = filedialog.askopenfilename(
            title="选择要导入的Excel文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        self.filepath = filepath
        self._file_label.config(text=os.path.basename(filepath), foreground="#000")

        try:
            if filepath.endswith(".xls") and not filepath.endswith(".xlsx"):
                self.df = pd.read_excel(filepath, engine="xlrd")
            else:
                self.df = pd.read_excel(filepath, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取文件：{e}")
            return

        if self.df.empty:
            messagebox.showwarning("无数据", "文件中没有数据。")
            return

        self._preview_label.config(
            text=f"共 {len(self.df)} 行 × {len(self.df.columns)} 列"
        )

        # Detect image columns
        self.image_col_names = _detect_image_columns(self.df)

        # Extract embedded images from Excel
        self.image_data = _extract_images_from_excel(filepath)
        if self.image_data:
            self._status_label.config(
                text=f"检测到 {sum(len(v) for v in self.image_data.values())} 张嵌入图片"
            )

        self._build_mapping_rows()
        self._import_btn.config(state="normal")

    def _build_mapping_rows(self):
        # Clear existing
        for w in self._mapping_rows_container.winfo_children():
            w.destroy()

        self._mapping_widgets = {}  # excel_col -> {"var": StringVar, "label": status_label}

        for excel_col in self.df.columns:
            row = ttk.Frame(self._mapping_rows_container)
            row.pack(fill="x", pady=1)

            # Column name
            ttk.Label(row, text=excel_col, width=20, anchor="w").pack(side="left", padx=2)

            # Sample value
            sample_vals = self.df[excel_col].dropna().head(2)
            sample_str = ""
            if len(sample_vals) > 0:
                val = str(sample_vals.iloc[0])
                if len(val) > 50:
                    val = val[:47] + "..."
                sample_str = val

            ttk.Label(row, text=sample_str, width=28, anchor="w",
                      foreground="#555").pack(side="left", padx=2)

            # DB field combobox
            combo = self._build_choices_combo(row, excel_col)

            # Update status when user changes selection
            def _on_change(event, col=excel_col):
                self._update_mapping_status(col)

            combo.bind("<<ComboboxSelected>>", _on_change)

        # Check for duplicate mappings
        self._check_duplicates()

    def _get_choices(self):
        """Build the current list of combobox choices, including dynamic fields from DB."""
        choices = ["-- 跳过此列 --"]
        skip_fields = {"id", "created_at", "updated_at"}

        # Start with known fields from RAW_FIELDS
        seen = set()
        for (fn, fl, _) in db.RAW_FIELDS:
            if fn in skip_fields:
                continue
            choices.append(f"{fl} ({fn})")
            seen.add(fn)

        # Also discover any columns in the actual DB table that aren't in RAW_FIELDS
        extra = db.get_raw_case_extra_columns()
        for fn in extra:
            if fn in skip_fields or fn in seen:
                continue
            choices.append(f"{fn} ({fn})")
            seen.add(fn)

        choices.append("-- 经纬度 (拆分为latitude+longitude) --")
        choices.append("-- 图片列 --")
        choices.append("-- 创建新字段... --")
        choices.append("-- 管理自定义字段... --")
        return choices

    def _build_choices_combo(self, parent, excel_col):
        """Create and return the combobox for a single Excel column."""
        choices = self._get_choices()
        var = tk.StringVar()
        combo = ttk.Combobox(parent, textvariable=var, values=choices,
                             state="readonly", width=26)
        combo.pack(side="left", padx=2)

        # Auto-match
        if excel_col in self.image_col_names:
            var.set("-- 图片列 --")
            status_text = "图片"
            status_color = "#1976d2"
        else:
            fname, flabel, confidence = _fuzzy_match(excel_col)
            if confidence >= 1.0:
                var.set(f"{flabel} ({fname})")
                status_text = "已匹配"
                status_color = "#2e7d32"
            elif confidence >= 0.8:
                var.set(f"{flabel} ({fname})")
                status_text = "推测"
                status_color = "#e65100"
            elif confidence >= 0.6:
                var.set(f"{flabel} ({fname})")
                status_text = "猜测"
                status_color = "#f57c00"
            else:
                var.set("-- 跳过此列 --")
                status_text = "待处理"
                status_color = "#c62828"

        status_label = ttk.Label(parent, text=status_text, width=8,
                                 foreground=status_color, font=("", 9, "bold"))
        status_label.pack(side="left", padx=2)

        self._mapping_widgets[excel_col] = {"var": var, "status": status_label, "combo": combo}
        return combo

    def _refresh_all_combos(self):
        """Rebuild choices for all comboboxes after a new field is added."""
        choices = self._get_choices()
        for excel_col, w in self._mapping_widgets.items():
            current = w["var"].get()
            w["combo"]["values"] = choices

    def _update_mapping_status(self, excel_col):
        w = self._mapping_widgets.get(excel_col)
        if not w:
            return
        val = w["var"].get()
        if val == "-- 跳过此列 --":
            w["status"].config(text="跳过", foreground="#999")
        elif val == "-- 图片列 --":
            w["status"].config(text="图片", foreground="#1976d2")
        elif val == "-- 经纬度 (拆分为latitude+longitude) --":
            w["status"].config(text="已映射", foreground="#2e7d32")
        elif val == "-- 创建新字段... --":
            self._create_new_field(excel_col)
        elif val == "-- 管理自定义字段... --":
            self._manage_fields(excel_col)
        elif val:
            w["status"].config(text="已映射", foreground="#2e7d32")
        self._check_duplicates()

    def _check_duplicates(self):
        """Check for duplicate DB field mappings and highlight them."""
        skip_values = {"-- 跳过此列 --", "-- 图片列 --",
                       "-- 创建新字段... --", "-- 管理自定义字段... --",
                       "-- 经纬度 (拆分为latitude+longitude) --"}
        field_counts = {}
        for excel_col, w in self._mapping_widgets.items():
            val = w["var"].get()
            if val and val not in skip_values:
                match = re.search(r"\((\w+)\)$", val)
                if match:
                    fname = match.group(1)
                    field_counts[fname] = field_counts.get(fname, 0) + 1

        for excel_col, w in self._mapping_widgets.items():
            val = w["var"].get()
            if val and val not in skip_values:
                match = re.search(r"\((\w+)\)$", val)
                if match and field_counts.get(match.group(1), 0) > 1:
                    w["status"].config(text="重复!", foreground="#d32f2f")
                elif match:
                    w["status"].config(text="已映射", foreground="#2e7d32")

    def _parse_mapping(self):
        """Parse the current mapping into a dict: excel_col -> db_field_name.
        Returns (mapping, unmapped_cols, image_cols).
        """
        mapping = {}
        unmapped = []
        image_cols = []
        for excel_col, w in self._mapping_widgets.items():
            val = w["var"].get()
            if not val or val == "-- 跳过此列 --":
                unmapped.append(excel_col)
                mapping[excel_col] = None
            elif val == "-- 图片列 --":
                image_cols.append(excel_col)
                mapping[excel_col] = None
            elif val == "-- 经纬度 (拆分为latitude+longitude) --":
                mapping[excel_col] = "__COORDINATES__"
            else:
                match = re.search(r"\((\w+)\)$", val)
                if match:
                    mapping[excel_col] = match.group(1)
                else:
                    unmapped.append(excel_col)
                    mapping[excel_col] = None
        return mapping, unmapped, image_cols

    def _create_new_field(self, excel_col):
        """Open a dialog to create a new database field for an unmatched Excel column."""
        w = self._mapping_widgets.get(excel_col)
        if not w:
            return

        # Small popup for field creation
        popup = tk.Toplevel(self.win)
        popup.title("创建新字段")
        popup.geometry("400x200")
        popup.transient(self.win)
        popup.grab_set()

        ttk.Label(popup, text=f"为Excel列 \"{excel_col}\" 创建新的数据库字段",
                  font=("", 10)).pack(padx=15, pady=(15, 10))

        # Field name (English)
        f1 = ttk.Frame(popup)
        f1.pack(fill="x", padx=15, pady=3)
        ttk.Label(f1, text="字段名(英文):", width=13).pack(side="left")
        name_var = tk.StringVar()
        # Auto-generate a field name from the column
        suggested = excel_col.strip().lower().replace(" ", "_").replace("（", "").replace("）", "")
        name_var.set(suggested)
        ttk.Entry(f1, textvariable=name_var, width=25).pack(side="left", padx=3)

        # Chinese label
        f2 = ttk.Frame(popup)
        f2.pack(fill="x", padx=15, pady=3)
        ttk.Label(f2, text="中文标签:", width=13).pack(side="left")
        label_var = tk.StringVar(value=excel_col.strip())
        ttk.Entry(f2, textvariable=label_var, width=25).pack(side="left", padx=3)

        # Field type
        f3 = ttk.Frame(popup)
        f3.pack(fill="x", padx=15, pady=3)
        ttk.Label(f3, text="字段类型:", width=13).pack(side="left")
        type_var = tk.StringVar(value="TEXT")
        type_combo = ttk.Combobox(f3, textvariable=type_var, values=["TEXT", "INTEGER", "REAL"],
                                  state="readonly", width=23)
        type_combo.pack(side="left", padx=3)

        error_label = ttk.Label(popup, text="", foreground="#c62828")
        error_label.pack(pady=(5, 0))

        def _do_create():
            fname = name_var.get().strip()
            flabel = label_var.get().strip()
            ftype = type_var.get().strip()
            if not fname:
                error_label.config(text="字段名不能为空")
                return
            if not re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*$", fname):
                error_label.config(text="字段名只能包含英文字母、数字和下划线")
                return
            if not flabel:
                flabel = excel_col.strip()

            # Check if field already exists
            if fname in db.RAW_FIELD_NAMES:
                error_label.config(text=f"字段 \"{fname}\" 已存在")
                return

            success = db.add_raw_field(fname, flabel, ftype)
            if not success:
                error_label.config(text=f"创建失败，字段 \"{fname}\" 可能已存在")
                return

            # Refresh all combobox choices and set this one to the new field
            popup.destroy()
            self._refresh_all_combos()
            new_choice = f"{flabel} ({fname})"
            w["var"].set(new_choice)
            w["status"].config(text="新字段", foreground="#00695c")

        btn_bar = ttk.Frame(popup)
        btn_bar.pack(fill="x", padx=15, pady=15)
        ttk.Button(btn_bar, text="创建", command=_do_create).pack(side="right", padx=3)
        ttk.Button(btn_bar, text="取消", command=popup.destroy).pack(side="right", padx=3)

    def _manage_fields(self, excel_col):
        """Open a dialog to manage (rename/delete) custom fields."""
        w = self._mapping_widgets.get(excel_col)
        extra = db.get_raw_case_extra_columns()
        # Filter out internal columns
        internal = {"id", "created_at", "updated_at"}
        custom = [c for c in extra if c not in internal]
        if not custom:
            messagebox.showinfo("管理字段", "当前没有自定义字段。")
            # Reset combobox back to skip
            if w:
                w["var"].set("-- 跳过此列 --")
                w["status"].config(text="跳过", foreground="#999")
            return

        popup = tk.Toplevel(self.win)
        popup.title("管理自定义字段")
        popup.geometry("520x350")
        popup.transient(self.win)
        popup.grab_set()

        ttk.Label(popup, text="自定义字段列表（可重命名或删除）",
                  font=("", 10, "bold")).pack(padx=15, pady=(15, 5))
        note = ("注意：删除字段将永久移除该列及其所有数据，不可恢复。\n"
                "重命名仅修改字段名，数据保留。")
        ttk.Label(popup, text=note, foreground="#666", font=("", 9)).pack(padx=15)

        # Scrollable list
        canvas = tk.Canvas(popup, highlightthickness=0, height=180)
        scrollbar = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", width=480)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=15, pady=5)
        scrollbar.pack(side="right", fill="y", pady=5)

        name_vars = {}  # fname -> StringVar

        for fname in custom:
            row = ttk.Frame(inner)
            row.pack(fill="x", pady=2)
            var = tk.StringVar(value=fname)
            name_vars[fname] = var
            ttk.Entry(row, textvariable=var, width=25).pack(side="left", padx=2)
            ttk.Button(row, text="保存", width=5,
                       command=lambda fn=fname, v=var: self._rename_field(fn, v.get(), popup)).pack(side="left", padx=2)
            ttk.Button(row, text="删除", width=5,
                       command=lambda fn=fname: self._delete_field(fn, popup)).pack(side="left", padx=2)
            ttk.Label(row, text=fname, foreground="#999", font=("", 8)).pack(side="left", padx=5)

        btn_bar = ttk.Frame(popup)
        btn_bar.pack(fill="x", padx=15, pady=10)
        ttk.Button(btn_bar, text="关闭", command=popup.destroy).pack(side="right", padx=3)

        # Reset the combobox selection back to skip
        if w:
            w["var"].set("-- 跳过此列 --")
            w["status"].config(text="跳过", foreground="#999")

    def _rename_field(self, old_name, new_name_var, popup):
        new_name = new_name_var.get().strip()
        if not new_name or new_name == old_name:
            return
        if not re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*$", new_name):
            messagebox.showerror("错误", "字段名只能包含英文字母、数字和下划线", parent=popup)
            return
        if new_name in db.RAW_FIELD_NAMES or new_name in db.get_raw_case_extra_columns():
            messagebox.showerror("错误", f"字段 \"{new_name}\" 已存在", parent=popup)
            return
        ok = db.rename_raw_field(old_name, new_name)
        if ok:
            self._refresh_all_combos()
            messagebox.showinfo("成功", f"字段已重命名：{old_name} → {new_name}", parent=popup)
        else:
            messagebox.showerror("失败", "重命名失败", parent=popup)

    def _delete_field(self, fname, popup):
        if not messagebox.askyesno("确认删除",
                                    f"确定要删除字段 \"{fname}\" 吗？\n\n"
                                    "该字段的所有数据将永久丢失！",
                                    parent=popup):
            return
        ok = db.delete_raw_field(fname)
        if ok:
            self._refresh_all_combos()
            messagebox.showinfo("成功", f"字段 \"{fname}\" 已删除", parent=popup)
        else:
            messagebox.showerror("失败", f"删除字段 \"{fname}\" 失败", parent=popup)

    def _do_import(self):
        if self.df is None:
            return

        mapping, unmapped, image_cols = self._parse_mapping()

        # Warn about unmapped columns
        if unmapped:
            unmapped_list = "\n".join(f"  - {c}" for c in unmapped)
            msg = ("以下列未被映射，其数据将不会导入：\n\n" + unmapped_list + "\n\n"
                   "是否继续导入？\n（可返回将未映射的列改为[合并到备注]以保留数据）")
            if not messagebox.askyesno("未映射的列", msg):
                return

        # Warn about duplicate field mappings
        field_counts = {}
        for v in mapping.values():
            if v and v != "__NOTES__":
                field_counts[v] = field_counts.get(v, 0) + 1
        duplicates = {k: c for k, c in field_counts.items() if c > 1}
        if duplicates:
            dup_list = "\n".join(f"  - {fn} ({count}个列映射)" for fn, count in duplicates.items())
            msg = ("以下字段被多个Excel列映射，后映射的列将覆盖先映射的列：\n\n"
                   + dup_list + "\n\n是否继续？")
            if not messagebox.askyesno("重复映射", msg):
                return

        imported, updated, skipped, errors = db.import_raw_from_dataframe(
            self.df, mapping,
            image_data=self.image_data,
            merge_extra_to_notes=False,
            update_existing=self._update_existing.get()
        )

        # Summary
        parts = [f"新增: {imported} 条"]
        if updated:
            parts.append(f"更新: {updated} 条")
        parts.append(f"跳过(空名称): {skipped} 条")
        msg = "导入完成！\n" + "\n".join(parts)
        if errors:
            msg += f"\n错误: {len(errors)} 条"
            if len(errors) <= 5:
                msg += "\n\n" + "\n".join(errors)
            else:
                msg += "\n\n" + "\n".join(errors[:5]) + f"\n... 及其他 {len(errors)-5} 条错误"

        messagebox.showinfo("导入结果", msg)

        self.win.destroy()

        if self.on_imported:
            self.on_imported(imported)

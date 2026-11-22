import sys
import json
import os
import re
import subprocess
import threading
import tkinter as tk
from collections import Counter
from tkinter import ttk, messagebox, filedialog
import urllib.request
import urllib.parse

import database as db
import config
from utils.import_export import import_excel, export_to_excel
from ui.import_dialog import ImportDialog
from utils.ai_extractor import extract_case_names, verify_case_by_search, cross_validate_cases, research_and_complete, verify_completed_case, verify_field_from_paste, verify_field_name_from_paste, verify_field_deepseek, verify_field_gemini, auto_complete_fields, auto_complete_with_gemini, search_images_with_gemini, _build_name_block, _call_yunwu, _extract_balanced_json, AUTO_COMPLETE_PROMPT, FIELD_VERIFY_DEEPSEEK_PROMPT, FIELD_VERIFY_DEEPSEEK_NAME_PROMPT
from ui.image_components import ImageThumbnailStrip, load_tree_thumbnail, merged_image_sources, parse_image_urls

from ui.search_tab import _ExtractServerHandler


def _tree_id(prefix, db_id):
    """Format an ID for tree display with prefix (Y/H/Z)."""
    return f"{prefix}-{db_id}"


def _parse_tree_id(value):
    """Strip Y-/H-/Z- prefix from a tree ID value, return int."""
    s = str(value)
    if len(s) > 2 and s[1] == "-" and s[0] in "YHZV":
        return int(s[2:])
    return int(s)


def _make_tree_draggable(tree, table):
    """Enable drag-and-drop reordering on a Treeview.

    Args:
        tree: ttk.Treeview widget
        table: database table name ("raw_cases" or "cases")
    """
    drag_data = {"item": None, "start_y": 0, "dragging": False}

    def _on_press(event):
        region = tree.identify_region(event.x, event.y)
        if region in ("heading", "separator"):
            return
        row = tree.identify_row(event.y)
        if row:
            drag_data["item"] = row
            drag_data["start_y"] = event.y
            drag_data["dragging"] = False

    def _on_motion(event):
        item = drag_data.get("item")
        if not item:
            return
        if not drag_data["dragging"]:
            if abs(event.y - drag_data["start_y"]) < 5:
                return
            drag_data["dragging"] = True
            tree.configure(cursor="fleur")

        target = tree.identify_row(event.y)
        if target and target != item:
            children = tree.get_children("")
            try:
                target_idx = children.index(target)
                tree.move(item, "", target_idx)
            except ValueError:
                pass

    def _on_release(event):
        if not drag_data.get("dragging"):
            drag_data["item"] = None
            return
        drag_data["dragging"] = False
        drag_data["item"] = None
        tree.configure(cursor="")

        ordered_ids = []
        for child in tree.get_children(""):
            vals = tree.item(child, "values")
            if vals:
                ordered_ids.append(_parse_tree_id(vals[0]))
        if ordered_ids:
            db.update_list_order(table, ordered_ids)

    tree.bind("<ButtonPress-1>", _on_press)
    tree.bind("<B1-Motion>", _on_motion)
    tree.bind("<ButtonRelease-1>", _on_release)


class OrganizeTab:
    """Top-level container with 3 sub-tabs for the data organization workflow."""

    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)

        self.notebook = ttk.Notebook(self.frame)
        self.notebook.pack(fill="both", expand=True)

        self.raw_panel = RawDataPanel(self.notebook, self)
        self.db_panel = DatabasePanel(self.notebook, self)
        self.verify_panel = VerificationPanel(self.notebook, self, self.raw_panel)

        self.notebook.add(self.raw_panel.frame, text="原始数据")
        self.notebook.add(self.verify_panel.frame, text="真实性核验")
        self.notebook.add(self.db_panel.frame, text="正式案例")


# ============================================================
#  Tab 1: Raw Data Panel
# ============================================================

class RawDataPanel:
    def __init__(self, parent, organizer):
        self.organizer = organizer
        self.frame = ttk.Frame(parent)
        self._sort_reverse = False
        self._sort_col = None

        self._build_toolbar()
        self._build_filter_bar()
        self._build_tree()
        self.refresh()

    def _build_toolbar(self):
        bar = ttk.Frame(self.frame)
        bar.pack(fill="x", padx=5, pady=(5, 0))

        ttk.Button(bar, text="案例详情", command=self._view_detail).pack(side="left", padx=3)
        ttk.Button(bar, text="提交核验（可多选）", command=self._submit_for_verification).pack(side="left", padx=3)
        ttk.Button(bar, text="导入 Excel/CSV", command=self._import_data).pack(side="left", padx=3)
        ttk.Button(bar, text="添加", command=self._add_raw_case).pack(side="left", padx=3)
        ttk.Button(bar, text="删除选中", command=self._delete_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="导出全部", command=self._export_all).pack(side="left", padx=3)

        ttk.Label(bar, text="状态筛选：").pack(side="left", padx=(15, 0))
        self.status_filter = ttk.Combobox(bar, values=["全部", "raw", "pending", "verified", "rejected"],
                                          state="readonly", width=10)
        self.status_filter.set("全部")
        self.status_filter.pack(side="left", padx=3)
        self.status_filter.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        self.status_label = ttk.Label(bar, text="")
        self.status_label.pack(side="right", padx=10)

        bar2 = ttk.Frame(self.frame)
        bar2.pack(fill="x", padx=5, pady=(0, 5))

        ttk.Button(bar2, text="刷新列表", command=self.refresh).pack(side="left", padx=3)
        ttk.Button(bar2, text="合并选中", command=self._merge_selected).pack(side="left", padx=3)
        ttk.Button(bar2, text="去重", command=self._show_duplicates).pack(side="left", padx=3)
        ttk.Button(bar2, text="智能合并重名", command=self._smart_merge).pack(side="left", padx=3)
        ttk.Button(bar2, text="清洗", command=self._clean_data).pack(side="left", padx=3)
        ttk.Button(bar2, text="查找检索词", command=self._lookup_keywords).pack(side="left", padx=3)
        self._undo_btn = ttk.Button(bar2, text="撤回", command=self._undo_last)
        self._undo_btn.pack(side="left", padx=3)
        self._update_undo_btn()

        self._redo_btn = ttk.Button(bar2, text="恢复", command=self._redo_last)
        self._redo_btn.pack(side="left", padx=3)
        self._update_redo_btn()

    def _build_filter_bar(self):
        bar = ttk.LabelFrame(self.frame, text=" 筛选 ", padding=5)
        bar.pack(fill="x", padx=5, pady=(2, 0))

        search_row = ttk.Frame(bar)
        search_row.pack(fill="x", pady=(0, 4))
        ttk.Label(search_row, text="搜索：").pack(side="left")
        self._search_var = tk.StringVar()
        e = ttk.Entry(search_row, textvariable=self._search_var, width=30)
        e.pack(side="left", padx=5)
        e.bind("<Return>", lambda e: self.refresh())
        ttk.Button(search_row, text="搜索", command=self.refresh).pack(side="left", padx=2)
        ttk.Button(search_row, text="清除筛选", command=self._clear_filters).pack(side="left", padx=5)

        filter_row = ttk.Frame(bar)
        filter_row.pack(fill="x")

        filter_fields = [
            ("country", "国家"),
            ("function", "功能"),
            ("spread_method", "传播方式"),
            ("building_status", "建筑状态"),
        ]

        self._filter_vars = {}
        self._filter_combos = {}

        for i, (field, label) in enumerate(filter_fields):
            pad_left = (0, 0) if i == 0 else (8, 0)
            ttk.Label(filter_row, text=f"{label}：").pack(side="left", padx=pad_left)
            var = tk.StringVar(value="全部")
            combo = ttk.Combobox(filter_row, textvariable=var, values=["全部"],
                                 state="readonly", width=12)
            combo.pack(side="left", padx=3)
            combo.bind("<<ComboboxSelected>>", lambda e, f=field: self.refresh())
            self._filter_vars[field] = var
            self._filter_combos[field] = combo

        self._load_filter_values()

    def _load_filter_values(self):
        for field, combo in self._filter_combos.items():
            current = self._filter_vars[field].get()
            col = "function" if field == "function" else field
            values = db.get_distinct_values("raw_cases", col)
            combo["values"] = ["全部"] + values
            if current == "全部" or current in values:
                self._filter_vars[field].set(current)
            else:
                self._filter_vars[field].set("全部")

    def _clear_filters(self):
        self._search_var.set("")
        for var in self._filter_vars.values():
            var.set("全部")
        self.status_filter.set("全部")
        self.refresh()

    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col
            self._sort_reverse = False
        rows = [(self.tree.set(item, col), item) for item in self.tree.get_children("")]
        try:
            rows.sort(key=lambda x: float(x[0]) if x[0].replace(".", "").isdigit() else x[0],
                      reverse=self._sort_reverse)
        except (ValueError, AttributeError, TypeError):
            rows.sort(key=lambda x: x[0], reverse=self._sort_reverse)
        for i, (_, item) in enumerate(rows):
            self.tree.move(item, "", i)

    def _apply_sort(self):
        col = self._sort_col
        reverse = self._sort_reverse
        rows = [(self.tree.set(item, col), item) for item in self.tree.get_children("")]
        try:
            rows.sort(key=lambda x: float(x[0]) if x[0].replace(".", "").isdigit() else x[0],
                      reverse=reverse)
        except (ValueError, AttributeError, TypeError):
            rows.sort(key=lambda x: x[0], reverse=reverse)
        for i, (_, item) in enumerate(rows):
            self.tree.move(item, "", i)

    def _build_tree(self):
        tree_frame = ttk.Frame(self.frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=2)

        columns = ("id", "name", "country", "city", "function", "construction_year")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")

        headers = {"id": "ID", "name": "建筑名称", "country": "国家", "city": "城市",
                   "function": "功能", "construction_year": "建造时间"}
        widths = {"id": 45, "name": 200, "country": 90, "city": 100,
                  "function": 80, "construction_year": 80}

        for col in columns:
            self.tree.heading(col, text=headers[col],
                              command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=widths.get(col, 80))

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda e: self._view_detail())
        # Cmd/Ctrl+A to select all
        self.tree.bind("<Command-a>" if sys.platform == "darwin" else "<Control-a>",
                       lambda e: self.tree.selection_set(self.tree.get_children()))

        _make_tree_draggable(self.tree, "raw_cases")

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        filter_val = self.status_filter.get()
        status = None if filter_val == "全部" else filter_val

        filters = {}
        if hasattr(self, '_filter_vars'):
            for field, var in self._filter_vars.items():
                val = var.get()
                if val and val != "全部":
                    filters[field] = val

        search = None
        if hasattr(self, '_search_var'):
            search = self._search_var.get().strip() or None

        cases = db.get_all_raw_cases(status=status, search=search,
                                     filters=filters if filters else None)

        for c in cases:
            self.tree.insert("", "end", values=(
                _tree_id("Y", c["id"]), c.get("name", ""), c.get("country", "") or "",
                c.get("city", "") or "", c.get("function", "") or "",
                c.get("construction_year", "") or "",
            ))
        self.status_label.config(text=f"共 {len(cases)} 条原始记录")
        self._update_undo_btn()
        self._update_redo_btn()

        if self._sort_col is not None:
            self._apply_sort()

        if hasattr(self, '_filter_combos'):
            self._load_filter_values()

    def _submit_for_verification(self):
        items = self.tree.selection()
        if not items:
            messagebox.showinfo("提示", "请先选择要提交核验的记录。")
            return

        already_in_queue = []
        for item in items:
            case_id = _parse_tree_id(self.tree.item(item, "values")[0])
            vc = db.get_verified_case_by_raw_id(case_id)
            if vc:
                raw = db.get_raw_case_by_id(case_id)
                name = raw.get("name", "?") if raw else str(case_id)
                already_in_queue.append((case_id, vc["id"], name))

        if already_in_queue:
            name_preview = "、".join(n[:15] for _, _, n in already_in_queue[:5])
            if len(already_in_queue) > 5:
                name_preview += f" 等共{len(already_in_queue)}条"
            if not messagebox.askyesno(
                "重复提交确认",
                f"以下 {len(already_in_queue)} 条案例已在核验队列中存在：\n\n"
                f"{name_preview}\n\n"
                f"是否仍然提交？将恢复其核验状态并刷新数据。"
            ):
                return

        count = 0
        for item in items:
            case_id = _parse_tree_id(self.tree.item(item, "values")[0])
            current = db.get_raw_case_by_id(case_id)
            if not current:
                continue
            db.set_raw_case_status(case_id, "pending")
            vc_id = db.copy_raw_to_verified(case_id)
            if vc_id:
                db.set_verified_case_status(vc_id, "unverified")
            db.insert_verification_log(case_id, "submit", "提交核验")
            count += 1
        total = len(db.get_all_raw_cases())
        self.refresh()
        self.organizer.verify_panel._verify_search_var.set("")
        self.organizer.verify_panel.refresh()
        messagebox.showinfo("完成", f"已提交 {count} 条。当前列表共 {total} 条案例。")

    def _delete_selected(self):
        items = self.tree.selection()
        if not items:
            messagebox.showinfo("提示", "请先选择要删除的记录。")
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(items)} 条记录吗？"):
            for item in items:
                case_id = self.tree.item(item, "values")[0]
                db.delete_raw_case(_parse_tree_id(case_id))
            self.refresh()

    def _add_raw_case(self):
        _AddRawCaseDialog(self.frame, on_done=lambda: self.refresh())

    def _undo_last(self):
        result = db.undo_last()
        if result:
            self.status_label.config(text=f"已{result}")
        else:
            self.status_label.config(text="没有可撤回的操作")
        self.refresh()
        self._update_undo_btn()
        self._update_redo_btn()

    def _redo_last(self):
        result = db.redo_last()
        if result:
            self.status_label.config(text=f"已{result}")
        else:
            self.status_label.config(text="没有可恢复的操作")
        self.refresh()
        self._update_undo_btn()
        self._update_redo_btn()

    def _update_undo_btn(self):
        count = db.get_undo_count()
        if count > 0:
            self._undo_btn.config(text=f"撤回({count})")
        else:
            self._undo_btn.config(text="撤回")

    def _update_redo_btn(self):
        count = db.get_redo_count()
        if count > 0:
            self._redo_btn.config(text=f"恢复({count})")
        else:
            self._redo_btn.config(text="恢复")

    def _import_data(self):
        ImportDialog(self.frame, on_imported=lambda count: self.refresh())

    def _export_all(self):
        from tkinter import filedialog
        import csv
        path = filedialog.asksaveasfilename(
            title="导出原始数据",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="raw_cases_export.csv"
        )
        if not path:
            return
        cases = db.get_all_raw_cases()
        if not cases:
            messagebox.showinfo("提示", "没有数据可导出。")
            return
        # Get all columns from the first case
        columns = list(cases[0].keys())
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            writer.writerows(cases)
        messagebox.showinfo("导出完成", f"已导出 {len(cases)} 条记录到:\n{path}")

    def _merge_selected(self):
        items = self.tree.selection()
        if len(items) < 2:
            messagebox.showinfo("提示", "请至少选择两条记录进行合并。\n\n第一条为主记录，其余记录的数据将合并到主记录中。")
            return

        cases = []
        for item in items:
            case_id = _parse_tree_id(self.tree.item(item, "values")[0])
            case = db.get_raw_case_by_id(case_id)
            if not case:
                messagebox.showerror("错误", f"无法读取选中记录 (ID={case_id})。")
                return
            cases.append(case)

        _MergeDialog(self.frame, cases, on_done=lambda: self.refresh())

    def _view_detail(self):
        sel = self.tree.selection()
        if not sel:
            return
        case_id = _parse_tree_id(self.tree.item(sel[0], "values")[0])
        case = db.get_raw_case_by_id(case_id)
        if case:
            _RawCaseDetailDialog(self.frame, case)

    def _smart_merge(self):
        groups = db.find_duplicates_by_name_country()
        if not groups:
            messagebox.showinfo("智能合并", "未发现名称+国家相同的重复记录。")
            return

        self._merge_groups = groups
        self._merge_idx = 0
        self._merge_total = len(groups)
        self._merge_skipped = 0

        def _process_next():
            # Refresh group data in case previous merges changed records
            while self._merge_idx < self._merge_total:
                g = self._merge_groups[self._merge_idx]
                cases = []
                for rid in g["ids"]:
                    case = db.get_raw_case_by_id(rid)
                    if case:
                        cases.append(case)
                self._merge_idx += 1
                if len(cases) >= 2:
                    dlg = _MergeDialog(self.frame, cases, on_done=_on_group_done,
                                       smart_defaults=True)
                    dlg.win.title(f"合并 ({self._merge_idx}/{self._merge_total}) - {len(cases)}条记录")
                    return
                else:
                    self._merge_skipped += 1

            # All groups processed
            msg = f"全部 {self._merge_total} 组合并完成。"
            if self._merge_skipped:
                msg += f"\n跳过 {self._merge_skipped} 组（记录不完整）。"
            messagebox.showinfo("智能合并", msg)
            self.refresh()

        def _on_group_done():
            _process_next()

        # Ask user if they want to proceed
        if not messagebox.askyesno("智能合并",
                                   f"发现 {self._merge_total} 组名称+国家相同的重复记录。\n\n"
                                   f"将从上到下依次打开每组合并对话框，\n"
                                   f"每组自动预选智能默认值，您确认后自动进入下一组。\n\n是否开始？"):
            return

        _process_next()

    def _show_duplicates(self):
        dupes = db.find_raw_duplicates_all_fields()
        if not dupes:
            messagebox.showinfo("去重检查", "未发现重复记录。")
            return

        win = tk.Toplevel(self.frame)
        win.title(f"发现 {len(dupes)} 组重复记录")
        win.geometry("700x400")
        win.transient(self.frame)

        ttk.Label(win, text="以下记录所有字段完全一致，判断为重复，请手动确认并删除：",
                  font=("", 10)).pack(padx=10, pady=5)

        tree = ttk.Treeview(win, columns=("name", "country", "count"), show="headings")
        tree.heading("name", text="建筑名称")
        tree.heading("country", text="国家")
        tree.heading("count", text="重复数量")
        tree.column("name", width=250)
        tree.column("country", width=150)
        tree.column("count", width=80)
        tree.pack(fill="both", expand=True, padx=10, pady=5)

        for i, d in enumerate(dupes):
            tree.insert("", "end", values=(d["name"], d.get("country", ""), d["count"]), tags=(str(i),))

        def delete_dupes():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("提示", "请先在列表中选中要处理的重复记录。")
                return
            for item in sel:
                idx = int(tree.item(item, "tags")[0])
                ids = dupes[idx]["ids"]
                for dup_id in ids[1:]:
                    db.delete_raw_case(dup_id)
            messagebox.showinfo("完成", "已处理选中的重复记录（保留每组第一条）。")
            win.destroy()
            self.refresh()

        ttk.Button(win, text="删除选中重复（保留第一条）", command=delete_dupes).pack(pady=8)

    def _clean_data(self):
        if not messagebox.askyesno("数据清洗", "将执行以下操作：\n- 去除所有文本字段首尾空白\n\n确定继续？"):
            return

        cases = db.get_all_raw_cases()
        cleaned = 0
        text_fields = [f[0] for f in db.RAW_FIELDS if "TEXT" in f[2]]
        for c in cases:
            updates = {}
            for field in text_fields:
                val = c.get(field)
                if isinstance(val, str) and val.strip() != val:
                    updates[field] = val.strip()
            if updates:
                db.update_raw_case(c["id"], updates)
                cleaned += 1

        messagebox.showinfo("清洗完成", f"已清洗 {cleaned} 条记录。")
        self.refresh()

    def _lookup_keywords(self):
        """Batch-update search_keywords for all raw cases by looking up
        their source URLs in raw_search_results."""
        if not messagebox.askyesno("查找检索词",
                                   "将为所有有来源链接的原始案例查找检索词。\n"
                                   "在搜索主列表中按URL匹配，找到则用搜索词覆盖。\n\n是否继续？"):
            return
        updated = db.batch_update_search_keywords()
        messagebox.showinfo("查找完成", f"已更新 {updated} 条记录的检索词。")
        self.refresh()


# ============================================================
#  Tab 2: Verification Panel
def _classify_verification_result(action, result):
    """Return (status_text, is_success, is_failure) for a verification log entry."""
    r = (result or "").strip()
    if not r:
        return ("", False, False)
    if action == "name":
        if "exception" in r or "parse_error" in r:
            return ("❌ 失败", False, True)
        if "correct:" in r:
            return ("✅ 成功", True, False)
        if r.startswith("{"):
            try:
                obj = json.loads(r)
                if obj.get("correct_value"):
                    return ("✅ 成功", True, False)
                return ("⚠️ 待确认", False, False)
            except Exception:
                pass
        return ("✅ 成功", True, False)
    if action == "fields":
        if "exception" in r or "parse_error" in r:
            return ("❌ 失败", False, True)
        if "filled:" in r:
            return ("✅ 成功", True, False)
        if "empty:" in r:
            return ("❌ 失败", False, True)
        return ("✅ 成功", True, False)
    return (r[:50], False, False)


def _load_batch_case_status(cid):
    """Return (name_status, name_reason, fields_status, fields_reason) for a batch case."""
    import database as db
    vcase = db.get_verified_case_by_id(cid)
    raw_log_id = vcase.get("raw_case_id", cid) if vcase else cid
    logs = db.get_verification_logs(raw_log_id)
    name_status = "⏳ 等待"
    name_reason = ""
    fields_status = "⏳ 等待"
    fields_reason = ""
    for log in logs:
        step = log.get("action", "")
        result = log.get("result", "")
        if step == "name" and name_status == "⏳ 等待":
            status, _, _ = _classify_verification_result("name", result)
            if status:
                name_status = status
                name_reason = result[:200] if result else ""
        elif step == "fields" and fields_status == "⏳ 等待":
            status, _, _ = _classify_verification_result("fields", result)
            if status:
                fields_status = status
                fields_reason = result[:200] if result else ""
    return name_status, name_reason, fields_status, fields_reason


# ============================================================

class BatchLogDialog(tk.Toplevel):
    def __init__(self, parent, panel):
        super().__init__(parent)
        self.panel = panel
        self.title("批量核验日志")
        self.geometry("700x450")
        self.transient(parent)

        ctrl = ttk.Frame(self)
        ctrl.pack(fill="x", padx=8, pady=(8, 4))
        self._summary_var = tk.StringVar(value="")
        ttk.Label(ctrl, textvariable=self._summary_var).pack(side="left", padx=16)
        ttk.Button(ctrl, text="刷新", command=self._refresh).pack(side="left", padx=2)
        ttk.Button(ctrl, text="补跑失败", command=self._retry_failed).pack(side="left", padx=2)

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=4)

        columns = ("case_name", "step", "status", "detail", "time")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("case_name", text="案例名")
        self.tree.column("case_name", width=180)
        self.tree.heading("step", text="步骤")
        self.tree.column("step", width=60)
        self.tree.heading("status", text="状态")
        self.tree.column("status", width=80)
        self.tree.heading("detail", text="详情")
        self.tree.column("detail", width=200)
        self.tree.heading("time", text="时间")
        self.tree.column("time", width=140)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self._refresh()

    def _refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        import database as db
        self._log_rows = []
        success_count = 0
        fail_count = 0
        for cid in self.panel._batch_rows:
            case_name = self.panel._batch_tree.item(self.panel._batch_rows[cid], "text")
            vcase = db.get_verified_case_by_id(cid)
            raw_id = vcase.get("raw_case_id", cid) if vcase else cid
            logs = db.get_verification_logs(raw_id)
            for log in logs:
                step = log.get("action", "?")
                if step not in ("name", "fields"):
                    continue
                result = log.get("result", "") or ""
                created = log.get("created_at", "") or ""
                status, is_success, is_fail = _classify_verification_result(step, result)
                if is_success:
                    success_count += 1
                elif is_fail:
                    fail_count += 1
                detail = result[:200] if result else ""
                item = self.tree.insert("", "end", values=(case_name, step, status, detail, created))
                self._log_rows.append((item, cid, step))
        self._summary_var.set(f"成功 {success_count}  ·  失败 {fail_count}")

    def _retry_failed(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中失败的日志行。")
            return
        to_retry = []
        for item in sel:
            vals = self.tree.item(item, "values")
            if vals and "失败" in str(vals[2]):
                for log_item, cid, step in self._log_rows:
                    if log_item == item:
                        batch_item = self.panel._batch_rows.get(cid)
                        if batch_item:
                            to_retry.append((cid, step, batch_item))
                        break
        if not to_retry:
            messagebox.showinfo("提示", "请选中失败的日志行。")
            return
        import threading
        t = threading.Thread(target=self.panel._run_batch_retry, args=(to_retry,), daemon=True)
        t.start()
        self.after(2000, self._refresh)

class VerificationPanel:
    def __init__(self, parent, organizer, raw_panel):
        self.organizer = organizer
        self.raw_panel = raw_panel
        self.frame = ttk.Frame(parent)
        self._api_key = tk.StringVar(value=config.get_llm("search_parse")[2])
        self._sort_reverse = False
        self._sort_col = None
        self._batch_name_result = None

        # Wrap in horizontally-scrollable canvas
        canvas = tk.Canvas(self.frame, highlightthickness=0)
        h_scroll = ttk.Scrollbar(self.frame, orient="horizontal", command=canvas.xview)
        canvas.configure(xscrollcommand=h_scroll.set)

        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_canvas_configure(event):
            canvas.itemconfig(inner_id, width=max(1600, event.width))
            canvas.itemconfig(inner_id, height=event.height)
        canvas.bind("<Configure>", _on_canvas_configure)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        paned = ttk.PanedWindow(inner, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        # Left: pending queue
        left = ttk.Frame(paned)
        self._build_queue(left)
        paned.add(left, weight=1)

        # Middle: case detail
        middle = ttk.Frame(paned)
        self._build_detail(middle)
        paned.add(middle, weight=2)

        # Right: batch verify
        right = ttk.Frame(paned)
        self._build_batch_verify(right)
        paned.add(right, weight=1)

        canvas.pack(side="top", fill="both", expand=True)
        h_scroll.pack(side="bottom", fill="x")

    def _build_queue(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 3))
        self.lang_var = tk.StringVar(value="中文")
        ttk.Label(bar, text="状态：").pack(side="left")
        self._verify_status_var = tk.StringVar(value="全部")
        status_combo = ttk.Combobox(bar, textvariable=self._verify_status_var,
                                  values=["全部", "pending", "name_verified", "verified", "rejected"],
                                  state="readonly", width=10)
        status_combo.pack(side="left", padx=2)
        status_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        ttk.Button(bar, text="核验详情", command=self._open_verify_detail).pack(side="left", padx=2)
        ttk.Button(bar, text="修改", command=self._enable_field_edit).pack(side="left", padx=2)
        ttk.Button(bar, text="保存", command=self._save_field_changes).pack(side="left", padx=2)
        ttk.Button(bar, text="删除", command=self._delete_selected).pack(side="left", padx=2)
        self.queue_label = ttk.Label(bar, text="")
        self.queue_label.pack(side="right", padx=3)

        bar2 = ttk.Frame(parent)
        bar2.pack(fill="x", pady=(0, 3))
        ttk.Button(bar2, text="刷新队列", command=self.refresh).pack(side="left", padx=2)
        ttk.Button(bar2, text="去重", command=self._show_duplicates).pack(side="left", padx=2)
        ttk.Button(bar2, text="智能去重", command=self._smart_dedup).pack(side="left", padx=2)
        ttk.Button(bar2, text="清洗", command=self._clean_data).pack(side="left", padx=2)
        ttk.Button(bar2, text="通过核验", command=self._batch_approve).pack(side="left", padx=(10, 2))

        # Filter bar
        filter_bar = ttk.Frame(parent)
        filter_bar.pack(fill="x", pady=(0, 3))
        ttk.Label(filter_bar, text="搜索：").pack(side="left")
        self._verify_search_var = tk.StringVar()
        e = ttk.Entry(filter_bar, textvariable=self._verify_search_var, width=14)
        e.pack(side="left", padx=3)
        e.bind("<Return>", lambda e: self.refresh())
        ttk.Button(filter_bar, text="搜索", command=self.refresh).pack(side="left", padx=2)
        ttk.Button(filter_bar, text="发送到批量核验", command=self._send_to_batch).pack(side="left", padx=(10, 2))

        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)

        columns = ("id", "country", "name")
        self.queue_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        headers = {"id": "ID", "country": "国家", "name": "建筑名称"}
        widths = {"id": 30, "country": 60, "name": 200}
        for col in columns:
            self.queue_tree.heading(col, text=headers[col],
                                    command=lambda c=col: self._sort_queue_by(c))
            self.queue_tree.column(col, width=widths.get(col, 60))

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.queue_tree.yview)
        self.queue_tree.configure(yscrollcommand=scroll_y.set)
        self.queue_tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self.queue_tree.bind("<<TreeviewSelect>>", self._on_select)
        self.queue_tree.bind("<Double-1>", self._on_queue_double_click)

        _make_tree_draggable(self.queue_tree, "raw_cases")

    def _build_detail(self, parent):
        # Scrollable container
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self._detail_inner = ttk.Frame(canvas)

        self._detail_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._detail_inner, anchor="nw", width=canvas.winfo_reqwidth())

        def _on_canvas_configure(event):
            canvas.itemconfig(1, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Bind mousewheel for scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 60)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        inner = self._detail_inner

        # Field-by-field verification buttons
        field_frame = ttk.LabelFrame(inner, text=" 逐字段核验 ", padding=5)
        field_frame.pack(fill="x", pady=(0, 5))

        # Header row
        header = ttk.Frame(field_frame)
        header.pack(fill="x", pady=(0, 2))
        ttk.Label(header, text="字段", width=8, font=("", 9, "bold")).pack(side="left", padx=1)
        ttk.Label(header, text="当前值", width=16, font=("", 9, "bold")).pack(side="left", padx=1)
        ttk.Label(header, text="操作", width=16, font=("", 9, "bold")).pack(side="left", padx=1)
        ttk.Label(header, text="核验记录", width=14, font=("", 9, "bold")).pack(side="left", padx=1)

        self._field_rows = {}  # field_name -> {"var": StringVar, "entry": entry, "status_label": label, "row": row}
        self._field_statuses = {}  # field_name -> "unverified" | "confirmed" | "conflict" | "found"
        self._active_field = None  # currently selected field for shared engines

        verify_fields = [
            ("country", "国家"),
            ("name", "建筑名称"),
            ("name_en", "英文名称"),
            ("name_official", "官方语言名称"),
            ("city", "城市"),
            ("construction_year", "建造时间"),
            ("demolition_year", "拆除时间"),
            ("restoration_year", "修复时间"),
            ("building_status", "建筑状态"),
            ("spread_method", "传播方式"),
            ("quantity", "数量"),
            ("function", "功能"),
            ("designer", "设计师"),
            ("coordinates", "经纬度"),
            ("image_urls", "图片"),
        ]

        for field_name, field_label in verify_fields:
            row = ttk.Frame(field_frame)
            row.pack(fill="x", pady=1)

            ttk.Label(row, text=field_label, width=8, anchor="e").pack(side="left", padx=1)
            var = tk.StringVar()
            entry = ttk.Entry(row, textvariable=var, width=16)
            entry.pack(side="left", padx=1)

            # Bind click on entry to select as active field for shared engines
            entry.bind("<Button-1>", lambda e, fn=field_name: self._set_active_field(fn))
            entry.bind("<FocusIn>", lambda e, fn=field_name: self._set_active_field(fn))

            btn_frame = ttk.Frame(row, width=16)
            btn_frame.pack(side="left", padx=1)

            if field_name == "name":
                ttk.Button(btn_frame, text="百度", width=4,
                           command=lambda fn=field_name, fl=field_label: self._open_field_search(fn, fl, "baidu")).pack(side="left", padx=1)
                ttk.Button(btn_frame, text="谷歌", width=4,
                           command=lambda fn=field_name, fl=field_label: self._open_field_search(fn, fl, "google")).pack(side="left", padx=1)
                ttk.Button(btn_frame, text="DS", width=4,
                           command=lambda fn=field_name, fl=field_label: self._verify_field_deepseek(fn, fl)).pack(side="left", padx=1)
                ttk.Button(btn_frame, text="G", width=3,
                           command=lambda fn=field_name, fl=field_label: self._verify_field_gemini(fn, fl)).pack(side="left", padx=1)
            elif field_name != "country":
                ttk.Button(btn_frame, text="选择字段", width=7,
                           command=lambda fn=field_name, fl=field_label: self._set_active_field(fn, fl)).pack(side="left", padx=1)

            status_label = ttk.Label(row, text="—", width=14, anchor="w", foreground="#999")
            status_label.pack(side="left", padx=2)

            self._field_rows[field_name] = {"var": var, "entry": entry, "status_label": status_label, "row": row}
            self._field_statuses[field_name] = "unverified"

        # Image preview area (below field rows)
        self._image_preview_frame = ttk.Frame(field_frame)
        self._image_preview_frame.pack(fill="x", pady=(3, 0))

        # Image refresh button
        img_btn_bar = ttk.Frame(field_frame)
        img_btn_bar.pack(fill="x", pady=(2, 0))
        ttk.Button(img_btn_bar, text="刷新图片", width=10,
                   command=lambda: self._refresh_image_preview(self._current_case) if hasattr(self, "_current_case") and self._current_case else None).pack(side="left", padx=1)
        self._undo_btn = ttk.Button(img_btn_bar, text="撤销删除", width=10,
                                    command=self._undo_delete_image, state="disabled")
        self._undo_btn.pack(side="left", padx=1)

        # Manual image download
        img_dl_bar = ttk.Frame(field_frame)
        img_dl_bar.pack(fill="x", pady=(2, 0))
        ttk.Label(img_dl_bar, text="图片URL：", font=("", 9)).pack(side="left")
        self._img_url_var = tk.StringVar()
        ttk.Entry(img_dl_bar, textvariable=self._img_url_var, width=50).pack(side="left", padx=2, fill="x", expand=True)

        img_dl_btn_bar = ttk.Frame(field_frame)
        img_dl_btn_bar.pack(fill="x", pady=(2, 0))
        self._download_btn = ttk.Button(img_dl_btn_bar, text="下载图片", width=10,
                                     command=self._download_image)
        self._download_btn.pack(side="left", padx=1)
        ttk.Button(img_dl_btn_bar, text="选择本地图片", width=12,
                   command=self._select_local_image).pack(side="left", padx=1)
        ttk.Button(img_dl_btn_bar, text="百度搜图", width=10,
                   command=self._baidu_image_search).pack(side="left", padx=1)
        ttk.Button(img_dl_btn_bar, text="谷歌搜图", width=10,
                   command=self._google_image_search).pack(side="left", padx=1)

        # Shared search engine bar at bottom — three rows
        engine_area = ttk.Frame(field_frame)
        engine_area.pack(fill="x", pady=(5, 0))

        # Row 1: prompt
        row1 = ttk.Frame(engine_area)
        row1.pack(fill="x")
        ttk.Label(row1, text="选中字段后选择搜索引擎：", font=("", 9)).pack(side="left")
        self._active_field_label = ttk.Label(row1, text="（未选中）", font=("", 9, "italic"), foreground="#999")
        self._active_field_label.pack(side="left", padx=8)

        # Row 2: four engine buttons
        row2 = ttk.Frame(engine_area)
        row2.pack(fill="x", pady=(3, 0))
        ttk.Button(row2, text="百度搜索", width=10,
                   command=lambda: self._search_active_field("baidu")).pack(side="left", padx=2)
        ttk.Button(row2, text="谷歌搜索", width=10,
                   command=lambda: self._search_active_field("google")).pack(side="left", padx=2)
        ttk.Button(row2, text="DS核验", width=10,
                   command=lambda: self._search_active_field("ds")).pack(side="left", padx=2)
        ttk.Button(row2, text="G核验", width=10,
                   command=lambda: self._verify_active_field_gemini()).pack(side="left", padx=2)

        # Row 3: one-click complete
        row3 = ttk.Frame(engine_area)
        row3.pack(fill="x", pady=(3, 0))
        ttk.Button(row3, text="一键DS补全", width=12,
                   command=self._auto_complete_all).pack(side="left", padx=2)
        ttk.Button(row3, text="一键G补全", width=12,
                   command=self._auto_complete_gemini).pack(side="left", padx=2)

        # Currently verifying field indicator
        self._verifying_field = None
        self._verify_field_label = ttk.Label(field_frame, text="", foreground="#1565c0", font=("", 9, "italic"))
        self._verify_field_label.pack(fill="x", pady=(3, 0))

        # Search results paste area
        paste_frame = ttk.LabelFrame(inner, text=" 搜索结果内容（Chrome插件自动回传，也可手动粘贴） ", padding=5)
        paste_frame.pack(fill="x", pady=(0, 2))

        self.paste_text = tk.Text(paste_frame, wrap="word", font=("", 9), height=4)
        paste_scroll = ttk.Scrollbar(paste_frame, orient="vertical", command=self.paste_text.yview)
        self.paste_text.configure(yscrollcommand=paste_scroll.set)
        self.paste_text.pack(side="left", fill="x", expand=True)
        paste_scroll.pack(side="right", fill="y")

        self._paste_sources = set()

        def _on_paste(event):
            self._paste_sources.clear()
            self._paste_sources.add("手动粘贴")
            self.paste_text.after(100, lambda: self._save_paste_content("手动粘贴"))

        self.paste_text.bind("<Control-v>", _on_paste)
        self.paste_text.bind("<Command-v>", _on_paste)

        # Analysis buttons below paste area
        paste_btn_bar = ttk.Frame(inner)
        paste_btn_bar.pack(fill="x", pady=(0, 5))
        ttk.Button(paste_btn_bar, text="名称字段信息分析", command=self._run_name_field_analysis).pack(side="left", padx=2)
        ttk.Button(paste_btn_bar, text="其他字段信息分析", command=self._parse_and_fill_from_text).pack(side="left", padx=2)
        ttk.Button(paste_btn_bar, text="清空内容", command=self._clear_paste_area).pack(side="left", padx=2)

        # Verification result — split into name and other fields
        # Name field result
        name_result_frame = ttk.LabelFrame(inner, text=" 名称字段核验结果 ", padding=5)
        name_result_frame.pack(fill="x", pady=(0, 3))
        self.name_result_text = tk.Text(name_result_frame, wrap="word", font=("", 10), height=4)
        self.name_result_text.pack(fill="x")

        # Other fields result
        other_result_frame = ttk.LabelFrame(inner, text=" 其他字段核验结果 ", padding=5)
        other_result_frame.pack(fill="x", pady=(0, 3))
        result_scroll_frame = tk.Frame(other_result_frame)
        result_scroll_frame.pack(fill="both", expand=True)
        self.result_text = tk.Text(result_scroll_frame, wrap="word", font=("", 10), height=12)
        result_scrollbar = tk.Scrollbar(result_scroll_frame, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)
        self.result_text.pack(side="left", fill="both", expand=True)
        result_scrollbar.pack(side="right", fill="y")

        # Shared result buttons
        result_btn_bar = ttk.Frame(inner)
        result_btn_bar.pack(fill="x", pady=(0, 3))
        ttk.Button(result_btn_bar, text="保存结果", command=self._save_manual_result).pack(side="left", padx=2)
        ttk.Button(result_btn_bar, text="清空结果", command=self._clear_result).pack(side="left", padx=2)

        # Overall assessment
        self.overall_var = tk.StringVar(value="选择左侧记录开始核验 — 点击字段旁的「选择字段」，再用下方引擎")
        ttk.Label(inner, textvariable=self.overall_var, wraplength=600,
                  font=("", 9, "italic"), foreground="#555").pack(fill="x", pady=(0, 5))

        # Action buttons
        btn_bar = ttk.Frame(inner)
        btn_bar.pack(fill="x", pady=(0, 5))

        ttk.Button(btn_bar, text="应用当前字段结果", command=self._apply_field_result).pack(side="left", padx=2)

    def _get_verify_stats(self, case_id):
        """Return (verified_count, total_fields, field_methods) for a case.
        field_methods: dict field_name -> [(method, count), ...] sorted by count desc.
        case_id can be either verified_case_id or raw_case_id.
        """
        verify_fields = ["name", "country", "city",
                        "construction_year", "demolition_year", "restoration_year",
                        "building_status", "spread_method", "quantity", "function", "designer",
                        "coordinates", "image_urls"]
        total = len(verify_fields)
        # Resolve raw_case_id for log lookup
        raw_log_id = case_id
        vcase = db.get_verified_case_by_id(case_id)
        if vcase:
            raw_log_id = vcase.get("raw_case_id", case_id)
        logs = db.get_verification_logs(raw_log_id)
        field_actions = {}  # field_name -> list of action types
        for log in logs:
            action = log.get("action", "")
            details = log.get("details", "")
            # Map action to display method
            method_map = {
                "field_verify_ai": "AI",
                "field_verify_deepseek": "DS",
                "field_verify_gemini": "G",
                "field_complete_gemini": "G",
            }
            method = method_map.get(action)
            if not method:
                continue
            # Extract field from details (e.g., "AI分析: 建筑名称" or "Gemini核验: country")
            for fn in verify_fields:
                field_labels_map = {
                    "name": "建筑名称", "country": "国家", "city": "城市",
                    "construction_year": "建造时间", "demolition_year": "拆除时间",
                    "restoration_year": "修复时间", "building_status": "建筑状态",
                    "spread_method": "传播方式",
                    "quantity": "数量", "function": "功能", "designer": "设计师",
                    "coordinates": "经纬度", "image_urls": "图片",
                }
                if field_labels_map.get(fn, "") in details or fn in details:
                    if fn not in field_actions:
                        field_actions[fn] = []
                    field_actions[fn].append(method)
                    break

        verified = len(field_actions)
        # Count methods per field
        field_methods = {}
        for fn, methods in field_actions.items():
            c = Counter(methods)
            field_methods[fn] = c.most_common()
        return verified, total, field_methods

    def _get_field_val(self, case, field_name):
        """Get display value for a field, handling virtual fields like coordinates."""
        if field_name == "coordinates":
            lat = case.get("latitude")
            lng = case.get("longitude")
            parts = [str(v) for v in [lat, lng] if v is not None and v != ""]
            return ", ".join(parts) if parts else ""
        return case.get(field_name, "") or ""

    def _extract_year(self, val_str):
        """Extract a single 4-digit year number from a string.

        Returns the year as a string (e.g. '1970'), or the original value if no
        year-like number is found.
        """
        if not val_str:
            return val_str
        match = re.search(r'(?<!\d)(1[0-9]{3}|20[0-2][0-9])(?!\d)', str(val_str))
        return match.group(1) if match else val_str

    def _normalize_building_status(self, val_str):
        """Normalize building_status to one of four standard values, or 未知.

        Only allows: 状态良好, 破损待修复, 损毁严重, 拆除, 未知
        """
        if not val_str:
            return val_str
        v = val_str.strip()
        # Direct match
        if v in ("状态良好", "破损待修复", "损毁严重", "损坏严重", "拆除", "未知"):
            if v == "损坏严重":
                return "损毁严重"
            return v
        # Only classify when keywords clearly indicate building condition
        if any(w in v for w in ("良好", "完好", "正常", "完好无损", "保存完好")):
            return "状态良好"
        if any(w in v for w in ("待修复", "破损", "需修缮", "部分损坏", "needs repair",
                                 "damaged", "年久失修")):
            return "破损待修复"
        if any(w in v for w in ("严重", "损毁", "严重损坏", "危房", "倒塌", "severe",
                                 "危", "严重破损", "collapsed", "ruin", "decay", "衰败")):
            return "损毁严重"
        if any(w in v for w in ("拆除", "demolished", "已拆", "torn down", "dismantled",
                                 "gone", "不再存在", "已不存在", "destroyed", "已毁")):
            return "拆除"
        # Cannot confidently classify → 未知
        return "未知"

    def _load_verify_filter_values(self):
        values = db.get_distinct_values("raw_cases", "country")
        self._verify_country_combo["values"] = ["全部"] + values

    def _clear_verify_filters(self):
        self._verify_search_var.set("")
        self._verify_country_var.set("全部")
        self.refresh()

    def _sort_queue_by(self, col):
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col
            self._sort_reverse = False
        self._apply_sort()

    def _apply_sort(self):
        col = self._sort_col
        if not col:
            return
        reverse = self._sort_reverse
        if col == "country":
            # Sort: entries with name+country first, country-only (no name) last
            items = [(self.queue_tree.set(item, "country"),
                      self.queue_tree.set(item, "name"),
                      item) for item in self.queue_tree.get_children("")]
            with_name = [(c, n, i) for c, n, i in items if c.strip() and n.strip()]
            without_name = [(c, n, i) for c, n, i in items if not (c.strip() and n.strip())]
            # Sort with_name by country count, then country name, then case name
            country_counts = {}
            for co, _, _ in items:
                co = co.strip()
                if co:
                    country_counts[co] = country_counts.get(co, 0) + 1
            def _key(x):
                return (
                    country_counts.get(x[0].strip(), 0) * (-1 if not reverse else 1),
                    x[0].strip(),
                    x[1].strip(),
                )
            with_name.sort(key=_key)
            sorted_items = with_name + without_name
            for i, (_, _, item) in enumerate(sorted_items):
                self.queue_tree.move(item, "", i)
        else:
            rows = [(self.queue_tree.set(item, col), item) for item in self.queue_tree.get_children("")]
            try:
                rows.sort(key=lambda x: float(x[0]) if x[0].replace(".", "").isdigit() else x[0],
                          reverse=reverse)
            except (ValueError, AttributeError, TypeError):
                rows.sort(key=lambda x: x[0], reverse=reverse)
            for i, (_, item) in enumerate(rows):
                self.queue_tree.move(item, "", i)

    def refresh(self):
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)

        search = None
        if hasattr(self, '_verify_search_var'):
            search = self._verify_search_var.get().strip() or None

        status_filter = self._verify_status_var.get() if hasattr(self, '_verify_status_var') else "全部"
        if status_filter and status_filter != "全部":
            cases = db.get_all_verified_cases(status=status_filter, search=search)
        else:
            cases = db.get_all_verified_cases(search=search)
        # Sort: countries with more cases first, then by country name, then by case name
        country_counts = {}
        for c in cases:
            co = (c.get("country") or "").strip()
            if co:
                country_counts[co] = country_counts.get(co, 0) + 1
        cases.sort(key=lambda c: (
            -country_counts.get((c.get("country") or "").strip(), 0),
            (c.get("country") or "").strip(),
            (c.get("name") or "").strip(),
        ))
        lang = self.lang_var.get() if hasattr(self, "lang_var") else "中文"
        for c in cases:
            self.queue_tree.insert("", "end", values=(
                _tree_id("V", c["id"]), c.get("country", "") or "",
                db.get_display_name(c, lang),
            ))
        total = len(cases)
        self.queue_label.config(text=f"共 {total} 条案例")
        self._clear_detail()

        if self._sort_col is not None:
            self._apply_sort()



    def _clear_detail(self):
        for field_name, row_info in self._field_rows.items():
            row_info["var"].set("")
            row_info["status_label"].config(text="—", foreground="#999")
            self._field_statuses[field_name] = "unverified"
        self.paste_text.configure(state="normal")
        self.paste_text.delete("1.0", "end")
        self.name_result_text.delete("1.0", "end")
        self.result_text.delete("1.0", "end")
        self._verify_field_label.config(text="")
        self._verifying_field = None
        self._page_count = 0
        self._last_verify_result = None
        self._last_verify_field_name = None
        self._last_paste_label = ""
        self._paste_sources.clear()
        _ExtractServerHandler.verify_receiver = None
        self.overall_var.set("选择左侧记录开始核验")
        self._current_case = None
        self._current_case_id = None
        self._current_raw_case_id = None
        self._verify_result = None
        self._active_field = None
        self._active_field_label.config(text="（未选中）", foreground="#999")

    def _set_active_field(self, field_name, field_label=None):
        """Set the active field for shared engine search."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return
        self._active_field = field_name
        if not field_label:
            field_labels_map = {
                "city": "城市", "construction_year": "建造时间",
                "demolition_year": "拆除时间", "restoration_year": "修复时间",
                "building_status": "建筑状态", "spread_method": "传播方式",
                "quantity": "数量", "function": "功能", "designer": "设计师",
                "coordinates": "经纬度", "image_urls": "图片",
            }
            field_label = field_labels_map.get(field_name, field_name)
        self._active_field_label.config(text=f"{field_label}", foreground="#1565c0")
        self.overall_var.set(f"已选中「{field_label}」，点击下方搜索引擎开始")

    def _search_active_field(self, engine):
        """Use shared engine to search for the active field."""
        if not self._active_field:
            messagebox.showinfo("提示", "请先点击某个字段的「选择字段」按钮。")
            return
        field_name = self._active_field
        field_labels_map = {
            "city": "城市", "construction_year": "建造时间",
            "demolition_year": "拆除时间", "restoration_year": "修复时间",
            "building_status": "建筑状态", "spread_method": "传播方式",
            "quantity": "数量", "function": "功能", "designer": "设计师",
            "coordinates": "经纬度", "image_urls": "图片",
        }
        field_label = field_labels_map.get(field_name, field_name)
        if engine == "ds":
            self._verify_field_deepseek(field_name, field_label)
        else:
            self._open_field_search(field_name, field_label, engine)

    def _verify_active_field_gemini(self):
        """Use Gemini to verify the active field."""
        if not self._active_field:
            messagebox.showinfo("提示", "请先点击某个字段的「选择字段」按钮。")
            return
        field_name = self._active_field
        field_labels_map = {
            "city": "城市", "construction_year": "建造时间",
            "demolition_year": "拆除时间", "restoration_year": "修复时间",
            "building_status": "建筑状态", "quantity": "数量",
            "function": "功能", "designer": "设计师",
            "coordinates": "经纬度", "image_urls": "图片",
        }
        field_label = field_labels_map.get(field_name, field_name)
        self._verify_field_gemini(field_name, field_label)

    def _auto_complete_all(self):
        """Direct DeepSeek API auto-complete — result goes to paste_text area."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return
        case = self._current_case
        if not case.get("name"):
            messagebox.showinfo("提示", "建筑名称为空，无法补全。")
            return

        DS_API_KEY = config.get_llm("verify_assist")[2]

        # Build name block from all available name variants
        name_block = _build_name_block(case)

        # Query ALL fields except country and image_urls, even if they have values
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

        self.overall_var.set("正在调用 DeepSeek 一键补全...")
        self._verify_field_label.config(text="DeepSeek 正在补全缺失字段...")
        self.frame.update()

        def _run():
            result = None
            try:
                result = _call_yunwu(prompt, DS_API_KEY, model="deepseek-chat",
                                    temperature=0.3, max_tokens=8000, timeout=90)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.frame.after(0, lambda: messagebox.showerror("一键DS补全", f"API调用失败: {e}"))
                return

            def _apply(r=result):
                try:
                    self.paste_text.configure(state="normal")
                    self.paste_text.insert("end", f"\n\n=== DeepSeek一键补全结果 ===\n")
                    self.paste_text.insert("end", json.dumps(r, ensure_ascii=False, indent=2))
                    self.paste_text.see("end")
                    self.paste_text.configure(state="disabled")
                    self._last_paste_label = "一键DS补全"
                    self._paste_sources.clear()
                    self._paste_sources.add("一键DS补全")
                    self._save_paste_content("一键DS补全")
                    self._apply_auto_completions(r)
                except Exception as e:
                    import traceback
                    traceback.print_exc()
                    try:
                        self.frame.after(0, lambda: messagebox.showerror("一键DS补全", f"结果显示失败: {e}"))
                    except Exception:
                        pass
            try:
                self.frame.after(0, _apply)
            except Exception:
                pass

        threading.Thread(target=_run, daemon=True).start()

    def _auto_complete_gemini(self):
        """Direct Gemini API auto-complete — no browser, leverages Gemini's multilingual knowledge."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return
        case = self._current_case
        if not case.get("name"):
            messagebox.showinfo("提示", "建筑名称为空，无法补全。")
            return

        GEMINI_API_KEY = config.get_llm("field_completion")[2]

        self.overall_var.set("正在调用 Gemini 一键补全...")
        self._verify_field_label.config(text="Gemini 正在补全缺失字段...")
        self.frame.update()

        def _run():
            result = None
            try:
                result = auto_complete_with_gemini(case, GEMINI_API_KEY)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.frame.after(0, lambda: messagebox.showerror("一键G补全", f"API调用失败: {e}"))
                return

            # Capture result in default arg to avoid closure issues
            def _apply(r=result):
                try:
                    self.paste_text.configure(state="normal")
                    self.paste_text.insert("end", f"\n\n=== Gemini一键补全结果 ===\n")
                    self.paste_text.insert("end", json.dumps(r, ensure_ascii=False, indent=2))
                    self.paste_text.see("end")
                    self.paste_text.configure(state="disabled")
                    self._last_paste_label = "一键G补全"
                    self._paste_sources.clear()
                    self._paste_sources.add("一键G补全")
                    self._save_paste_content("一键G补全")
                    self._apply_auto_completions(r)
                except Exception as e:
                    import traceback
                    traceback.print_exc()
                    try:
                        self.frame.after(0, lambda: messagebox.showerror("一键G补全", f"结果显示失败: {e}"))
                    except Exception:
                        pass
            try:
                self.frame.after(0, _apply)
            except Exception:
                # Frame might have been destroyed; ignore
                pass

        threading.Thread(target=_run, daemon=True).start()

    def _fill_completions(self, result, label="一键补全"):
        """Fill fields from completions dict — no UI output to result_text."""
        if not isinstance(result, dict):
            return {}
        completions = result.get("completions", {})
        # If no completions wrapper, check if top-level keys are field names
        if not completions:
            completable = {"city", "construction_year", "demolition_year",
                          "restoration_year", "building_status", "quantity",
                          "function", "designer", "coordinates", "latitude", "longitude",
                          "spread_method"}
            if any(k in completable for k in result):
                completions = result
        if not completions:
            self.overall_var.set("未找到可补全的字段")
            self._verify_field_label.config(text="— 未找到补全信息")
            return {}

        completable = ["city", "construction_year", "demolition_year",
                       "restoration_year", "building_status", "quantity",
                       "function", "designer", "coordinates", "latitude", "longitude",
                       "spread_method"]
        updates = {}
        for f in completable:
            info = completions.get(f, {})
            if not isinstance(info, dict):
                continue
            val = info.get("value")
            if not val or str(val).strip() == "" or str(val).lower() == "not_found":
                continue
            val_str = str(val).strip()

            if f == "coordinates":
                parts = [p.strip() for p in val_str.replace("，", ",").split(",")]
                if len(parts) >= 2:
                    try:
                        updates["latitude"] = float(parts[0])
                        updates["longitude"] = float(parts[1])
                    except ValueError:
                        updates["latitude"] = val_str
                else:
                    updates["latitude"] = val_str
            elif f == "function":
                updates[f] = "寺庙" if "worship" in val_str.lower() else val_str
            elif f == "building_status":
                updates[f] = self._normalize_building_status(val_str)
            elif f == "image_urls":
                image_urls_list = val if isinstance(val, list) else [val_str]
                updates["image_urls"] = json.dumps(image_urls_list, ensure_ascii=False)
            elif f == "quantity":
                num_match = re.search(r'\d+', str(val_str))
                if num_match:
                    updates[f] = int(num_match.group(0))
            elif f in ("construction_year", "demolition_year", "restoration_year"):
                updates[f] = self._extract_year(val_str)
            else:
                updates[f] = val_str

        if not updates:
            self.overall_var.set("未找到可用的补全数据")
            self._verify_field_label.config(text="— 未找到补全数据")
            return {}

        try:
            db.update_verified_case(self._current_case_id, updates)
            db.sync_raw_to_case(self._current_raw_case_id)
            source_short = "DS" if "DeepSeek" in label else ("G" if "Gemini" in label else label)
            db.insert_verification_log(
                self._current_raw_case_id, "auto_complete",
                f"一键{source_short}补全",
                json.dumps(result, ensure_ascii=False))

            case_new = db.get_verified_case_by_id(self._current_case_id)
            if not case_new:
                return
            self._current_case = case_new
            lang = self.lang_var.get()
            for fn, row_info in self._field_rows.items():
                if fn == "name":
                    val = db.get_display_name(case_new, lang)
                else:
                    val = self._get_field_val(case_new, fn)
                row_info["var"].set(str(val) if val else "")

            for item in self.queue_tree.get_children():
                item_id = self.queue_tree.item(item, "values")[0]
                if _parse_tree_id(item_id) == self._current_case_id:
                    self.queue_tree.item(item, values=(
                        _tree_id("V", case_new["id"]), case_new.get("country", "") or "",
                        db.get_display_name(case_new, lang),
                    ))
                    break

            self.overall_var.set(f"{label}完成！已填入 {len(updates)} 个字段")
            self._verify_field_label.config(text=f"✓ {label} {len(updates)} 个字段")
            if case_new:
                self._refresh_image_preview(case_new)
        except Exception as e:
            self.overall_var.set(f"保存失败: {e}")
        return updates
        """Display auto-complete results in result_text and fill fields."""
        label = getattr(self, '_last_paste_label', '') or '一键补全'
        if not isinstance(result, dict):
            self._show_verify_error("补全", f"AI返回格式异常: {str(result)[:100]}")
            return

        completions = result.get("completions", {})
        if not completions:
            error_msg = result.get("error", "")
            notes = result.get("notes", "")
            msg = error_msg or notes or "未找到可补全的字段"
            self.overall_var.set(msg)
            self._verify_field_label.config(text=f"— {msg}")
            messagebox.showinfo(label, msg)
            return

        completable = ["city", "construction_year", "demolition_year",
                       "restoration_year", "building_status", "quantity",
                       "function", "designer", "coordinates", "latitude", "longitude",
                       "spread_method"]
        field_labels = {"city": "城市",
                        "construction_year": "建造时间", "demolition_year": "拆除时间",
                        "restoration_year": "修复时间", "building_status": "建筑状态",
                        "spread_method": "传播方式",
                        "quantity": "数量", "function": "功能", "designer": "设计师",
                        "coordinates": "经纬度", "latitude": "纬度", "longitude": "经度"}

        self.result_text.delete("1.0", "end")
        self.result_text.insert("end", f"【{label}】\n\n")

        for f in completable:
            info = completions.get(f, {})
            if not isinstance(info, dict):
                self.result_text.insert("end", f"— {field_labels.get(f, f)}：未找到\n")
                continue
            val = info.get("value")
            if not val or str(val).strip() == "" or str(val).lower() == "not_found":
                self.result_text.insert("end", f"— {field_labels.get(f, f)}：未找到\n")
                continue

            val_str = str(val).strip()

            if f == "coordinates":
                parts = [p.strip() for p in val_str.replace("，", ",").split(",")]
                if len(parts) >= 2:
                    try:
                        lat = float(parts[0])
                        lng = float(parts[1])
                        self.result_text.insert("end", f"✓ 经纬度：{lat}, {lng}\n")
                    except ValueError:
                        self.result_text.insert("end", f"✓ 经纬度：{val_str}\n")
                else:
                    self.result_text.insert("end", f"✓ 经纬度：{val_str}\n")
            elif f == "function":
                label = "寺庙" if "worship" in val_str.lower() else val_str
                self.result_text.insert("end", f"✓ 功能：{label}（原值：{val_str}）\n")
            elif f == "building_status":
                normalized = self._normalize_building_status(val_str)
                self.result_text.insert("end", f"✓ 建筑状态：{normalized}（原值：{val_str}）\n")
            elif f == "image_urls":
                image_urls_list = val if isinstance(val, list) else [val_str]
                self.result_text.insert("end", f"✓ 图片：{len(image_urls_list)} 个链接\n")
            elif f == "quantity":
                num_match = re.search(r'\d+', str(val_str))
                if num_match:
                    qty = int(num_match.group(0))
                    self.result_text.insert("end", f"✓ {field_labels.get(f, f)}：{qty}（提取自：{val_str}）\n")
                else:
                    self.result_text.insert("end", f"— {field_labels.get(f, f)}：无法解析数字（{val_str}）\n")
            elif f in ("construction_year", "demolition_year", "restoration_year"):
                year_val = self._extract_year(val_str)
                if year_val != val_str:
                    self.result_text.insert("end", f"✓ {field_labels.get(f, f)}：{year_val}（提取自：{val_str}）\n")
                else:
                    self.result_text.insert("end", f"✓ {field_labels.get(f, f)}：{year_val}\n")
            else:
                self.result_text.insert("end", f"✓ {field_labels.get(f, f)}：{val_str}\n")

            # 显示结论、依据和参考链接
            conclusion = info.get("conclusion", "")
            if conclusion:
                self.result_text.insert("end", f"  结论：{conclusion}\n")
            reasoning = info.get("reasoning", "")
            if reasoning:
                self.result_text.insert("end", f"  依据：{reasoning}\n")
            refs = info.get("references", [])
            if refs:
                if isinstance(refs, str):
                    refs = [refs]
                self.result_text.insert("end", "  参考链接：\n")
                for ref in refs:
                    if ref and str(ref).strip():
                        self.result_text.insert("end", f"    • {ref}\n")
            self.result_text.insert("end", "\n")

        notes = result.get("notes", "")
        if notes:
            self.result_text.insert("end", f"\n备注：{notes}\n")

        self._save_other_field_text(
            getattr(self, '_last_paste_label', '') or '一键补全')

        # Actually fill the fields
        self._fill_completions(result, label)

    def _refresh_image_preview(self, case):
        """Refresh the image thumbnail strip for the current case."""
        for w in self._image_preview_frame.winfo_children():
            w.destroy()

        image_urls, merged_path, _all_urls = merged_image_sources(case)

        def on_delete_image(src):
            """Soft-delete a dbimg:// image (move to trash for undo)."""
            if src.startswith("dbimg://"):
                try:
                    img_id = int(src.replace("dbimg://", ""))
                    trash_id = db.soft_delete_case_image(img_id)
                    if trash_id is None:
                        return
                    # Update image_path in raw_cases
                    current = db.get_raw_case_by_id(case["id"])
                    cur_path = current.get("image_path", "") if current else ""
                    refs = []
                    if cur_path:
                        try:
                            refs = json.loads(cur_path) if isinstance(cur_path, str) else cur_path
                            if not isinstance(refs, list):
                                refs = [refs]
                        except (json.JSONDecodeError, TypeError):
                            refs = []
                    refs = [r for r in refs if r != src]
                    db.update_raw_case(case["id"], {"image_path": json.dumps(refs, ensure_ascii=False) if refs else ""})
                    # Enable undo button
                    self._undo_btn.config(state="normal")
                    # Refresh if no db images left
                    remaining = db.get_images_for_raw_case(case["id"])
                    if not remaining and not image_urls:
                        self._refresh_image_preview(case)
                except Exception:
                    pass

        def on_reorder_images(new_dbimg_order):
            """Persist reordered dbimg:// list to image_path."""
            db.update_raw_case(case["id"], {"image_path": json.dumps(new_dbimg_order, ensure_ascii=False)})
            case["image_path"] = json.dumps(new_dbimg_order, ensure_ascii=False)
            self._refresh_image_preview(case)

        def on_crop_save(cropped_pil, src):
            """Save cropped image to data/images/ and add to case."""
            from io import BytesIO
            buf = BytesIO()
            # Determine format from the source extension
            ext = os.path.splitext(src.split("?")[0])[-1] if "." in src else ""
            fmt = "JPEG" if ext.lower() in (".jpg", ".jpeg", "") else "PNG"
            cropped_pil.save(buf, format=fmt)
            data = buf.getvalue()
            fn = f"cropped_{os.path.basename(src.rsplit('/', 1)[-1].split('?')[0][:80])}" if "/" in src else "cropped_image"
            if not fn.endswith((".jpg", ".jpeg", ".png")):
                fn += ".jpg"
            img_id = db.insert_case_image(case["id"], data, filename=fn, source_url=src)
            new_ref = f"dbimg://{img_id}"
            cur_path = case.get("image_path", "") or ""
            refs = []
            if cur_path:
                try:
                    refs = json.loads(cur_path) if isinstance(cur_path, str) else cur_path
                    if not isinstance(refs, list):
                        refs = [refs]
                except (json.JSONDecodeError, TypeError):
                    refs = []
            refs.append(new_ref)
            db.update_raw_case(case["id"], {"image_path": json.dumps(refs, ensure_ascii=False)})
            case["image_path"] = json.dumps(refs, ensure_ascii=False)
            db.sync_raw_to_case(case["id"])
            self._refresh_image_preview(case)
            self.overall_var.set("裁剪图片已保存。")

        if image_urls or merged_path:
            # Fetch per-image metadata from both case_images table and image_metadata JSON
            image_metadata = {}
            images = db.get_images_for_raw_case(case["id"])
            for img in images:
                key = f"dbimg://{img['id']}"
                image_metadata[key] = {"author": img.get("author", ""), "image_time": img.get("image_time", "")}
            # Merge in metadata from raw_cases.image_metadata JSON (for URL images)
            raw_meta = case.get("image_metadata", "")
            if raw_meta:
                try:
                    all_meta = json.loads(raw_meta)
                    for src, meta in all_meta.items():
                        if src not in image_metadata:
                            image_metadata[src] = meta
                        else:
                            if not image_metadata[src].get("author") and meta.get("author"):
                                image_metadata[src]["author"] = meta["author"]
                            if not image_metadata[src].get("image_time") and meta.get("image_time"):
                                image_metadata[src]["image_time"] = meta["image_time"]
                except (json.JSONDecodeError, TypeError):
                    pass
            ImageThumbnailStrip(self._image_preview_frame, image_urls, merged_path,
                               on_delete=on_delete_image,
                               on_reorder=on_reorder_images,
                               case_id=case["id"],
                               on_crop_save=on_crop_save,
                               image_metadata=image_metadata).pack(fill="x")

    def _undo_delete_image(self):
        """Restore the last soft-deleted image for the current case."""
        if not hasattr(self, "_current_case") or not self._current_case:
            return
        case_id = self._current_raw_case_id
        new_id = db.undo_delete_image(case_id)
        if new_id is None:
            # No more trash entries, disable button
            self._undo_btn.config(state="disabled")
            return
        new_ref = f"dbimg://{new_id}"
        # Update image_path
        current = db.get_raw_case_by_id(case_id)
        cur_path = current.get("image_path", "") if current else ""
        refs = []
        if cur_path:
            try:
                refs = json.loads(cur_path) if isinstance(cur_path, str) else cur_path
                if not isinstance(refs, list):
                    refs = [refs]
            except (json.JSONDecodeError, TypeError):
                refs = []
        refs.append(new_ref)
        db.update_raw_case(case_id, {"image_path": json.dumps(refs, ensure_ascii=False)})
        db.update_verified_case(self._current_case_id, {"image_path": json.dumps(refs, ensure_ascii=False)})
        # Check if more trash entries remain
        self._refresh_image_preview(self._current_case)

    def _download_image(self):
        """Download an image from URL and save to the current case (runs in background thread)."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showwarning("提示", "请先在左侧列表选中一个案例。")
            return
        url = self._img_url_var.get().strip()
        if not url:
            messagebox.showwarning("提示", "请输入图片URL。")
            return
        case_id = self._current_raw_case_id
        self._img_url_var.set("")
        # Disable button to prevent double-click
        self._download_btn.config(state="disabled", text="下载中...")

        def _do_download():
            try:
                from ui.image_components import fetch_image_bytes
                data, _ = fetch_image_bytes(url, timeout=30)
                if not data:
                    raise Exception("无法获取图片数据")
                import database as db_module
                img_id = db_module.insert_case_image(case_id, data, filename=url.rsplit("/", 1)[-1][:100] if "/" in url else "image")
                new_ref = f"dbimg://{img_id}"
                cur_path = self._current_case.get("image_path", "") or ""
                refs = []
                if cur_path:
                    try:
                        refs = json.loads(cur_path) if isinstance(cur_path, str) else cur_path
                        if not isinstance(refs, list):
                            refs = [refs]
                    except (json.JSONDecodeError, TypeError):
                        refs = []
                refs.append(new_ref)
                db.update_raw_case(case_id, {"image_path": json.dumps(refs, ensure_ascii=False)})
                db.update_verified_case(self._current_case_id, {"image_path": json.dumps(refs, ensure_ascii=False)})
                self._current_case["image_path"] = json.dumps(refs, ensure_ascii=False)
                db.sync_raw_to_case(case_id)

                def _on_success():
                    self._download_btn.config(state="normal", text="下载图片")
                    self._refresh_image_preview(self._current_case)
                    self.overall_var.set("图片已下载并保存。")
                self.frame.after(100, _on_success)
            except Exception as e:
                err_msg = str(e)
                def _on_error():
                    self._download_btn.config(state="normal", text="下载图片")
                    self.overall_var.set(f"下载失败：{err_msg}")
                self.frame.after(100, _on_error)

        threading.Thread(target=_do_download, daemon=True).start()

    def _select_local_image(self):
        """Select a local image file and save to the current case."""
        if not hasattr(self, "_current_case") or not self._current_case:
            self.overall_var.set("请先在左侧列表选中一个案例。")
            return
        path = filedialog.askopenfilename(
            title="选择图片文件",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.gif *.bmp *.webp"), ("所有文件", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "rb") as f:
                data = f.read()
            import database as db_module
            img_id = db_module.insert_case_image(self._current_raw_case_id, data,
                                              filename=os.path.basename(path))
            new_ref = f"dbimg://{img_id}"
            cur_path = self._current_case.get("image_path", "") or ""
            refs = []
            if cur_path:
                try:
                    refs = json.loads(cur_path) if isinstance(cur_path, str) else cur_path
                    if not isinstance(refs, list):
                        refs = [refs]
                except (json.JSONDecodeError, TypeError):
                    refs = []
            refs.append(new_ref)
            db.update_raw_case(self._current_raw_case_id, {"image_path": json.dumps(refs, ensure_ascii=False)})
            db.update_verified_case(self._current_case_id, {"image_path": json.dumps(refs, ensure_ascii=False)})
            self._current_case["image_path"] = json.dumps(refs, ensure_ascii=False)
            db.sync_raw_to_case(self._current_raw_case_id)
            self._refresh_image_preview(self._current_case)
            self.overall_var.set("图片已添加。")
        except Exception as e:
            self.overall_var.set(f"添加失败：{e}")

    # ── Baidu image search ──

    def _fetch_baidu_image_urls(self, query, count=15):
        """Call Baidu image AJAX API, return list of (thumb_url, download_url) tuples."""
        import requests
        sess = requests.Session()
        sess.headers.update({
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept-Language": "zh-CN,zh;q=0.9",
        })
        # Get cookies from main page and search page
        try:
            sess.get("https://image.baidu.com/", timeout=10)
            search_url = "https://image.baidu.com/search/index?tn=baiduimage&word=" + urllib.parse.quote(query)
            sess.get(search_url, timeout=10)
        except Exception:
            pass

        results = []
        pn = 0
        while len(results) < count:
            try:
                r = sess.get("https://image.baidu.com/search/acjson", params={
                    "tn": "resultjson_com", "word": query, "pn": pn, "rn": min(count, 30)
                }, headers={"Referer": search_url}, timeout=15)
                text = r.text
                if "try{" in text:
                    text = text[text.index("try{") + 4:]
                    if text.rstrip().endswith("}catch(e){}"):
                        text = text[:text.rindex("}catch(e){}")]
                obj_start = text.find('{"')
                if obj_start >= 0:
                    text = text[obj_start:]
                data = json.loads(text)
                if data.get("antiFlag"):
                    break
                items = data.get("data", [])
                if not items:
                    break
                for item in items:
                    thumb = item.get("thumbURL", "")
                    if not thumb:
                        continue
                    # Prefer replaceUrl ObjURL for full-size, fallback to middleURL
                    dl_url = ""
                    replace_urls = item.get("replaceUrl", [])
                    if replace_urls:
                        dl_url = replace_urls[0].get("ObjURL") or replace_urls[0].get("ObjUrl") or ""
                    if not dl_url:
                        dl_url = item.get("middleURL", "") or thumb
                    results.append((thumb, dl_url))
                    if len(results) >= count:
                        break
                pn += len(items)
            except Exception:
                break
        return results

    def _show_baidu_image_dialog(self, query, urls, source="百度"):
        """Show a dialog with thumbnails + checkboxes for user to select images.
        Returns list of selected (obj_url) strings, or empty list if cancelled."""
        if not urls:
            messagebox.showinfo("提示", "未找到相关图片。")
            return []

        dialog = tk.Toplevel(self.frame)
        dialog.title(f"{source}图片搜索结果 — {query}")
        dialog.geometry("720x550")
        dialog.transient(self.frame)
        dialog.grab_set()

        # Scrollable thumbnail area
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side="right", fill="y", pady=10)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        dialog.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        check_vars = []
        thumb_labels = []

        from PIL import Image, ImageTk
        import requests
        from io import BytesIO

        # Progress indicator
        progress_var = tk.StringVar(value=f"加载缩略图中... 0/{len(urls)}")
        ttk.Label(scroll_frame, textvariable=progress_var, font=("", 9)).pack(pady=5)

        for i, (thumb_url, obj_url) in enumerate(urls):
            var = tk.BooleanVar(value=False)
            check_vars.append((var, obj_url))

            row = ttk.Frame(scroll_frame)
            row.pack(fill="x", pady=3, padx=5)

            # Checkbox
            cb = ttk.Checkbutton(row, variable=var)
            cb.pack(side="left", padx=(0, 5))

            # Thumbnail — handle both HTTP URLs and local file paths
            try:
                is_local = thumb_url.startswith("/") or not thumb_url.startswith("http")
                if is_local:
                    img = Image.open(thumb_url)
                else:
                    r = requests.get(thumb_url, timeout=8, headers={
                        "User-Agent": "Mozilla/5.0",
                        "Referer": "https://image.baidu.com/"
                    })
                    if r.status_code != 200:
                        raise Exception("HTTP error")
                    img = Image.open(BytesIO(r.content))
                img.thumbnail((120, 120), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                lbl = ttk.Label(row, image=photo)
                lbl.image = photo
                lbl.pack(side="left")
                thumb_labels.append(lbl)
            except Exception:
                ttk.Label(row, text="[加载失败]", width=15).pack(side="left")

            # Source tag
            from urllib.parse import urlparse
            try:
                if obj_url.startswith("/") or not obj_url.startswith("http"):
                    host = "Gemini"
                else:
                    host = urlparse(obj_url).hostname or "unknown"
            except Exception:
                host = "Gemini"
            ttk.Label(row, text=f"来源: {host}", font=("", 7), foreground="#999").pack(side="left", padx=8)

            progress_var.set(f"加载缩略图中... {i+1}/{len(urls)}")

        progress_var.set(f"共 {len(urls)} 张图片，勾选想要下载的图片后点击下方按钮")

        # Bottom buttons
        btn_bar = ttk.Frame(dialog)
        btn_bar.pack(fill="x", pady=(0, 10), padx=10)

        result = {"selected": []}

        def _on_select_all():
            select_all = not all(v.get() for v, _ in check_vars)
            for var, _ in check_vars:
                var.set(select_all)

        ttk.Button(btn_bar, text="全选/取消全选", command=_on_select_all).pack(side="left", padx=2)

        def _on_download():
            result["selected"] = [url for var, url in check_vars if var.get()]
            if not result["selected"]:
                messagebox.showinfo("提示", "请至少勾选一张图片。")
                return
            dialog.destroy()

        ttk.Button(btn_bar, text="下载选中图片", command=_on_download).pack(side="right", padx=2)
        ttk.Button(btn_bar, text="取消", command=dialog.destroy).pack(side="right", padx=2)

        self.frame.wait_window(dialog)
        return result["selected"]

    def _download_selected_images(self, selected_urls):
        """Download selected images and save to current case via insert_case_image.

        Handles both HTTP URLs (Baidu) and local file paths (Gemini temp files).
        """
        if not selected_urls:
            return
        case = self._current_case
        case_id = case["id"]
        import requests
        downloaded = 0
        for obj_url in selected_urls:
            try:
                is_local = obj_url.startswith("/") or not obj_url.startswith("http")
                if is_local:
                    with open(obj_url, "rb") as f:
                        content = f.read()
                    fn = os.path.basename(obj_url)
                    source = "Gemini"
                else:
                    r = requests.get(obj_url, timeout=30, headers={
                        "User-Agent": "Mozilla/5.0",
                        "Referer": "https://image.baidu.com/"
                    })
                    if r.status_code != 200 or len(r.content) <= 500:
                        continue
                    content = r.content
                    fn = urllib.parse.unquote(obj_url.rsplit("/", 1)[-1][:100]) if "/" in obj_url else "image"
                    source = obj_url

                img_id = db.insert_case_image(case_id, content, filename=fn, source_url=source)
                new_ref = f"dbimg://{img_id}"
                cur_path = case.get("image_path", "") or ""
                refs = []
                if cur_path:
                    try:
                        refs = json.loads(cur_path) if isinstance(cur_path, str) else cur_path
                        if not isinstance(refs, list):
                            refs = [refs]
                    except (json.JSONDecodeError, TypeError):
                        refs = []
                refs.append(new_ref)
                db.update_raw_case(case_id, {"image_path": json.dumps(refs, ensure_ascii=False)})
                case["image_path"] = json.dumps(refs, ensure_ascii=False)
                downloaded += 1
            except Exception:
                continue
        if downloaded > 0:
            db.sync_raw_to_case(case_id)
            try:
                self._refresh_image_preview(case)
            except Exception:
                pass
            self.overall_var.set(f"已下载 {downloaded} 张图片。")
        else:
            self.overall_var.set("下载失败：没有图片下载成功。")

    def _baidu_image_search(self):
        """Open Baidu image search dialog, let user select images, download selected."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return
        case = self._current_case

        # Build query from all available name variants for better accuracy
        def _s(key):
            return (case.get(key) or "").strip()
        query_parts = []
        for key in ["name", "name_en", "name_official"]:
            v = _s(key)
            if v:
                query_parts.append(v)
        for key in ["country", "city"]:
            v = _s(key)
            if v:
                query_parts.append(v)
        # common_names / other_names may be JSON strings or lists
        for key in ["common_names", "other_names"]:
            raw = case.get(key)
            if raw:
                if isinstance(raw, list):
                    for item in raw:
                        query_parts.append(str(item))
                elif isinstance(raw, str) and raw.strip():
                    try:
                        items = json.loads(raw)
                        for item in items:
                            query_parts.append(str(item))
                    except (json.JSONDecodeError, TypeError):
                        query_parts.append(raw.strip())
        query = " ".join(p for p in query_parts if p)

        if not query.strip():
            messagebox.showinfo("提示", "案例名称为空，无法搜索图片。")
            return

        self.overall_var.set(f"正在搜索百度图片：「{query}」...")
        self.frame.update_idletasks()

        # Fetch in background thread to avoid blocking UI
        def _run():
            urls = self._fetch_baidu_image_urls(query, count=15)

            def _show():
                if not urls:
                    self.overall_var.set(f"百度图片搜索「{query}」未找到结果。")
                    messagebox.showinfo("提示", "未找到相关图片。")
                    return
                selected = self._show_baidu_image_dialog(query, urls)
                if selected:
                    self._download_selected_images(selected)
                else:
                    self.overall_var.set("已取消图片选择。")
            self.frame.after(0, _show)

        threading.Thread(target=_run, daemon=True).start()

    def _google_image_search(self):
        """Use Gemini API to search for building images, show in-app for selection."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return
        case = self._current_case
        if not case.get("name"):
            messagebox.showinfo("提示", "案例名称为空，无法搜索图片。")
            return

        GEMINI_API_KEY = config.get_llm("field_completion")[2]
        name_block = _build_name_block(case)
        country = (case.get("country") or "").strip()

        self.overall_var.set("Gemini 正在搜图，约需1-2分钟...")
        self._verify_field_label.config(text="Gemini 正在搜索10张图片，请稍候...")
        self.frame.update_idletasks()

        def _run():
            images = []
            try:
                images = search_images_with_gemini(name_block, country, GEMINI_API_KEY, count=10)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.frame.after(0, lambda: messagebox.showerror("谷歌搜图", f"API调用失败: {e}"))
                return

            def _show():
                if not images:
                    self.overall_var.set("Gemini 未找到图片")
                    self._verify_field_label.config(text="— 未找到图片")
                    messagebox.showinfo("提示", "Gemini 未找到相关图片，可尝试百度搜图。")
                    return
                selected = self._show_baidu_image_dialog(
                    f"{case.get('name', '')} ({country})", images, source="谷歌Gemini")
                if selected:
                    self._download_selected_images(selected)
                else:
                    self.overall_var.set("已取消图片选择。")
                # Clean up temp files that were not selected
                for path, _, _ in images:
                    if path not in selected:
                        try:
                            os.unlink(path)
                        except Exception:
                            pass
            self.frame.after(0, _show)

        threading.Thread(target=_run, daemon=True).start()

    def _on_select(self, event):
        sel = self.queue_tree.selection()
        if not sel:
            self._clear_detail()
            return

        case_id = _parse_tree_id(self.queue_tree.item(sel[0], "values")[0])
        case = db.get_verified_case_by_id(case_id)
        if not case:
            return

        self._current_case = case
        self._current_case_id = case_id
        self._current_raw_case_id = case.get("raw_case_id")

        # Populate field rows with current values
        lang = self.lang_var.get()
        _, _, field_methods = self._get_verify_stats(case_id)
        for field_name, row_info in self._field_rows.items():
            if field_name == "name":
                val = db.get_display_name(case, lang)
            else:
                val = self._get_field_val(case, field_name)
            row_info["var"].set(str(val) if val else "")
            # Show verification method and count
            methods = field_methods.get(field_name, [])
            if methods:
                parts = [f"{m}×{n}" for m, n in methods]
                status_text = " ".join(parts)
                row_info["status_label"].config(text=status_text, foreground="#2e7d32")
                self._field_statuses[field_name] = "verified"
            else:
                row_info["status_label"].config(text="—", foreground="#999")
                self._field_statuses[field_name] = "unverified"

        # Refresh image preview
        self._refresh_image_preview(case)

        # Clear paste area and results
        self.paste_text.configure(state="normal")
        self.paste_text.delete("1.0", "end")
        self.paste_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self._verify_field_label.config(text="")
        self._verifying_field = None
        self._page_count = 0
        self._last_verify_result = None
        self._last_verify_field_name = None

        # Clear verification receiver when switching cases
        _ExtractServerHandler.verify_receiver = None

        self.overall_var.set("点击字段旁的「百度」或「谷歌」按钮开始核验")

    def _on_lang_change(self):
        """Refresh display when language selector changes."""
        self.refresh()
        if hasattr(self, "_current_case") and self._current_case:
            case = self._current_case
            lang = self.lang_var.get()
            if "name" in self._field_rows:
                val = db.get_display_name(case, lang)
                self._field_rows["name"]["var"].set(str(val) if val else "")

    def _enable_field_edit(self):
        """Enable all field entries for editing."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return
        for field_name, row_info in self._field_rows.items():
            row_info["entry"].configure(state="normal")
        self.overall_var.set("字段已解锁，修改后点击「保存」")

    def _save_field_changes(self):
        """Save manual edits from the field entries to the database and log the change."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return

        updates = {}
        changes = []
        case = self._current_case

        for field_name, row_info in self._field_rows.items():
            new_val = row_info["var"].get().strip()
            if field_name == "coordinates":
                lat = case.get("latitude")
                lng = case.get("longitude")
                old_parts = []
                if lat is not None and lat != "":
                    old_parts.append(str(lat))
                if lng is not None and lng != "":
                    old_parts.append(str(lng))
                old_val = ", ".join(old_parts) if old_parts else ""
                if new_val != old_val:
                    # Parse "lat, lng" back to separate values
                    parts = [p.strip() for p in new_val.split(",") if p.strip()]
                    if len(parts) >= 1:
                        updates["latitude"] = parts[0]
                    else:
                        updates["latitude"] = None
                    if len(parts) >= 2:
                        updates["longitude"] = parts[1]
                    else:
                        updates["longitude"] = None
                    changes.append(("coordinates", old_val, new_val))
            else:
                old_val = case.get(field_name)
                if old_val is None:
                    old_val = ""
                old_val = str(old_val)

                if new_val != old_val:
                    updates[field_name] = new_val if new_val else None
                    changes.append((field_name, old_val, new_val))

        if not updates:
            self.overall_var.set("没有修改")
            return

        try:
            db.update_verified_case(self._current_case_id, updates)
            raw_id = self._current_raw_case_id
            for field_name, old_val, new_val in changes:
                db.insert_verification_log(
                    raw_id, "manual_edit",
                    f"手动修改 {field_name}: 「{old_val}」→「{new_val}」")
            case = db.get_verified_case_by_id(self._current_case_id)
            if case:
                self._current_case = case
            self.overall_var.set(f"已保存 {len(changes)} 处修改")

            # Sync to linked cases record
            linked_id = db.sync_raw_to_case(raw_id)
            if linked_id:
                self.overall_var.set(f"已保存 {len(changes)} 处修改 — 已同步到正式案例")
                self.organizer.db_panel.refresh()

            # Update the queue list to reflect new values
            lang = self.lang_var.get()
            for item in self.queue_tree.get_children():
                item_id = self.queue_tree.item(item, "values")[0]
                if _parse_tree_id(item_id) == self._current_case_id:
                    self.queue_tree.item(item, values=(
                        _tree_id("V", case["id"]), case.get("country", "") or "",
                        db.get_display_name(case, lang),
                    ))
                    break

        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            self.overall_var.set(f"保存失败: {e}")

    def _on_queue_double_click(self, event):
        """Double-click a case in the queue to open detail dialog."""
        sel = self.queue_tree.selection()
        if not sel:
            return
        case_id = _parse_tree_id(self.queue_tree.item(sel[0], "values")[0])
        case = db.get_verified_case_by_id(case_id)
        if case:
            _VerificationResultDialog(self.frame, case)

    def _open_verify_detail(self):
        """Button click to open verification detail dialog for selected case."""
        sel = self.queue_tree.selection()
        if not sel:
            return
        case_id = _parse_tree_id(self.queue_tree.item(sel[0], "values")[0])
        case = db.get_verified_case_by_id(case_id)
        if case:
            _VerificationResultDialog(self.frame, case)

    def _delete_selected(self):
        """Remove selected cases from the verification queue (multi-select)."""
        sel = self.queue_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择记录。")
            return
        names = []
        for item in sel:
            vals = self.queue_tree.item(item, "values")
            names.append(vals[2] if len(vals) > 2 else "?")
        name_preview = "、".join(n[:12] for n in names[:5])
        if len(names) > 5:
            name_preview += f" 等共{len(names)}条"
        if not messagebox.askyesno("确认删除", f"确定要删除以下 {len(names)} 条核验记录吗？\n\n{name_preview}\n\n删除后可从原始数据重新提交核验。"):
            return
        try:
            for item in sel:
                verified_case_id = _parse_tree_id(self.queue_tree.item(item, "values")[0])
                vc = db.get_verified_case_by_id(verified_case_id)
                if vc and vc.get("raw_case_id"):
                    db.set_raw_case_status(vc["raw_case_id"], "raw")
                db.delete_verified_case(verified_case_id)
            self._clear_detail()
            self.refresh()
            self.overall_var.set(f"已删除: {len(sel)} 条")
        except Exception as e:
            messagebox.showerror("操作失败", str(e))

    # ── Field verification methods ──

    def _open_field_search(self, field_name, field_label, engine):
        """Open Chrome with Baidu/Google search for the given field, then register as receiver."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return

        case = self._current_case
        name = case.get("name", "")
        country = case.get("country", "")
        city = case.get("city", "")

        query_parts = [name]
        if field_name not in ("name",):
            if country:
                query_parts.append(country)
            if city:
                query_parts.append(city)
            query_parts.append(field_label)
        query = " ".join(p for p in query_parts if p)

        if engine == "baidu":
            url = f"https://www.baidu.com/s?wd={urllib.parse.quote(query)}"
        else:
            url = f"https://www.google.com/search?q={urllib.parse.quote(query)}"

        self._verifying_field = field_name
        self._page_count = 0
        engine_labels = {"baidu": "百度搜索插件回传", "google": "谷歌搜索插件回传"}
        src = engine_labels.get(engine, "浏览器搜索插件回传")
        self._paste_sources.clear()
        self._paste_sources.add(src)
        self._verify_field_label.config(
            text=f"等待接收「{field_label}」的搜索结果... 请在Chrome中查看搜索结果后，点击插件图标回传内容（可多次点击追加多个页面）")

        def receive_for_verify(text, url):
            self.frame.after(0, lambda: self._handle_verify_receive(field_name, field_label, text, url))

        _ExtractServerHandler.verify_receiver = receive_for_verify

        subprocess.run(["open", "-a", "Google Chrome", url])

        self.overall_var.set(f"已打开 {engine} 搜索「{field_label}」— 查看结果后点击Chrome插件图标回传")

    def _verify_field_gemini(self, field_name, field_label):
        """Call Gemini via Yunwu API to verify a field directly (no browser needed)."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return

        case = self._current_case
        current_val = self._get_field_val(case, field_name)
        if not current_val:
            current_val = "（空，待补充）"

        GEMINI_API_KEY = config.get_llm("field_completion")[2]

        self.overall_var.set(f"正在调用 Gemini 核验「{field_label}」...")
        self._verify_field_label.config(text=f"Gemini 正在核验「{field_label}」...")
        self.frame.update()

        def _run():
            try:
                result = verify_field_gemini(
                    field_name, field_label,
                    str(current_val), case, GEMINI_API_KEY,
                )
            except Exception as e:
                self.frame.after(0, lambda: self._show_verify_error(field_label, str(e)))
                return

            def _show():
                self._last_verify_result = result
                self._last_verify_field_name = field_name
                text_desc = result.pop("text_description", "") if isinstance(result, dict) else ""
                self._append_to_paste(field_name, field_label, result, text_desc, source=f"Gemini核验返回-{field_label}")
                self._show_verify_result(field_name, field_label, result, text_desc)

                try:
                    db.insert_verification_log(
                        self._current_raw_case_id, "field_verify_gemini",
                        f"Gemini核验: {field_label}",
                        json.dumps(result, ensure_ascii=False),
                    )
                except Exception:
                    pass

                self._update_log_display()

            self.frame.after(0, _show)

        threading.Thread(target=_run, daemon=True).start()

    def _verify_field_deepseek(self, field_name, field_label):
        """Call DeepSeek via API to verify a field directly (no browser needed)."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return

        case = self._current_case
        current_val = self._get_field_val(case, field_name)
        if not current_val:
            current_val = "（空，待补充）"

        DS_API_KEY = config.get_llm("verify_assist")[2]

        self.overall_var.set(f"正在调用 DeepSeek 核验「{field_label}」...")
        self._verify_field_label.config(text=f"DeepSeek 正在核验「{field_label}」...")
        self.frame.update()

        def _run():
            try:
                result = verify_field_deepseek(
                    field_name, field_label,
                    str(current_val), case, DS_API_KEY,
                )
            except Exception as e:
                self.frame.after(0, lambda: self._show_verify_error(field_label, str(e)))
                return

            def _show():
                self._last_verify_result = result
                self._last_verify_field_name = field_name
                # For name field: append to paste_text as search result content
                # (for comparison with Gemini results), not to name_result_text
                if field_name == "name":
                    self._append_to_paste(field_name, field_label, result, "", source=f"DeepSeek核验返回-{field_label}")
                else:
                    self._append_to_paste(field_name, field_label, result, "", source=f"DeepSeek核验返回-{field_label}")
                    self._show_verify_result(field_name, field_label, result, "")

                try:
                    db.insert_verification_log(
                        self._current_raw_case_id, "field_verify_deepseek",
                        f"DeepSeek核验: {field_label}",
                        json.dumps(result, ensure_ascii=False),
                    )
                except Exception:
                    pass

                self._update_log_display()

            self.frame.after(0, _show)

        threading.Thread(target=_run, daemon=True).start()

    def _append_to_paste(self, field_name, field_label, result, text_desc, source="核验返回"):
        """Append verification result to the paste text area."""
        self.paste_text.configure(state="normal")
        self.paste_text.insert("end", f"\n\n=== {field_label} 核验结果 ===\n")
        if text_desc:
            self.paste_text.insert("end", text_desc + "\n")
        if isinstance(result, dict):
            self.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
        else:
            self.paste_text.insert("end", str(result))
        self.paste_text.see("end")
        self.paste_text.configure(state="disabled")
        self._paste_sources.clear()
        if "DeepSeek" in source:
            self._last_paste_label = "DS核验"
            self._save_paste_content("DS核验")
            self._paste_sources.add("DS核验")
        elif "Gemini" in source:
            self._last_paste_label = "G核验"
            self._save_paste_content("G核验")
            self._paste_sources.add("G核验")
        else:
            self._last_paste_label = source
            self._save_paste_content(source)
            self._paste_sources.add(source)

    def _open_field_deepseek(self, field_name, field_label):
        """Open DeepSeek Chat in Chrome, auto-fill the prompt into the dialog."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return

        case = self._current_case
        current_val = self._get_field_val(case, field_name)
        if not current_val:
            current_val = "（空，待补充）"

        if field_name == "name":
            country = case.get("country") or "未知"
            prompt = FIELD_VERIFY_DEEPSEEK_NAME_PROMPT.format(
                country=country,
                current_value=str(current_val),
            )
        else:
            prompt = FIELD_VERIFY_DEEPSEEK_PROMPT.format(
                field_label=field_label,
                current_value=str(current_val),
            )

        self.frame.clipboard_clear()
        self.frame.clipboard_append(prompt)

        self._verifying_field = field_name
        self._page_count = 0
        self._verify_field_label.config(
            text=f"正在打开 DeepSeek 并自动填入核验问题...")

        def receive_for_verify(text, url):
            self.frame.after(0, lambda: self._handle_verify_receive(field_name, field_label, text, url))

        _ExtractServerHandler.verify_receiver = receive_for_verify

        ds_url = "https://chat.deepseek.com"
        subprocess.run(["open", "-a", "Google Chrome", ds_url])

        self.overall_var.set(
            f"已打开 DeepSeek — 正在自动填入核验「{field_label}」的问题..."
        )

        def _auto_fill():
            prompt_js = json.dumps(prompt, ensure_ascii=False)
            js_code = (
                '(function(){'
                'var ta=document.querySelector("textarea");'
                'if(!ta)return "no-textarea";'
                'var s=Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype,"value").set;'
                f's.call(ta,{prompt_js});'
                'ta.dispatchEvent(new Event("input",{bubbles:true}));'
                'ta.dispatchEvent(new Event("change",{bubbles:true}));'
                'setTimeout(function(){'
                'var btns=document.querySelectorAll("button");'
                'for(var i=0;i<btns.length;i++){'
                'if(!btns[i].disabled&&btns[i].offsetParent!==null){btns[i].click();break;}'
                '}'
                '},800);'
                'return "ok";'
                '})()'
            )
            as_script = f'''
            tell application "Google Chrome"
                activate
                repeat 10 times
                    try
                        tell active tab of front window
                            set jsResult to execute javascript "{js_code.replace(chr(34), chr(92) + chr(34))}"
                            if jsResult is not "no-textarea" then return "ok"
                        end tell
                    end try
                    delay 1
                end repeat
                return "timeout"
            end tell
            '''
            try:
                result = subprocess.run(['osascript', '-e', as_script],
                                      capture_output=True, text=True, timeout=20)
                if result.stdout.strip() == "ok":
                    self.frame.after(0, lambda: self._verify_field_label.config(
                        text="已在 DeepSeek 对话框中填入问题并发送 — 等待回复后点击插件图标回传"))
                    self.frame.after(0, lambda: self.overall_var.set(
                        f"DeepSeek 已发送「{field_label}」核验问题，等待回复后点击Chrome插件图标回传结果"))
                else:
                    self.frame.after(0, lambda: self._verify_field_label.config(
                        text="自动填入失败，请手动在DeepSeek对话框中 Cmd+V 粘贴"))
            except Exception as e:
                self.frame.after(0, lambda: self._verify_field_label.config(
                    text=f"自动填入失败，请手动 Cmd+V 粘贴 — {e}"))

        threading.Thread(target=_auto_fill, daemon=True).start()

    def _handle_verify_receive(self, field_name, field_label, text, url):
        """Called when extension sends page content while verification mode is active."""
        self.paste_text.configure(state="normal")
        existing = self.paste_text.get("1.0", "end-1c").strip()
        if existing:
            self.paste_text.insert("end", f"\n\n--- [来源] {url} ---\n\n{text}")
        else:
            self.paste_text.insert("1.0", f"--- [来源] {url} ---\n\n{text}")

        total_chars = len(self.paste_text.get("1.0", "end-1c"))
        self._page_count = getattr(self, "_page_count", 0) + 1
        self._verify_field_label.config(
            text=f"✓ 已收到 {self._page_count} 个页面（累计 {total_chars} 字符）— 可继续浏览并点击插件图标追加更多页面")

        self.overall_var.set(f"「{field_label}」已收集 {self._page_count} 个搜索结果页 — 收集完毕后点击「AI分析」")

        # 保存搜索结果内容到核验日志，核验详情页可查看
        try:
            db.insert_verification_log(
                self._current_raw_case_id, "search_result",
                f"搜索结果: {field_label} (来源: {url[:100]})",
                json.dumps({"search_content": text[:5000], "field": field_name, "field_label": field_label, "url": url}, ensure_ascii=False),
            )
        except Exception:
            pass

    def _show_verify_result(self, field_name, field_label, result, text_description=""):
        """Display verification result in the appropriate text area."""
        status = result.get("status", "unverified")
        case_val = self._get_field_val(self._current_case, field_name) if self._current_case else ""

        correct_value = result.get("correct_value") or result.get("web_value", "")
        suggestion = result.get("suggestion") or result.get("conclusion", "")
        detail = result.get("analysis") or result.get("reasoning", "")
        references = result.get("references", [])
        all_variations = result.get("all_variations", [])

        # Choose target text widget based on field_name
        target = self.name_result_text if field_name == "name" else self.result_text
        if field_name == "name":
            _prev_end_name = self.name_result_text.index("end-1c")
        else:
            _prev_end = self.result_text.index("end-1c")

        # Add separator before appending new result (if there's existing content)
        existing = target.get("1.0", "end-1c").strip()
        if existing:
            target.insert("end", "\n\n" + "─" * 50 + "\n\n")

        # Extract year for year-type fields
        display_value = str(correct_value) if correct_value else ""
        if field_name in ("construction_year", "demolition_year", "restoration_year") and display_value:
            display_value = self._extract_year(display_value)

        target.insert("end", f"当前值：{case_val if case_val else '（空）'}\n")
        if display_value and display_value != str(case_val):
            if display_value != str(correct_value):
                target.insert("end", f"正确值：{display_value}（提取自：{correct_value}）\n")
            else:
                target.insert("end", f"正确值：{correct_value}\n")
        elif display_value:
            target.insert("end", f"正确值：{correct_value}（与当前一致）\n")

        if suggestion:
            target.insert("end", f"\n【结论】\n{suggestion}\n")

        if detail:
            target.insert("end", f"\n【分析依据】\n{detail}\n")

        if references:
            target.insert("end", f"\n【参考来源】\n")
            for ref in references:
                if isinstance(ref, str):
                    target.insert("end", f"  • {ref}\n")
                elif isinstance(ref, dict):
                    title = ref.get("title", "")
                    url = ref.get("url", "")
                    target.insert("end", f"  • {title} — {url}\n")

        chinese_name = result.get("chinese_name", "") or result.get("name_zh", "")
        english_name = result.get("english_name", "") or result.get("name_en", "")
        official_name = result.get("official_language_name", "") or result.get("name_official", "")
        common_names = result.get("common_names", [])
        other_names = result.get("other_names", [])
        if chinese_name or english_name or official_name or common_names or other_names:
            target.insert("end", "\n【名称详情】\n")
            if chinese_name:
                target.insert("end", f"  中文：{chinese_name}\n")
            if english_name:
                target.insert("end", f"  English：{english_name}\n")
            if official_name:
                target.insert("end", f"  官方语言：{official_name}\n")
            if common_names:
                target.insert("end", f"  常用名：{', '.join(common_names)}\n")
            if other_names:
                target.insert("end", f"  其他名：{', '.join(other_names)}\n")

        if all_variations:
            variation_counts = result.get("variation_counts", {})
            target.insert("end", "\n【名称变体统计】\n")
            if variation_counts:
                for var, count in sorted(variation_counts.items(), key=lambda x: -x[1] if isinstance(x[1], int) else 0):
                    target.insert("end", f"  • {var}：{count}次\n")
            else:
                target.insert("end", f"  变体列表：{', '.join(all_variations)}\n")

        if field_name != "name":
            new_text = self.result_text.get(_prev_end, "end-1c").strip()
            if new_text:
                self._save_other_field_text(
                    getattr(self, '_last_paste_label', '') or 'AI分析')
        else:
            new_text = self.name_result_text.get(_prev_end_name, "end-1c").strip()
            if new_text:
                self._save_name_field_text(
                    getattr(self, '_last_paste_label', '') or 'AI分析')

        status_texts = {
            "confirmed": ("✓ 已确认", "#2e7d32"),
            "conflict": ("✗ 有冲突", "#c62828"),
            "found": ("+ 新发现", "#1565c0"),
            "unverified": ("? 未验证", "#999"),
        }
        # After verification, refresh field status to show method+count
        if field_name in self._field_rows:
            _, _, field_methods = self._get_verify_stats(self._current_case_id)
            methods = field_methods.get(field_name, [])
            if methods:
                parts = [f"{m}×{n}" for m, n in methods]
                status_text = " ".join(parts)
                self._field_rows[field_name]["status_label"].config(text=status_text, foreground="#2e7d32")
                self._field_statuses[field_name] = "verified"
            else:
                status_text, status_color = status_texts.get(status, (status, "#999"))
                self._field_rows[field_name]["status_label"].config(text=status_text, foreground=status_color)
                self._field_statuses[field_name] = status

        self._verify_field_label.config(text=f"「{field_label}」核验完成 — {status_text}")
        self.overall_var.set(f"「{field_label}」核验完成 — {suggestion}")

    def _save_other_field_text(self, source_label):
        text = self.result_text.get("1.0", "end-1c").strip()
        if not text:
            return
        try:
            db.insert_verification_log(
                self._current_raw_case_id, "field_verify_text",
                source_label, text
            )
        except Exception:
            pass

    def _save_paste_content(self, source_label):
        content = self.paste_text.get("1.0", "end-1c").strip()
        if not content:
            return
        try:
            db.insert_verification_log(
                self._current_raw_case_id, "search_result",
                f"{source_label}",
                json.dumps({"search_content": content, "source": source_label}, ensure_ascii=False),
            )
        except Exception:
            pass

    def _save_name_field_text(self, source_label):
        text = self.name_result_text.get("1.0", "end-1c").strip()
        if not text:
            return
        try:
            db.insert_verification_log(
                self._current_raw_case_id, "field_verify_text",
                f"[名称] {source_label}", text
            )
        except Exception:
            pass

    def _run_auto_complete_analysis(self, text, api_key):
        """Analyze search results to find missing field values and auto-fill."""
        case = self._current_case
        name = case.get("name", "")
        country = case.get("country", "") or ""

        # All fields that can be auto-completed (excluding name and country)
        completable = ["city", "construction_year", "demolition_year",
                       "restoration_year", "building_status", "quantity",
                       "function", "designer", "coordinates", "image_urls"]
        field_labels = {"city": "城市",
                        "construction_year": "建造时间", "demolition_year": "拆除时间",
                        "restoration_year": "修复时间", "building_status": "建筑状态",
                        "quantity": "数量", "function": "功能", "designer": "设计师",
                        "coordinates": "经纬度", "image_urls": "图片"}

        prompt = f"""我是一名研究人员，正在补全一个海外中式建筑的信息。

建筑名称：{name}
所在国家：{country}

请从以下搜索结果中，提取该建筑以下所有字段的信息：
{', '.join(field_labels.get(f, f) for f in completable)}

搜索结果：
{text[:8000]}

请对每个字段，从搜索结果中提取信息：
1. 如果能找到，给出具体值
2. 如果找不到，标注为 "not_found"

以JSON返回：
{{"completions": {{"city": {{"value": "旧金山"}}, "coordinates": {{"value": "37.79, -122.41"}}, "image_urls": {{"value": "http://..."}}, ...}}, "notes": "备注"}}
只返回JSON。"""

        self.overall_var.set("AI 正在分析搜索结果，补全缺失字段...")
        self._verify_field_label.config(text="⏳ AI 分析中...")

        def _run():
            try:
                result = _call_yunwu(prompt, api_key)
            except Exception as e:
                self.frame.after(0, lambda: self._show_verify_error("一键补全", str(e)))
                return

            def _apply():
                self._last_paste_label = "AI补全"
                self._apply_auto_completions(result)
            self.frame.after(0, _apply)

        threading.Thread(target=_run, daemon=True).start()

    def _parse_and_fill_from_text(self):
        """解析粘贴内容中所有JSON（多来源），对比后填入。冲突时弹窗让用户选择。"""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return

        text = self.paste_text.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showinfo("提示", "请先粘贴搜索结果内容。")
            return

        # --- Step 1: extract all JSON sections ---
        sections = re.split(r'=== .*?(?:核验结果|补全结果)\s*===', text)
        headers = re.findall(r'(=== .*?(?:核验结果|补全结果)\s*===)', text)
        targets = []
        for i, header in enumerate(headers):
            section_text = sections[i + 1] if i + 1 < len(sections) else ""
            source_label = "未知来源"
            hl = header.lower()
            if "gemini" in hl or "g" in header:
                source_label = "Gemini"
            elif "ds" in hl or "deepseek" in hl or "百度" in header:
                source_label = "DeepSeek"
            targets.append((section_text, source_label))
        if not targets:
            targets = [(text, "未标记来源")]

        # --- Step 2: parse JSON from each section ---
        all_data = []  # [{data, source}, ...]
        for section_text, source_label in targets:
            if not section_text.strip():
                continue
            json_str = None
            md_json = re.search(r'```json\s*([\s\S]*?)\s*```', section_text)
            if md_json:
                json_str = md_json.group(1).strip()
            else:
                json_str, _, _ = _extract_balanced_json(section_text)
            if json_str:
                try:
                    data = json.loads(json_str)
                    all_data.append({"data": data, "source": source_label})
                except json.JSONDecodeError:
                    pass

        if not all_data:
            self.result_text.delete("1.0", "end")
            self.result_text.insert("end", text + "\n")
            self.overall_var.set("未在粘贴内容中找到结构化JSON数据")
            return

        # --- Step 3: collect field values from all sources ---
        field_labels = {"city": "城市",
                       "construction_year": "建造时间", "demolition_year": "拆除时间",
                       "restoration_year": "修复时间", "building_status": "建筑状态",
                       "quantity": "数量", "function": "功能", "designer": "设计师",
                       "coordinates": "经纬度", "image_urls": "图片",
                       "spread_method": "传播方式"}

        field_sources = {}  # field_name -> [{"value": ..., "confidence": ..., "source": ...}, ...]

        for entry in all_data:
            data = entry["data"]
            src = entry["source"]
            completions = data.get("completions", {})
            # If no completions wrapper, check if top-level keys are field names
            if not completions:
                if any(k in field_labels for k in data):
                    completions = data
            status = data.get("status", "")

            if completions:
                for f in field_labels:
                    info = completions.get(f, {})
                    if not isinstance(info, dict):
                        continue
                    val = info.get("value")
                    if not val or (isinstance(val, str) and (val.strip() == "" or val.lower() == "not_found")):
                        continue
                    if f not in field_sources:
                        field_sources[f] = []
                    field_sources[f].append({
                        "value": val,
                        "source": src,
                        "conclusion": info.get("conclusion", ""),
                        "reasoning": info.get("reasoning", ""),
                        "references": info.get("references", []),
                    })

            elif status in ("confirmed", "conflict", "found", "unverified"):
                # Single-field result: determine field from details or _verifying_field
                vf = getattr(self, "_verifying_field", None)
                # Also try to identify field from the "field" key in data
                data_field = data.get("field", "")
                if data_field and data_field in field_labels:
                    vf = data_field
                if vf and vf in field_labels:
                    cv = data.get("correct_value") or data.get("web_value", "")
                    if cv:
                        if vf not in field_sources:
                            field_sources[vf] = []
                        field_sources[vf].append({
                            "value": cv,
                            "source": src,
                            "conclusion": data.get("conclusion", ""),
                            "reasoning": data.get("reasoning", ""),
                            "references": data.get("references", []),
                            "current_value": data.get("current_value", ""),
                        })

        if not field_sources:
            self.overall_var.set("未找到可用的字段数据")
            self._verify_field_label.config(text="— 未找到数据")
            return

        # --- Step 4: display each source's results sequentially ---
        self.result_text.delete("1.0", "end")
        self.result_text.insert("end", f"共解析 {len(all_data)} 个来源\n")
        self.result_text.insert("end", "─" * 50 + "\n\n")

        all_updates = {}

        for entry in all_data:
            data = entry["data"]
            src = entry["source"]
            completions = data.get("completions", {})
            if not completions and any(k in field_labels for k in data):
                completions = data

            self.result_text.insert("end", f"【来源：{src}】\n")

            for f, label in field_labels.items():
                info = completions.get(f, {})
                if not isinstance(info, dict):
                    self.result_text.insert("end", f"  — {label}：未找到\n")
                    continue
                val = info.get("value")
                if not val or str(val).strip() == "" or str(val).lower() == "not_found":
                    self.result_text.insert("end", f"  — {label}：未找到\n")
                    continue

                val_str = str(val).strip()
                self.result_text.insert("end", f"  ✓ {label}：{val_str}\n")

                conclusion = info.get("conclusion", "")
                if conclusion:
                    self.result_text.insert("end", f"    结论：{conclusion}\n")
                reasoning = info.get("reasoning", "")
                if reasoning:
                    self.result_text.insert("end", f"    依据：{reasoning}\n")
                refs = info.get("references", [])
                if refs:
                    if isinstance(refs, str):
                        refs = [refs]
                    self.result_text.insert("end", "    参考链接：\n")
                    for ref in refs:
                        if ref and str(ref).strip():
                            self.result_text.insert("end", f"      • {ref}\n")

                # Collect for auto-apply (last source wins)
                self._collect_field_update(f, val_str, all_updates)

            notes = data.get("notes", "")
            if notes:
                self.result_text.insert("end", f"  备注：{notes}\n")
            self.result_text.insert("end", "\n")

        if self._paste_sources:
            src_label = "+".join(sorted(self._paste_sources))
        else:
            src_label = getattr(self, '_last_paste_label', '') or "字段分析"
        self._save_other_field_text(src_label)
        self._save_paste_content(src_label)
        self._paste_sources.clear()

        # --- Step 5: save structured results and apply updates ---
        if all_updates:
            # Save each source's structured results to verification log
            paste_content = self.paste_text.get("1.0", "end-1c").strip()
            for entry in all_data:
                data = entry["data"]
                src = entry["source"]
                completions = data.get("completions", {})
                if not completions and any(k in field_labels for k in data):
                    completions = data
                if completions:
                    try:
                        db.insert_verification_log(
                            self._current_raw_case_id, "info_analysis",
                            f"信息分析结果 ({src})",
                            json.dumps({
                                "completions": completions,
                                "search_content": paste_content[:5000],
                            }, ensure_ascii=False),
                        )
                    except Exception:
                        pass

            self._apply_field_updates(all_updates, skip_log=True)
        else:
            self.overall_var.set("未找到可用的补全数据")
        self._clear_paste_area()

    def _collect_field_update(self, field_name, raw_value, updates):
        """Process a field value through field-specific cleaning and add to updates dict."""
        if isinstance(raw_value, list):
            val_str = json.dumps(raw_value, ensure_ascii=False)
        elif not isinstance(raw_value, str):
            val_str = str(raw_value)
        else:
            val_str = raw_value.strip()

        if field_name == "coordinates":
            parts = [p.strip() for p in val_str.replace("，", ",").split(",")]
            if len(parts) >= 2:
                try:
                    updates["latitude"] = float(parts[0])
                    updates["longitude"] = float(parts[1])
                except ValueError:
                    updates["latitude"] = val_str
            else:
                updates["latitude"] = val_str
        elif field_name == "function":
            if "worship" in val_str.lower():
                updates[field_name] = "寺庙"
            else:
                updates[field_name] = val_str
        elif field_name == "quantity":
            num_match = re.search(r'\d+', str(val_str))
            if num_match:
                updates[field_name] = int(num_match.group(0))
        elif field_name in ("construction_year", "demolition_year", "restoration_year"):
            updates[field_name] = self._extract_year(val_str)
        elif field_name == "building_status":
            updates[field_name] = self._normalize_building_status(val_str)
        else:
            updates[field_name] = val_str

    def _show_field_conflict_dialog(self, conflicts, callback):
        """Show dialog for user to pick which source's value to use for each conflicting field."""
        dlg = tk.Toplevel(self.frame)
        dlg.title("字段冲突 — 请选择")
        dlg.geometry("700x500")
        dlg.transient(self.frame)
        dlg.grab_set()

        ttk.Label(dlg, text="以下字段在不同来源中有不同结论，请逐字段选择采纳哪个：",
                 font=("", 11, "bold")).pack(padx=15, pady=(15, 10))

        field_labels = {"city": "城市",
                       "construction_year": "建造时间", "demolition_year": "拆除时间",
                       "restoration_year": "修复时间", "building_status": "建筑状态",
                       "quantity": "数量", "function": "功能", "designer": "设计师",
                       "coordinates": "经纬度", "image_urls": "图片",
                       "spread_method": "传播方式"}

        canvas = tk.Canvas(dlg)
        scrollbar = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=660)
        canvas.configure(yscrollcommand=scrollbar.set)

        choices = {}  # field_name -> tk.StringVar

        for f, sources in conflicts.items():
            label = field_labels.get(f, f)
            group = ttk.LabelFrame(scroll_frame, text=f" {label} ", padding=8)
            group.pack(fill="x", padx=10, pady=5)

            var = tk.StringVar()
            choices[f] = var

            for i, s in enumerate(sources):
                val_text = str(s["value"])
                src = s.get("source", "")
                rb = ttk.Radiobutton(group, text=f"{src}：{val_text}",
                                    variable=var, value=f"{i}")
                rb.pack(anchor="w")
                if i == 0:
                    var.set(f"{i}")  # Default: first option

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _confirm():
            result = {}
            for f, var in choices.items():
                idx = int(var.get())
                result[f] = conflicts[f][idx]
            dlg.destroy()
            callback(result)

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确认应用", command=_confirm).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=dlg.destroy).pack(side="left", padx=5)

    def _apply_field_updates(self, updates, skip_log=False):
        """Write updates to DB and refresh UI."""
        if not updates:
            self.overall_var.set("未找到可用的字段数据")
            self._verify_field_label.config(text="— 未找到数据")
            return

        try:
            db.update_verified_case(self._current_case_id, updates)
            raw_id = self._current_raw_case_id
            db.sync_raw_to_case(raw_id)
            self.organizer.db_panel.refresh()
            if not skip_log:
                paste_content = self.paste_text.get("1.0", "end-1c").strip()
                db.insert_verification_log(
                    raw_id, "info_analysis",
                    f"信息分析补全 {len(updates)} 个字段: {', '.join(updates.keys())}",
                    json.dumps({"updates": updates, "search_content": paste_content[:5000]}, ensure_ascii=False))

            case_new = db.get_verified_case_by_id(self._current_case_id)
            if case_new:
                self._current_case = case_new
            lang = self.lang_var.get()
            for fn, row_info in self._field_rows.items():
                if fn == "name":
                    val = db.get_display_name(case_new, lang)
                else:
                    val = self._get_field_val(case_new, fn)
                row_info["var"].set(str(val) if val else "")
                row_info["entry"].configure(state="disabled")

            for item in self.queue_tree.get_children():
                item_id = self.queue_tree.item(item, "values")[0]
                if _parse_tree_id(item_id) == self._current_case_id:
                    self.queue_tree.item(item, values=(
                        _tree_id("V", case_new["id"]), case_new.get("country", "") or "",
                        db.get_display_name(case_new, lang),
                    ))
                    break

            self.overall_var.set(f"信息分析完成！已填入 {len(updates)} 个字段")
            self._verify_field_label.config(text=f"✓ 信息分析补全 {len(updates)} 个字段")
            if case_new:
                self._refresh_image_preview(case_new)
        except Exception as e:
            self.overall_var.set(f"保存失败: {e}")

    def _show_verify_error(self, field_label, error):
        """Show verification error in the appropriate area."""
        target = self.name_result_text if self._verifying_field == "name" else self.result_text
        target.delete("1.0", "end")
        target.insert("1.0", f"AI分析失败：{error}")
        self._verify_field_label.config(text=f"✗「{field_label}」AI分析失败")
        self.overall_var.set(f"「{field_label}」核验失败: {error}")

    def _run_field_ai_analysis(self):
        """调用AI API分析粘贴内容（适用于百度/谷歌搜索结果等非结构化文本）。"""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return

        text = self.paste_text.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showinfo("提示", "请先粘贴搜索结果内容，或通过Chrome插件自动回传。")
            return

        api_key = self._api_key.get().strip()
        if not api_key:
            messagebox.showinfo("提示", "请输入 API 密钥。")
            return

        field_name = self._verifying_field
        if not field_name:
            messagebox.showinfo("提示", "请先点击某个字段旁的「百度」或「谷歌」按钮。")
            return

        # --- Auto-complete mode ---
        if field_name == "__auto_complete__":
            self._run_auto_complete_analysis(text, api_key)
            return

        field_labels = {"name": "建筑名称", "country": "国家", "city": "城市",
                        "construction_year": "建造时间", "demolition_year": "拆除时间",
                        "restoration_year": "修复时间", "building_status": "建筑状态",
                        "quantity": "数量", "function": "功能", "designer": "设计师",
                        "coordinates": "经纬度", "image_urls": "图片"}
        field_label = field_labels.get(field_name, field_name)

        case = self._current_case
        current_val = self._get_field_val(case, field_name)
        case_context = {k: v for k, v in case.items() if v is not None and k != field_name}

        self.overall_var.set(f"正在AI分析「{field_label}」...")

        def _run():
            try:
                if field_name == "name":
                    result = verify_field_name_from_paste(
                        current_val,
                        case_context, text, api_key,
                    )
                else:
                    result = verify_field_from_paste(
                        field_name, field_label,
                        current_val,
                        case_context, text, api_key,
                    )
            except Exception as e:
                self.frame.after(0, lambda: self._show_verify_error(field_label, str(e)))
                return

            def _show():
                self._last_verify_result = result
                self._last_verify_field_name = field_name
                self._show_verify_result(field_name, field_label, result)

                try:
                    db.insert_verification_log(
                        self._current_raw_case_id, "field_verify_ai",
                        f"AI分析: {field_label}",
                        json.dumps(result, ensure_ascii=False),
                    )
                except Exception:
                    pass

                self._update_log_display()

            self.frame.after(0, _show)

        threading.Thread(target=_run, daemon=True).start()

    def _extract_all_name_results(self, text):
        """从粘贴文本中提取所有名称核验结果JSON，返回 [{result, source_label}, ...]"""
        import re as re_module
        all_results = []

        # Strategy 1: split by === ... 核验结果 === headers
        sections = re_module.split(r'=== .*?核验结果\s*===', text)
        headers = re_module.findall(r'(=== .*?核验结果\s*===)', text)

        # Each section follows its header; sections[0] is text before first header
        targets = []
        for i, header in enumerate(headers):
            section_text = sections[i + 1] if i + 1 < len(sections) else ""
            # Determine source from header text
            source_label = "未知来源"
            hl = header.lower()
            if "gemini" in hl or "g" in header:
                source_label = "Gemini"
            elif "ds" in hl or "deepseek" in hl or "百度" in header:
                source_label = "DeepSeek"
            # If header doesn't identify source, check content for Gemini/DS markers
            if source_label == "未知来源":
                section_lower = section_text.lower()
                if "json output" in section_lower or "architectural historian" in section_lower:
                    source_label = "Gemini"
                elif "deepseek" in section_lower or "百度" in section_text or "文字描述" in section_text:
                    source_label = "DeepSeek"
            targets.append((section_text, source_label))

        # If no headers found, treat entire text as one block
        if not targets:
            targets = [(text, "未标记来源")]

        for section_text, source_label in targets:
            if not section_text.strip():
                continue
            # Try ```json block first
            md_json = re_module.search(r'```json\s*([\s\S]*?)\s*```', section_text)
            json_str = None
            if md_json:
                json_str = md_json.group(1)
            else:
                # Try balanced {} with name-specific fields
                for pattern in [r'\{[^{}]*"status"[^{}]*"correct_value"[^{}]*\}',
                              r'\{[^{}]*"correct_value"[^{}]*"chinese_name"[^{}]*\}',
                              r'\{[^{}]*"status"[^{}]*\}',
                              r'\{[^{}]*"chinese_name"[^{}]*\}']:
                    obj_match = re_module.search(pattern, section_text, re_module.DOTALL)
                    if obj_match:
                        start = obj_match.start()
                        depth = 0
                        for i, ch in enumerate(section_text[start:], start):
                            if ch == '{':
                                depth += 1
                            elif ch == '}':
                                depth -= 1
                                if depth == 0:
                                    json_str = section_text[start:i + 1]
                                    break
                        break
                # Fallback: try every balanced {} block
                if not json_str:
                    for m in re_module.finditer(r'\{', section_text):
                        start = m.start()
                        depth = 0
                        for i, ch in enumerate(section_text[start:], start):
                            if ch == '{':
                                depth += 1
                            elif ch == '}':
                                depth -= 1
                                if depth == 0:
                                    try:
                                        candidate = json.loads(section_text[start:i + 1])
                                        if any(k in candidate for k in ["correct_value", "chinese_name", "status"]):
                                            json_str = section_text[start:i + 1]
                                    except json.JSONDecodeError:
                                        pass
                                    break
                        if json_str:
                            break

            if json_str:
                try:
                    result = json.loads(json_str)
                    if any(k in result for k in ["correct_value", "chinese_name"]):
                        all_results.append({"result": result, "source": source_label})
                except json.JSONDecodeError:
                    pass

        return all_results

    def _merge_name_results(self, all_results):
        """合并多个名称核验结果，后方来源覆盖前方的同名字段（Gemini优先）"""
        merged = {}
        # DS first, Gemini later (overrides)
        ordered = sorted(all_results, key=lambda r: 1 if r["source"] == "Gemini" else 0)
        for entry in ordered:
            r = entry["result"]
            for key in ["correct_value", "chinese_name", "english_name",
                       "official_language_name", "common_names", "other_names",
                       "status", "conclusion", "reasoning", "references"]:
                if r.get(key):
                    merged[key] = r[key]
        # Track which sources contributed
        merged["_sources"] = list(set(e["source"] for e in all_results))
        return merged

    def _show_name_conflict_dialog(self, name_groups, case, all_results):
        """多个核验来源对名称有不同结论时，弹出对话框让用户选择。"""

        def _apply_choice():
            idx = listbox.curselection()
            if not idx:
                messagebox.showinfo("提示", "请选择一个来源。")
                return
            chosen_cv = options[idx[0]]
            chosen_results = name_groups[chosen_cv]
            if len(chosen_results) == 1:
                chosen = chosen_results[0]
                # Also merge in fields from the other group if available
                for other_cv, other_results in name_groups.items():
                    if other_cv == chosen_cv:
                        continue
                    for r in other_results:
                        for k in ["english_name", "official_language_name", "common_names", "other_names"]:
                            if r.get(k) and not chosen.get(k):
                                chosen[k] = r[k]
                self._apply_name_result(chosen, case, self._current_case_id)
            else:
                merged = self._merge_name_results([{"result": r, "source": "multi"} for r in chosen_results])
                self._apply_name_result(merged, case, self._current_case_id)
            dlg.destroy()

        dlg = tk.Toplevel(self.frame)
        dlg.title("名称冲突 — 请选择")
        dlg.geometry("650x500")
        dlg.transient(self.frame)
        dlg.grab_set()

        ttk.Label(dlg, text="不同来源对建筑名称有不同结论，请选择采纳哪个：",
                 font=("", 11, "bold")).pack(padx=15, pady=(15, 10))

        # Build info for each group
        options = list(name_groups.keys())
        info_texts = []
        for cv, results in name_groups.items():
            sources = [r.get("_source", "未知") for r in results]
            info = f"名称：{cv}\n来源：{', '.join(sources)}\n"
            # Add key details from first result
            r0 = results[0]
            if r0.get("english_name"):
                info += f"英文名：{r0['english_name']}\n"
            if r0.get("official_language_name"):
                info += f"官方语言名：{r0['official_language_name']}\n"
            if r0.get("conclusion"):
                info += f"结论：{r0['conclusion'][:120]}\n"
            info_texts.append(info)

        listbox_frame = ttk.Frame(dlg)
        listbox_frame.pack(fill="both", expand=True, padx=15, pady=5)
        listbox = tk.Listbox(listbox_frame, selectmode="single", font=("", 10))
        scroll = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scroll.set)
        for i, info in enumerate(info_texts):
            listbox.insert("end", f"【选项 {i+1}】\n{info}")
        listbox.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="应用选中", command=_apply_choice).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=dlg.destroy).pack(side="left", padx=5)

    def _apply_name_result(self, result, case, case_id):
        """将名称核验结果写入数据库并刷新UI。"""
        correct_value = result.get("correct_value") or result.get("web_value", "")
        updates = {}
        current_name = case.get("name", "")

        if correct_value and str(correct_value) != str(current_name):
            updates["name"] = str(correct_value)
        if result.get("english_name"):
            updates["name_en"] = result["english_name"]
        if result.get("official_language_name"):
            updates["name_official"] = result["official_language_name"]

        if updates:
            try:
                db.update_verified_case(case_id, updates)
                db.sync_raw_to_case(case.get("raw_case_id", self._current_raw_case_id))
                for k, v in updates.items():
                    case[k] = v
                self._current_case = case
                if correct_value and "name" in self._field_rows:
                    self._field_rows["name"]["var"].set(str(correct_value))
                if "name" in updates:
                    self.overall_var.set(f"✓ 名称已修订：{current_name} → {correct_value}")
                    self._verify_field_label.config(text=f"✓ 名称已修订为「{correct_value}」")
                    if "name" in self._field_rows:
                        self._field_rows["name"]["status_label"].config(text="✓ 已修订", foreground="#2e7d32")
                else:
                    field_names = []
                    if "name_en" in updates:
                        field_names.append("英文名")
                    if "name_official" in updates:
                        field_names.append("官方语言名")
                    self.overall_var.set(f"✓ 名称与当前值一致，已更新{'、'.join(field_names)}")
                    self._verify_field_label.config(text=f"✓ 已更新{'、'.join(field_names)}")
            except Exception as e:
                self.overall_var.set(f"更新失败: {e}")
        elif correct_value:
            self.overall_var.set(f"名称「{correct_value}」与当前值一致，无需修订")
            if "name" in self._field_rows:
                self._field_rows["name"]["var"].set(str(correct_value))
        else:
            chinese_name = result.get("chinese_name", "")
            if chinese_name and chinese_name != current_name:
                if "name" in self._field_rows:
                    self._field_rows["name"]["var"].set(str(chinese_name))
                self._verify_field_label.config(text=f"请检查名称是否需要更新为「{chinese_name}」")

    def _run_name_field_analysis(self):
        """解析所有搜索结果中的名称核验JSON，对比后填入。多来源冲突时让用户选择。"""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先在左侧选择一条待核验记录。")
            return

        text = self.paste_text.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showinfo("提示", "请先粘贴搜索结果内容。")
            return

        case = self._current_case
        case_id = self._current_case_id

        # Step 1: extract all name results from paste text
        all_results = self._extract_all_name_results(text)

        if not all_results:
            self.overall_var.set("未在粘贴内容中找到名称核验结果JSON")
            self._verify_field_label.config(text="未找到名称核验JSON数据")
            return

        # Step 2: tag each result dict with its source
        for entry in all_results:
            entry["result"]["_source"] = entry["source"]

        # Step 3: group by correct_value
        name_groups = {}
        for entry in all_results:
            r = entry["result"]
            cv = r.get("correct_value") or r.get("web_value", "")
            if cv not in name_groups:
                name_groups[cv] = []
            name_groups[cv].append(r)

        # Step 4: display all results in name result area
        field_name = "name"
        field_label = "建筑名称"
        self._last_verify_result = all_results[0]["result"]
        self._last_verify_field_name = field_name
        # Show all results side by side
        self.name_result_text.delete("1.0", "end")
        for i, entry in enumerate(all_results):
            r = entry["result"]
            if r.get("correct_value"):
                self.name_result_text.insert("end", f"correct_value: {r['correct_value']}\n")
            if r.get("chinese_name"):
                self.name_result_text.insert("end", f"chinese_name: {r['chinese_name']}\n")
            if r.get("english_name"):
                self.name_result_text.insert("end", f"english_name: {r['english_name']}\n")
            if r.get("official_language_name"):
                self.name_result_text.insert("end", f"official_language_name: {r['official_language_name']}\n")
            if r.get("common_names"):
                cn = r["common_names"]
                self.name_result_text.insert("end", f"common_names: {cn if isinstance(cn, str) else ', '.join(cn)}\n")
            if r.get("other_names"):
                on = r["other_names"]
                self.name_result_text.insert("end", f"other_names: {on if isinstance(on, str) else ', '.join(on)}\n")
            if r.get("conclusion"):
                self.name_result_text.insert("end", f"结论：{r['conclusion'][:200]}\n")
            if r.get("reasoning"):
                self.name_result_text.insert("end", f"依据：{r['reasoning'][:200]}\n")
            if r.get("references"):
                refs = r["references"]
                if isinstance(refs, list):
                    self.name_result_text.insert("end", f"参考链接({len(refs)}条)：\n")
                    for ref in refs[:3]:
                        self.name_result_text.insert("end", f"  • {ref}\n")
                else:
                    self.name_result_text.insert("end", f"参考链接：{refs}\n")
            self.name_result_text.insert("end", "\n")

        # Step 5: handle conflicts
        if len(name_groups) > 1:
            self._verify_field_label.config(text=f"⚠️ {len(all_results)}个来源名称不一致，请在弹出的对话框中选择")
            self.overall_var.set(f"名称冲突：{len(name_groups)}种不同结论，请选择")
            self._show_name_conflict_dialog(name_groups, case, all_results)
        else:
            # All agree — merge and apply
            merged = self._merge_name_results(all_results)
            self._apply_name_result(merged, case, case_id)

        try:
            db.insert_verification_log(
                self._current_raw_case_id, "field_verify_ai",
                f"名称字段信息分析: {field_label}",
                json.dumps({
                    "results": [e["result"] for e in all_results],
                    "search_content": text[:5000],
                }, ensure_ascii=False),
            )
        except Exception:
            pass

        self._save_paste_content("名称字段信息分析")

        if self._paste_sources:
            src_label = "+".join(sorted(self._paste_sources))
        else:
            src_label = "名称字段信息分析"
        self._save_name_field_text(src_label)
        self._paste_sources.clear()

        self._update_log_display()
        self._clear_paste_area()

    def _apply_field_result(self):
        """Apply the most recent verification result to the database."""
        if not self._last_verify_result or not self._last_verify_field_name:
            messagebox.showinfo("提示", "没有可应用的核验结果。请先完成一次字段核验。")
            return

        result = self._last_verify_result
        field_name = self._last_verify_field_name
        status = result.get("status", "")

        if status == "confirmed":
            messagebox.showinfo("提示", f"该字段已确认正确，无需更新。")
            return

        new_value = result.get("correct_value") or result.get("web_value") or result.get("suggestion")
        if not new_value:
            messagebox.showinfo("提示", "该字段没有可应用的值。")
            return

        try:
            if field_name == "coordinates":
                parts = [p.strip() for p in new_value.split(",") if p.strip()]
                updates = {}
                if len(parts) >= 1:
                    updates["latitude"] = parts[0]
                if len(parts) >= 2:
                    updates["longitude"] = parts[1]
            else:
                updates = {field_name: new_value}
            db.update_verified_case(self._current_case_id, updates)
            raw_id = self._current_raw_case_id
            db.sync_raw_to_case(raw_id)
            db.insert_verification_log(raw_id, "apply_field",
                                       f"应用核验建议: {field_name} = {new_value}")

            if field_name in self._field_rows:
                self._field_rows[field_name]["var"].set(str(new_value))
                self._field_rows[field_name]["status_label"].config(text="✓ 已更新", foreground="#2e7d32")

            case = db.get_verified_case_by_id(self._current_case_id)
            if case:
                self._current_case = case

            self.overall_var.set(f"已应用: {field_name} = {new_value}")
            self._update_log_display()
            messagebox.showinfo("完成", f"已更新 {field_name}: {new_value}")
        except Exception as e:
            messagebox.showerror("更新失败", str(e))

    def _clear_paste_area(self):
        """Clear the paste text area."""
        self.paste_text.configure(state="normal")
        self.paste_text.delete("1.0", "end")
        self._page_count = 0

    def _save_manual_result(self):
        """Save manually entered result text to verification log."""
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return
        name_text = self.name_result_text.get("1.0", "end-1c").strip()
        other_text = self.result_text.get("1.0", "end-1c").strip()
        text = f"[名称字段]\n{name_text}\n\n[其他字段]\n{other_text}" if name_text or other_text else ""
        if not text.strip():
            messagebox.showinfo("提示", "核验结果为空，请先输入内容。")
            return
        try:
            result_json = json.dumps({
                "status": "manual",
                "conclusion": text,
                "current_value": "",
            }, ensure_ascii=False)
            source_label = ""
            db.insert_verification_log(
                self._current_raw_case_id, "manual_result",
                source_label, result_json)
            self.overall_var.set("核验结果已保存")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _clear_result(self):
        """Clear both result text areas."""
        self.name_result_text.delete("1.0", "end")
        self.result_text.delete("1.0", "end")

    def _stop_and_clear(self):
        """Stop receiving and clear everything."""
        _ExtractServerHandler.verify_receiver = None
        self._verifying_field = None
        self._page_count = 0
        self.paste_text.configure(state="normal")
        self.paste_text.delete("1.0", "end")
        self._verify_field_label.config(text="")
        self.overall_var.set("已停止接收，内容已清空 — 点击字段旁的按钮重新开始")

    def _clear_verify_receiver(self):
        """Clear the verify receiver and reset verifying state."""
        _ExtractServerHandler.verify_receiver = None
        self._verifying_field = None
        self._page_count = 0
        self._verify_field_label.config(text="")
        self.overall_var.set("已停止接收，插件内容将回到AI提取框")

    def _update_log_display(self):
        """Log display moved to _VerificationResultDialog."""
        pass

    def _batch_approve(self):
        """Batch approve selected cases in the queue."""
        sel = self.queue_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请在队列中勾选要批量通过的案例。")
            return
        names = []
        raw_case_ids = []
        for item in sel:
            vals = self.queue_tree.item(item, "values")
            verified_case_id = _parse_tree_id(vals[0])
            vc = db.get_verified_case_by_id(verified_case_id)
            if not vc:
                continue
            raw_case_id = vc.get("raw_case_id")
            if not raw_case_id:
                continue
            raw_case_ids.append(raw_case_id)
            names.append(vc.get("name", "?"))
        if not raw_case_ids:
            messagebox.showinfo("提示", "未找到有效的案例。")
            return
        name_preview = "、".join(n[:15] for n in names[:5])
        if len(names) > 5:
            name_preview += f" 等共{len(names)}条"
        if not messagebox.askyesno("批量通过确认",
            f"将把以下 {len(names)} 条案例转入正式数据库：\n\n{name_preview}\n\n确定批量通过？"):
            return
        success = 0
        errors = []
        for raw_case_id in raw_case_ids:
            try:
                new_id = db.promote_raw_to_case(raw_case_id)
                if new_id:
                    success += 1
                else:
                    errors.append(f"案例ID {raw_case_id}: promote_raw_to_case 返回 None")
            except Exception as e:
                errors.append(f"案例ID {raw_case_id}: {e}")
        self.refresh()
        self.raw_panel.refresh()
        self.organizer.db_panel.refresh()
        msg = f"批量通过完成：{success}/{len(raw_case_ids)} 条已转入正式数据库。"
        if errors:
            msg += f"\n\n失败 {len(errors)} 条：\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += f"\n... 等共{len(errors)}条失败"
        messagebox.showinfo("完成", msg)

    def _approve(self):
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条待核验记录。")
            return

        if messagebox.askyesno("确认", "通过核验后，该案例将被转入正式数据库，列表中保留并显示「已通过」。确定通过？"):
            try:
                new_id = db.promote_raw_to_case(self._current_raw_case_id)
                if new_id:
                    messagebox.showinfo("完成", f"已转入正式数据库，案例 ID: {new_id}")
                    self.refresh()
                    self.raw_panel.refresh()
                    self.organizer.db_panel.refresh()
                else:
                    messagebox.showerror("操作失败", "转入正式数据库失败")
            except Exception as e:
                messagebox.showerror("操作失败", str(e))

    def _reject(self):
        if not hasattr(self, "_current_case") or not self._current_case:
            messagebox.showinfo("提示", "请先选择一条待核验记录。")
            return

        if messagebox.askyesno("确认", "确定拒绝该案例吗？被拒绝的案例不会被删除，仅标记状态。"):
            db.set_verified_case_status(self._current_case_id, "rejected")
            db.insert_verification_log(self._current_raw_case_id, "reject", "核验被拒绝", "rejected")
            messagebox.showinfo("完成", "已标记为拒绝。")
            self.refresh()
            self.raw_panel.refresh()

    def _show_duplicates(self):
        """全字段匹配去重：找出所有字段完全一致的案例组，自动保留第一条，删除其余。"""
        dupes = db.find_verified_duplicates_all_fields()
        if not dupes:
            messagebox.showinfo("去重检查", "未发现重复记录。")
            return

        total_deleted = 0
        for d in dupes:
            ids = d["ids"]
            for dup_id in ids[1:]:
                db.delete_verified_case(dup_id)
                total_deleted += 1

        self.refresh()
        messagebox.showinfo("去重完成", f"共处理 {len(dupes)} 组重复记录，删除 {total_deleted} 条。")

    def _smart_dedup(self):
        """智能去重：按名称+国家分组，逐组弹出合并对话框。"""
        groups = db.find_verified_duplicates_by_name_country()
        if not groups:
            messagebox.showinfo("智能去重", "未发现名称+国家相同的重复记录。")
            return

        self._dedup_groups = groups
        self._dedup_idx = 0
        self._dedup_total = len(groups)
        self._dedup_merged = 0

        def _process_next():
            while self._dedup_idx < self._dedup_total:
                g = self._dedup_groups[self._dedup_idx]
                cases = []
                for rid in g["ids"]:
                    case = db.get_verified_case_by_id(rid)
                    if case:
                        cases.append(case)
                self._dedup_idx += 1
                if len(cases) >= 2:
                    dlg = _MergeDialog(self.frame, cases, on_done=_on_group_done,
                                       smart_defaults=True, table="verified")
                    dlg.win.title(
                        f"智能去重 ({self._dedup_idx}/{self._dedup_total})"
                        f" - {cases[0].get('name', '?')[:20]}"
                        f" ({len(cases)}条记录)")
                    return
            # All done
            msg = f"全部 {self._dedup_total} 组处理完成，共合并 {self._dedup_merged} 组。"
            messagebox.showinfo("智能去重", msg)
            self.refresh()

        def _on_group_done():
            self._dedup_merged += 1
            _process_next()

        _process_next()

    def _clean_data(self):
        if not messagebox.askyesno("数据清洗", "将执行以下操作：\n- 去除所有文本字段首尾空白\n\n确定继续？"):
            return

        cases = db.get_all_verified_cases()
        cleaned = 0
        text_fields = [f[0] for f in db.RAW_FIELDS if "TEXT" in f[2]]
        for c in cases:
            updates = {}
            for field in text_fields:
                val = c.get(field)
                if isinstance(val, str) and val.strip() != val:
                    updates[field] = val.strip()
            if updates:
                db.update_verified_case(c["id"], updates)
                cleaned += 1

        messagebox.showinfo("清洗完成", f"已清洗 {cleaned} 条记录。")
        self.refresh()

    def _send_to_batch(self):
        sel = self.queue_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在队列中选中案例。")
            return
        # Build case list directly from selected queue items
        cases = []
        ids = set()
        for item in sel:
            vals = self.queue_tree.item(item, "values")
            if vals:
                cid = _parse_tree_id(vals[0])
                country = vals[1] if len(vals) > 1 else "?"
                name = vals[2] if len(vals) > 2 else "?"
                cases.append((cid, name, country))
                ids.add(cid)
        # Rebuild batch tree with only selected cases
        for item in self._batch_tree.get_children():
            self._batch_tree.delete(item)
        self._batch_rows.clear()
        self._batch_status.clear()
        self._batch_pending = cases

        for cid, name, country in cases:
            name_status, name_reason, fields_status, fields_reason = _load_batch_case_status(cid)
            item = self._batch_tree.insert("", "end", text=name, values=(country, name_status, fields_status))
            self._batch_rows[cid] = item
            self._batch_status[cid] = {
                "name": (name_status, name_reason),
                "fields": (fields_status, fields_reason)
            }

        self._batch_progress_var.set(f"已选 {len(self._batch_pending)} 个案例")

    def _open_batch_verify(self):
        # Try to restore last checkpoint
        if not self._batch_pending:
            self._restore_batch_checkpoint()
        if not self._batch_pending:
            self._load_batch_cases()
        self._save_batch_checkpoint()
    # ── Batch verify methods ──

    def _build_batch_verify(self, parent):
        self._batch_paused = False
        self._batch_stopped = False
        self._batch_running = False
        self._batch_pending = []
        self._batch_rows = {}  # case_id -> tree item id
        self._batch_status = {}  # case_id -> {name: (status,reason), fields: (status,reason)}

        # Row 1: dropdowns + log button
        ctrl1 = ttk.Frame(parent)
        ctrl1.pack(fill="x", pady=(5, 2))

        ttk.Label(ctrl1, text="范围：").pack(side="left")
        self._batch_scope_var = tk.StringVar(value="全部待核验")
        ttk.Combobox(ctrl1, textvariable=self._batch_scope_var,
                     values=["全部待核验", "选中案例"],
                     state="readonly", width=12).pack(side="left", padx=3)

        ttk.Label(ctrl1, text="名称：").pack(side="left", padx=(8, 0))
        self._batch_name_model = tk.StringVar(value="G")
        ttk.Combobox(ctrl1, textvariable=self._batch_name_model,
                     values=["G", "DS"], state="readonly", width=4).pack(side="left", padx=2)

        ttk.Label(ctrl1, text="字段：").pack(side="left", padx=(8, 0))
        self._batch_fields_model = tk.StringVar(value="DS")
        ttk.Combobox(ctrl1, textvariable=self._batch_fields_model,
                     values=["DS", "G"], state="readonly", width=4).pack(side="left", padx=2)

        ttk.Button(ctrl1, text="日志", command=self._open_batch_log).pack(side="right", padx=4)

        # Row 2: buttons + progress
        ctrl2 = ttk.Frame(parent)
        ctrl2.pack(fill="x", pady=(2, 4))

        self._batch_start_btn = ttk.Button(ctrl2, text="开始", command=self._start_batch)
        self._batch_start_btn.pack(side="left")
        self._batch_pause_btn = ttk.Button(ctrl2, text="暂停", command=self._pause_batch, state="disabled")
        self._batch_pause_btn.pack(side="left", padx=3)
        self._batch_stop_btn = ttk.Button(ctrl2, text="停止", command=self._stop_batch, state="disabled")
        self._batch_stop_btn.pack(side="left", padx=3)

        self._batch_progress_var = tk.StringVar(value="")
        ttk.Label(ctrl2, textvariable=self._batch_progress_var).pack(side="left", padx=8)

        self._batch_progress_bar = ttk.Progressbar(ctrl2, length=160, mode="determinate")
        self._batch_progress_bar.pack(side="left", padx=3)

        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True, pady=4)

        columns = ("country", "name_status", "fields_status")
        self._batch_tree = ttk.Treeview(tree_frame, columns=columns,
                                        show="tree headings", selectmode="extended")
        self._batch_tree.heading("#0", text="案例名")
        self._batch_tree.column("#0", width=200)
        self._batch_tree.heading("country", text="国家")
        self._batch_tree.column("country", width=60)
        self._batch_tree.heading("name_status", text="名称")
        self._batch_tree.column("name_status", width=80)
        self._batch_tree.heading("fields_status", text="其他字段")
        self._batch_tree.column("fields_status", width=80)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self._batch_tree.yview)
        self._batch_tree.configure(yscrollcommand=scrollbar.set)
        self._batch_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self._batch_tree.bind("<<TreeviewSelect>>", self._on_batch_select)
        self._batch_tree.bind("<Button-2>" if sys.platform == "darwin" else "<Button-3>",
                              self._on_batch_right_click)

        bottom = ttk.Frame(parent)
        bottom.pack(fill="x", pady=(4, 5))

    def _load_batch_cases(self):
        for item in self._batch_tree.get_children():
            self._batch_tree.delete(item)
        self._batch_rows.clear()
        self._batch_status.clear()
        children = self.queue_tree.get_children()
        cases = []
        for item in children:
            vals = self.queue_tree.item(item, "values")
            if vals:
                cid = _parse_tree_id(vals[0])
                country = vals[1] if len(vals) > 1 else "?"
                name = vals[2] if len(vals) > 2 else "?"
                cases.append((cid, name, country))
        self._batch_pending = cases

        for cid, name, country in cases:
            name_status, name_reason, fields_status, fields_reason = _load_batch_case_status(cid)
            item = self._batch_tree.insert("", "end", text=name, values=(country, name_status, fields_status))
            self._batch_rows[cid] = item
            self._batch_status[cid] = {
                "name": (name_status, name_reason),
                "fields": (fields_status, fields_reason)
            }

        self._batch_progress_var.set(f"共 {len(cases)} 个案例待核验")

    def _save_batch_checkpoint(self):
        """Save current batch state to a checkpoint file."""
        import json, os
        pending = []
        for cid in self._batch_rows:
            pending.append({
                "case_id": cid,
                "name": self._batch_tree.item(self._batch_rows[cid], "text"),
                "country": self._batch_tree.item(self._batch_rows[cid], "values")[0] if self._batch_tree.item(self._batch_rows[cid], "values") else "",
            })
        data = {
            "pending": pending,
            "name_model": self._batch_name_model.get(),
            "fields_model": self._batch_fields_model.get(),
        }
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "batch_checkpoint.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _restore_batch_checkpoint(self):
        """Restore batch state from checkpoint file if it exists."""
        import json, os, database as db
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "batch_checkpoint.json")
        if not os.path.exists(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return
        pending = data.get("pending", [])
        if not pending:
            return
        answer = messagebox.askyesno("恢复核验任务",
            f"上次有未完成的核验任务（{len(pending)} 个案例）。\n是否恢复？")
        if not answer:
            os.remove(path)
            return
        self._batch_name_model.set(data.get("name_model", "G"))
        self._batch_fields_model.set(data.get("fields_model", "DS"))
        for item in self._batch_tree.get_children():
            self._batch_tree.delete(item)
        self._batch_rows.clear()
        self._batch_status.clear()
        cases = []
        for p in pending:
            cid = p["case_id"]
            cases.append((cid, p["name"], p["country"]))
            name_status, name_reason, fields_status, fields_reason = _load_batch_case_status(cid)
            item = self._batch_tree.insert("", "end", text=p["name"], values=(p["country"], name_status, fields_status))
            self._batch_rows[cid] = item
            self._batch_status[cid] = {"name": (name_status, name_reason), "fields": (fields_status, fields_reason)}
        self._batch_pending = cases
        self._batch_progress_var.set(f"共 {len(cases)} 个案例待核验")

    def _on_batch_select(self, event):
        sel = self._batch_tree.selection()
        if not sel:
            return
        item = sel[0]
        for cid, tree_item in self._batch_rows.items():
            if tree_item == item:
                for q_item in self.queue_tree.get_children():
                    vals = self.queue_tree.item(q_item, "values")
                    if vals and _parse_tree_id(vals[0]) == cid:
                        self.queue_tree.selection_set(q_item)
                        self.queue_tree.see(q_item)
                        self._on_select(None)
                        break
                # Show latest verification results from batch status
                if cid in self._batch_status:
                    bs = self._batch_status[cid]
                    name_status, name_reason = bs.get("name", ("", ""))
                    fields_status, fields_reason = bs.get("fields", ("", ""))
                    if "成功" in name_status or "失败" in name_status:
                        self.name_result_text.delete("1.0", "end")
                        self.name_result_text.insert("1.0", f"名称核验: {name_status}\n{name_reason}")
                    if "成功" in fields_status or "失败" in fields_status:
                        self.result_text.delete("1.0", "end")
                        self.result_text.insert("1.0", f"字段核验: {fields_status}\n{fields_reason}")
                break

    def _set_step_status(self, case_id, step, status, reason):
        status_text = {"running": "⏳ 进行中", "success": "✅ 成功",
                       "failed": "❌ 失败", "conflict": "⚠️ 冲突"}.get(status, status)
        if case_id in self._batch_status:
            self._batch_status[case_id][step] = (status_text, reason)
        row = self._batch_rows.get(case_id)
        if row:
            col = "name_status" if step == "name" else "fields_status"
            self.frame.after(0, lambda r=row, c=col, s=status_text: self._batch_tree.set(r, c, s))

    def _start_batch(self):
        if self._batch_running:
            return
        self._batch_running = True
        self._batch_paused = False
        self._batch_stopped = False
        self._batch_start_btn.config(state="disabled")
        self._batch_pause_btn.config(state="normal")
        self._batch_stop_btn.config(state="normal")

        if self._batch_scope_var.get() == "选中案例":
            sel = self.queue_tree.selection()
            ids = set()
            for item in sel:
                vals = self.queue_tree.item(item, "values")
                if vals:
                    ids.add(_parse_tree_id(vals[0]))
            cases = [(c_id, n, c) for c_id, n, c in self._batch_pending if c_id in ids]
        else:
            cases = list(self._batch_pending)

        if not cases:
            messagebox.showinfo("提示", "没有待核验的案例。")
            self._on_batch_done()
            return

        self._batch_progress_bar["maximum"] = len(cases) * 2
        self._batch_progress_bar["value"] = 0

        self._save_batch_checkpoint()
        name_model = self._batch_name_model.get()
        fields_model = self._batch_fields_model.get()
        import threading
        t = threading.Thread(target=self._run_batch, args=(cases, name_model, fields_model), daemon=True)
        t.start()

    def _pause_batch(self):
        self._batch_paused = not self._batch_paused
        if self._batch_paused:
            self._batch_pause_btn.config(text="继续")
            self._batch_progress_var.set("已暂停")
            self._save_batch_checkpoint()
        else:
            self._batch_pause_btn.config(text="暂停")
            self._batch_progress_var.set("运行中...")

    def _stop_batch(self):
        self._batch_stopped = True
        self._save_batch_checkpoint()
        self._on_batch_done()

    def _on_batch_done(self):
        self._batch_running = False
        self._batch_start_btn.config(state="normal")
        self._batch_pause_btn.config(state="disabled")
        self._batch_stop_btn.config(state="disabled")
        if self._batch_paused:
            self._batch_pause_btn.config(text="暂停")

    def _run_batch(self, cases, name_model, fields_model):
        import database as db
        import time

        GEMINI_KEY = config.get_llm("field_completion")[2]
        DS_KEY = config.get_llm("search_parse")[2]

        for i, (case_id, case_name, country) in enumerate(cases):
            if self._batch_stopped:
                break
            while self._batch_paused and not self._batch_stopped:
                time.sleep(0.5)
            if self._batch_stopped:
                break

            self._batch_name_result = None  # Clear stale result from previous case

            case = db.get_verified_case_by_id(case_id)
            if not case:
                continue
            raw_id = case.get("raw_case_id", case_id)

            # Skip if already verified
            existing_logs = db.get_verification_logs(raw_id)
            name_done = any(l.get("action") == "name" and "correct:" in (l.get("result") or "") for l in existing_logs)
            fields_done = any(l.get("action") == "fields" and "filled:" in (l.get("result") or "") for l in existing_logs)
            if name_done and fields_done:
                self._set_step_status(case_id, "name", "success", "已有核验结果，跳过")
                self._set_step_status(case_id, "fields", "success", "已有核验结果，跳过")
                self.frame.after(0, lambda v=i*2+2: self._batch_progress_bar.configure({"value": v}))
                continue

            self._set_step_status(case_id, "name", "running", "")
            try:
                name_val = case.get("name", "") or ""
                if name_model == "G":
                    result = verify_field_gemini("name", "建筑名称", str(name_val), case, GEMINI_KEY)
                else:
                    result = verify_field_deepseek("name", "建筑名称", str(name_val), case, DS_KEY)
                if isinstance(result, dict):
                    correct = result.get("correct_value", name_val)
                    conf = result.get("confidence", "")
                    self._set_step_status(case_id, "name", "success",
                                          f"置信度:{conf} 正确值:{correct}")
                    db.insert_verification_log(raw_id, "name",
                                               f"current:{name_val}",
                                               json.dumps(result, ensure_ascii=False))
                    db.set_verified_case_status(case_id, "name_verified")
                    # Update name in DB if corrected
                    if correct and str(correct) != str(name_val):
                        db.update_verified_case(case_id, {"name": str(correct)})
                    # Store name result for later use in paste_text (Step 2)
                    self._batch_name_result = (result, name_model)
                    # Update detail panel and refresh queue
                    self._select_and_update_detail(case_id, result, step="name")
                else:
                    self._set_step_status(case_id, "name", "failed", str(result)[:200])
                    db.insert_verification_log(raw_id, "name", f"current:{name_val}", "parse_error")
            except Exception as e:
                self._set_step_status(case_id, "name", "failed", str(e)[:200])
                db.insert_verification_log(raw_id, "name", "exception", str(e)[:200])

            self.frame.after(0, lambda v=i*2+1: self._batch_progress_bar.configure({"value": v}))

            if self._batch_stopped:
                break
            while self._batch_paused and not self._batch_stopped:
                time.sleep(0.5)
            if self._batch_stopped:
                break

            self._set_step_status(case_id, "fields", "running", "")
            try:
                if fields_model == "G":
                    result = auto_complete_with_gemini(case, GEMINI_KEY)
                else:
                    result = auto_complete_fields(case, DS_KEY, GEMINI_KEY)
                if isinstance(result, dict) and not result.get("error"):
                    completions = result.get("completions", {})
                    filled = len(completions)
                    notes = result.get("notes", "")
                    if filled == 0 and ("失败" in notes or "解析" in notes):
                        self._set_step_status(case_id, "fields", "failed", notes[:200])
                        db.insert_verification_log(raw_id, "fields", "auto_complete", f"empty:{notes[:200]}")
                    else:
                        self._set_step_status(case_id, "fields", "success",
                                              f"补全{filled}个字段 {notes}")
                        db.insert_verification_log(raw_id, "fields",
                                                   "auto_complete", f"filled:{filled}")
                        db.set_verified_case_status(case_id, "verified")
                        # Put result into paste area and run info analysis (same as manual flow)
                        model_label = "G" if fields_model == "G" else "DS"
                        self.frame.after(0, lambda cid=case_id, r=result, ml=model_label: (
                            self._batch_paste_and_analyze(cid, r, ml)
                        ))
                else:
                    self._set_step_status(case_id, "fields", "failed", str(result)[:200])
                    db.insert_verification_log(raw_id, "fields", "auto_complete", "parse_error")
            except Exception as e:
                self._set_step_status(case_id, "fields", "failed", str(e)[:200])
                db.insert_verification_log(raw_id, "fields", "exception", str(e)[:200])

            self.frame.after(0, lambda v=i*2+2: self._batch_progress_bar.configure({"value": v}))

        self.frame.after(0, self._on_batch_done)
        self.frame.after(0, lambda: self._batch_progress_var.set(f"完成 {len(cases)} 个案例"))

    def _batch_paste_and_analyze(self, case_id, result, model_label):
        """Put batch fields result into paste area and run info analysis to apply."""
        # Select the case
        for item in self.queue_tree.get_children():
            vals = self.queue_tree.item(item, "values")
            if vals and _parse_tree_id(vals[0]) == case_id:
                self.queue_tree.selection_set(item)
                self.queue_tree.see(item)
                self._on_select(None)
                break
        if not self._current_case:
            return
        # Put results in paste text (name first, then fields — same order as manual)
        self.paste_text.configure(state="normal")
        self.paste_text.delete("1.0", "end")
        # Step 1 result: name verification
        name_result_data = getattr(self, '_batch_name_result', None)
        if name_result_data:
            name_result, name_model_label = name_result_data
            nm_label = "G" if name_model_label == "G" else "DS"
            self.paste_text.insert("end", f"=== 名称 {nm_label}核验结果 ===\n")
            self.paste_text.insert("end", json.dumps(name_result, ensure_ascii=False, indent=2))
            self._paste_sources.clear()
            self._paste_sources.add(f"{nm_label}核验")
        # Step 2 result: fields completion
        self.paste_text.insert("end", f"\n\n=== {model_label}补全结果 ===\n")
        self.paste_text.insert("end", json.dumps(result, ensure_ascii=False, indent=2))
        self.paste_text.configure(state="disabled")
        if name_result_data:
            self._paste_sources.add(f"{model_label}补全")
            self._last_paste_label = f"{model_label}补全"
        else:
            self._last_paste_label = f"{model_label}补全"
        # Run info analysis to extract and apply fields
        self._parse_and_fill_from_text()
        # Clean up stored name result
        self._batch_name_result = None

    def _open_batch_log(self):
        BatchLogDialog(self.frame, self)

    def _on_batch_right_click(self, event):
        sel = self._batch_tree.selection()
        if not sel:
            return
        menu = tk.Menu(self.frame, tearoff=0)
        menu.add_command(label="补跑此条", command=lambda: self._retry_batch_selected(sel))
        menu.post(event.x_root, event.y_root)

    def _retry_batch_selected(self, items=None):
        if items is None:
            items = self._batch_tree.selection()
        if not items:
            return
        to_retry = []
        for item in items:
            cid = None
            for k, v in self._batch_rows.items():
                if v == item:
                    cid = k
                    break
            if cid is None:
                continue
            vals = self._batch_tree.item(item, "values")
            if vals:
                if "失败" in str(vals[1]):
                    to_retry.append((cid, "name", item))
                if "失败" in str(vals[2]):
                    to_retry.append((cid, "fields", item))
        if not to_retry:
            messagebox.showinfo("提示", "请选中包含失败步骤的案例。")
            return
        import threading
        t = threading.Thread(target=self._run_batch_retry, args=(to_retry,), daemon=True)
        t.start()

    def _select_and_update_detail(self, case_id, result, step):
        """Select case in queue and write result to detail panel text boxes."""
        def _update():
            # Find and select the case in queue tree
            for item in self.queue_tree.get_children():
                vals = self.queue_tree.item(item, "values")
                if vals and _parse_tree_id(vals[0]) == case_id:
                    # Refresh queue tree item display after batch update
                    case = db.get_verified_case_by_id(case_id)
                    if case:
                        lang = self.lang_var.get()
                        self.queue_tree.item(item, values=(
                            _tree_id("V", case["id"]),
                            case.get("country", "") or "",
                            db.get_display_name(case, lang),
                        ))
                    self.queue_tree.selection_set(item)
                    self.queue_tree.see(item)
                    self._on_select(None)
                    break
            if self._current_case:
                if step == "name":
                    # Format name result into name_result_text
                    from datetime import datetime
                    text = f"核验时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    text += f"当前值: {result.get('current_value', '')}\n"
                    text += f"正确值: {result.get('correct_value', '')}\n"
                    text += f"中文名: {result.get('chinese_name', '')}\n"
                    text += f"英文名: {result.get('english_name', '')}\n"
                    text += f"官方语言名: {result.get('official_language_name', '')}\n"
                    text += f"常用名: {result.get('common_names', '')}\n"
                    text += f"置信度: {result.get('confidence', '')}\n"
                    text += f"结论: {result.get('conclusion', '')}\n"
                    text += f"推理: {result.get('reasoning', '')}\n"
                    refs = result.get('references', [])
                    if refs:
                        text += "参考来源:\n"
                        for r in refs:
                            text += f"  - {r}\n"
                    self.name_result_text.configure(state="normal")
                    self.name_result_text.delete("1.0", "end")
                    self.name_result_text.insert("1.0", text)
                elif step == "fields":
                    # Format fields result into result_text
                    text = ""
                    completions = result.get("completions", {})
                    for field, info in completions.items():
                        if isinstance(info, dict):
                            val = info.get("value", "")
                            conf = info.get("confidence", "")
                            text += f"[{field}] {val} (置信度: {conf})\n"
                        else:
                            text += f"[{field}] {info}\n"
                    notes = result.get("notes", "")
                    if notes:
                        text += f"\n备注: {notes}"
                    self.result_text.configure(state="normal")
                    self.result_text.delete("1.0", "end")
                    self.result_text.insert("1.0", text)
        self.frame.after(0, _update)

    def _run_batch_retry(self, to_retry):
        import database as db

        GEMINI_KEY = config.get_llm("field_completion")[2]
        DS_KEY = config.get_llm("search_parse")[2]

        for case_id, step, tree_item in to_retry:
            self._batch_name_result = None  # Clear stale result
            case = db.get_verified_case_by_id(case_id)
            if not case:
                continue
            raw_id = case.get("raw_case_id", case_id)

            col = "name_status" if step == "name" else "fields_status"
            self.frame.after(0, lambda _item=tree_item, _col=col: (
                self._batch_tree.set(_item, _col, "⏳ 进行中"),
            ))

            try:
                if step == "name":
                    name_val = case.get("name", "") or ""
                    name_model = self._batch_name_model.get()
                    if name_model == "G":
                        result = verify_field_gemini("name", "建筑名称", str(name_val), case, GEMINI_KEY)
                    else:
                        result = verify_field_deepseek("name", "建筑名称", str(name_val), case, DS_KEY)
                    if isinstance(result, dict):
                        correct = result.get("correct_value", name_val)
                        conf = result.get("confidence", "")
                        self._batch_name_result = (result, name_model)
                        self.frame.after(0, lambda _item=tree_item, _col=col, _c=conf, _v=correct: (
                            self._batch_tree.set(_item, _col, "✅ 成功"),
                        ))
                        db.insert_verification_log(raw_id, "name", f"retry:{name_val}",
                                                   json.dumps(result, ensure_ascii=False))
                    else:
                        self.frame.after(0, lambda _item=tree_item, _col=col, _r=str(result)[:200]: (
                            self._batch_tree.set(_item, _col, "❌ 失败"),
                        ))
                else:
                    fields_model = self._batch_fields_model.get()
                    if fields_model == "G":
                        result = auto_complete_with_gemini(case, GEMINI_KEY)
                    else:
                        result = auto_complete_fields(case, DS_KEY, GEMINI_KEY)
                    if isinstance(result, dict) and not result.get("error"):
                        completions = result.get("completions", {})
                        filled = len(completions)
                        notes = result.get("notes", "")
                        if filled == 0 and ("失败" in notes or "解析" in notes):
                            self.frame.after(0, lambda _item=tree_item, _col=col, _n=notes: (
                                self._batch_tree.set(_item, _col, "❌ 失败"),
                            ))
                        else:
                            self.frame.after(0, lambda _item=tree_item, _col=col, _f=filled, _n=notes: (
                                self._batch_tree.set(_item, _col, "✅ 成功"),
                            ))
                            ml = "G" if fields_model == "G" else "DS"
                            self.frame.after(0, lambda cid=case_id, r=result, mlabel=ml: (
                                self._batch_paste_and_analyze(cid, r, mlabel)
                            ))
                    else:
                        self.frame.after(0, lambda _item=tree_item, _col=col, _r=str(result)[:200]: (
                            self._batch_tree.set(_item, _col, "❌ 失败"),
                        ))
            except Exception as e:
                self.frame.after(0, lambda _item=tree_item, _col=col, _e=str(e)[:200]: (
                    self._batch_tree.set(_item, _col, "❌ 失败"),
                ))
                db.insert_verification_log(raw_id, step, "retry_exception", str(e)[:200])


class DatabasePanel:
    def __init__(self, parent, organizer):
        self.organizer = organizer
        self.frame = ttk.Frame(parent)
        self._api_key = tk.StringVar(value=config.get_llm("search_parse")[2])
        self._photo_cache = {}
        self._sort_reverse = False
        self._build_toolbar()
        self._build_filter_bar()
        self._build_tree()
        self.refresh()

    def _build_toolbar(self):
        bar = ttk.Frame(self.frame)
        bar.pack(fill="x", padx=5, pady=(5, 0))

        ttk.Button(bar, text="导出中文 Excel", command=self._export_data_cn).pack(side="left", padx=3)
        ttk.Button(bar, text="导出英文 Excel", command=self._export_data_en).pack(side="left", padx=3)
        ttk.Button(bar, text="新增案例", command=self._add_case).pack(side="left", padx=8)
        ttk.Button(bar, text="编辑选中", command=self._edit_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="删除选中", command=self._delete_selected).pack(side="left", padx=3)

        self.status_label = ttk.Label(bar, text="")
        self.status_label.pack(side="right", padx=10)

        bar2 = ttk.Frame(self.frame)
        bar2.pack(fill="x", padx=5, pady=(0, 5))

        ttk.Button(bar2, text="刷新列表", command=self.refresh).pack(side="left", padx=3)
        ttk.Button(bar2, text="查找重复", command=self._show_duplicates).pack(side="left", padx=3)
        ttk.Button(bar2, text="数据清洗", command=self._clean_data).pack(side="left", padx=3)
        ttk.Button(bar2, text="缺失检查", command=self._check_missing).pack(side="left", padx=3)

    def _build_filter_bar(self):
        bar = ttk.LabelFrame(self.frame, text=" 筛选 ", padding=5)
        bar.pack(fill="x", padx=5, pady=(2, 0))

        # Row 0: text search + clear button
        search_row = ttk.Frame(bar)
        search_row.pack(fill="x", pady=(0, 4))
        ttk.Label(search_row, text="搜索：").pack(side="left")
        self._search_var = tk.StringVar()
        search_entry = ttk.Entry(search_row, textvariable=self._search_var, width=35)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<Return>", lambda e: self.refresh())
        ttk.Button(search_row, text="搜索", command=self.refresh).pack(side="left", padx=2)
        ttk.Button(search_row, text="清除筛选", command=self._clear_filters).pack(side="left", padx=5)

        # Row 1: sort controls
        sort_row = ttk.Frame(bar)
        sort_row.pack(fill="x", pady=(0, 4))
        ttk.Label(sort_row, text="排序：").pack(side="left")

        sort_fields = [("updated_at", "默认（最近更新）"), ("name", "建筑名称"),
                       ("country", "国家"), ("city", "城市"),
                       ("year_built", "建造时间"), ("building_type", "功能"),
                       ("architect", "设计师"), ("condition", "建筑状态"),
                       ("spread_method", "传播方式")]

        self._sort_field_var = tk.StringVar(value="updated_at")
        sort_combo = ttk.Combobox(sort_row, textvariable=self._sort_field_var,
                                  values=[sf[1] for sf in sort_fields],
                                  state="readonly", width=18)
        sort_combo.pack(side="left", padx=5)
        sort_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self._sort_field_map = {sf[1]: sf[0] for sf in sort_fields}

        self._sort_desc_var = tk.BooleanVar(value=True)
        self._sort_dir_btn = ttk.Button(sort_row, text="↓ 降序", command=self._toggle_sort_dir)
        self._sort_dir_btn.pack(side="left", padx=3)

        # Row 2: filter dropdowns
        filter_row = ttk.Frame(bar)
        filter_row.pack(fill="x")

        filter_fields = [
            ("country", "国家"),
            ("building_type", "功能"),
            ("condition", "建筑状态"),
            ("spread_method", "传播方式"),
        ]

        self._filter_vars = {}
        self._filter_combos = {}

        for i, (field, label) in enumerate(filter_fields):
            pad_left = (0, 0) if i == 0 else (8, 0)
            ttk.Label(filter_row, text=f"{label}：").pack(side="left", padx=pad_left)
            var = tk.StringVar(value="全部")
            combo = ttk.Combobox(filter_row, textvariable=var, values=["全部"],
                                 state="readonly", width=14)
            combo.pack(side="left", padx=3)
            combo.bind("<<ComboboxSelected>>", lambda e, f=field: self.refresh())
            self._filter_vars[field] = var
            self._filter_combos[field] = combo

        self._load_filter_values()

    def _load_filter_values(self):
        for field, combo in self._filter_combos.items():
            current = self._filter_vars[field].get()
            values = db.get_distinct_values("cases", field)
            combo["values"] = ["全部"] + values
            if current == "全部" or current in values:
                self._filter_vars[field].set(current)
            else:
                self._filter_vars[field].set("全部")

    def _toggle_sort_dir(self):
        if self._sort_desc_var.get():
            self._sort_desc_var.set(False)
            self._sort_dir_btn.config(text="↑ 升序")
        else:
            self._sort_desc_var.set(True)
            self._sort_dir_btn.config(text="↓ 降序")
        self.refresh()

    def _clear_filters(self):
        self._search_var.set("")
        for var in self._filter_vars.values():
            var.set("全部")
        self._sort_field_var.set("默认（最近更新）")
        self._sort_desc_var.set(True)
        self._sort_dir_btn.config(text="↓ 降序")
        self.refresh()

    def _build_tree(self):
        tree_frame = ttk.Frame(self.frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=2)

        columns = ("id", "name", "country", "city", "year_built",
                   "demolition_year", "restoration_year", "quantity",
                   "building_type", "architect", "condition", "spread_method")

        # Custom style for taller rows to fit thumbnail images
        style = ttk.Style()
        style.configure("Thumbnail.Treeview", rowheight=85)

        self.tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings",
                                 selectmode="extended", style="Thumbnail.Treeview")

        self.tree.heading("#0", text="图片")
        self.tree.column("#0", width=100, minwidth=80)

        headers = {"id": "ID", "name": "建筑名称", "country": "国家", "city": "城市",
                   "year_built": "建造时间", "demolition_year": "拆除时间",
                   "restoration_year": "修复时间", "quantity": "数量",
                   "building_type": "功能", "architect": "设计师",
                   "condition": "建筑状态", "spread_method": "传播方式"}
        widths = {"id": 40, "name": 160, "country": 80, "city": 80,
                  "year_built": 70, "demolition_year": 70,
                  "restoration_year": 70, "quantity": 50,
                  "building_type": 70, "architect": 80,
                  "condition": 65, "spread_method": 80}

        for col in columns:
            self.tree.heading(col, text=headers[col],
                              command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=widths.get(col, 80))

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self._on_db_tree_click)

        _make_tree_draggable(self.tree, "cases")

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._photo_cache.clear()

        # Build filters from dropdown selections
        filters = {}
        search = None
        if hasattr(self, '_filter_vars'):
            for field, var in self._filter_vars.items():
                val = var.get()
                if val and val != "全部":
                    filters[field] = val
        if hasattr(self, '_search_var'):
            search = self._search_var.get().strip() or None

        # Build sort selection
        order_by = None
        sort_field = None
        sort_desc = True
        if hasattr(self, '_sort_field_var'):
            sort_label = self._sort_field_var.get()
            sort_field = self._sort_field_map.get(sort_label)
            if sort_field and sort_field != "updated_at":
                sort_desc = self._sort_desc_var.get()
                order_by = (sort_field, "DESC" if sort_desc else "ASC")

        cases = db.get_all_cases(search=search, filters=filters if filters else None,
                                  order_by=order_by)

        # Sort by country case count (countries with more cases first)
        if sort_field == "country":
            country_counts = {}
            for c in cases:
                co = (c.get("country") or "").strip()
                if co:
                    country_counts[co] = country_counts.get(co, 0) + 1
            cases.sort(key=lambda c: (
                -country_counts.get((c.get("country") or "").strip(), 0) * (1 if sort_desc else -1),
                0 if (c.get("name") or "").strip() else 1,
            ))
        for c in cases:
            _img_urls, _merged_path, urls = merged_image_sources(c)
            count_text = f"[{len(urls)}张]" if urls else "—"

            item_id = self.tree.insert("", "end", text=count_text, values=(
                _tree_id("Z", c["id"]), c.get("name", ""), c.get("country", "") or "",
                c.get("city", "") or "", c.get("year_built", "") or "",
                c.get("demolition_year", "") or "", c.get("restoration_year", "") or "",
                c.get("quantity", "") or "", c.get("building_type", "") or "",
                c.get("architect", "") or "",
                c.get("condition") or "",
                c.get("spread_method") or "",
            ))

            if urls:
                thumb_url = urls[0]
                for u in urls:
                    if u.startswith("dbimg://"):
                        thumb_url = u
                        break
                threading.Thread(
                    target=load_tree_thumbnail,
                    args=(self.tree, item_id, thumb_url, c["id"], self._photo_cache),
                    daemon=True
                ).start()

        self.status_label.config(text=f"共 {len(cases)} 条记录")

        if hasattr(self, '_filter_combos'):
            self._load_filter_values()

    def _sort_by(self, col):
        """Sort tree contents by the given column, toggling asc/desc."""
        rows = [(self.tree.set(item, col), item) for item in self.tree.get_children("")]
        reverse = getattr(self, "_sort_reverse", False)
        # Separate empty and non-empty; empty always at end regardless of direction
        non_empty = [(v, i) for v, i in rows if v.strip()]
        empty = [(v, i) for v, i in rows if not v.strip()]
        def _key(x):
            v = x[0].strip()
            try:
                return float(v) if v.replace(".", "").isdigit() else v
            except (ValueError, AttributeError, TypeError):
                return v
        non_empty.sort(key=_key, reverse=reverse)
        sorted_rows = non_empty + empty
        for i, (_, item) in enumerate(sorted_rows):
            self.tree.move(item, "", i)
        self._sort_reverse = not reverse

    # -- Existing functionality (unchanged from before) --

    # --- Translation maps for English export ---
    _CN_EN_COUNTRY = {
        "印度": "India", "印度尼西亚": "Indonesia", "蒙古": "Mongolia", "越南": "Vietnam",
        "日本": "Japan", "韩国": "South Korea", "朝鲜": "North Korea",
        "泰国": "Thailand", "缅甸": "Myanmar", "柬埔寨": "Cambodia", "老挝": "Laos",
        "马来西亚": "Malaysia", "新加坡": "Singapore", "菲律宾": "Philippines",
        "文莱": "Brunei", "东帝汶": "East Timor",
        "尼泊尔": "Nepal", "不丹": "Bhutan", "孟加拉国": "Bangladesh", "斯里兰卡": "Sri Lanka",
        "巴基斯坦": "Pakistan", "阿富汗": "Afghanistan",
        "哈萨克斯坦": "Kazakhstan", "乌兹别克斯坦": "Uzbekistan", "吉尔吉斯斯坦": "Kyrgyzstan",
        "塔吉克斯坦": "Tajikistan", "土库曼斯坦": "Turkmenistan",
        "伊朗": "Iran", "伊拉克": "Iraq", "叙利亚": "Syria", "约旦": "Jordan",
        "黎巴嫩": "Lebanon", "以色列": "Israel", "巴勒斯坦": "Palestine",
        "沙特阿拉伯": "Saudi Arabia", "阿联酋": "UAE", "卡塔尔": "Qatar",
        "巴林": "Bahrain", "科威特": "Kuwait", "阿曼": "Oman", "也门": "Yemen",
        "土耳其": "Turkey", "俄罗斯": "Russia", "白俄罗斯": "Belarus", "乌克兰": "Ukraine",
        "波兰": "Poland", "捷克": "Czech Republic", "斯洛伐克": "Slovakia",
        "匈牙利": "Hungary", "罗马尼亚": "Romania", "保加利亚": "Bulgaria",
        "塞尔维亚": "Serbia", "克罗地亚": "Croatia", "斯洛文尼亚": "Slovenia",
        "德国": "Germany", "法国": "France", "英国": "United Kingdom", "意大利": "Italy",
        "西班牙": "Spain", "葡萄牙": "Portugal", "荷兰": "Netherlands", "比利时": "Belgium",
        "瑞士": "Switzerland", "奥地利": "Austria", "瑞典": "Sweden", "挪威": "Norway",
        "丹麦": "Denmark", "芬兰": "Finland", "爱尔兰": "Ireland", "希腊": "Greece",
        "美国": "United States", "加拿大": "Canada", "墨西哥": "Mexico",
        "巴西": "Brazil", "阿根廷": "Argentina", "智利": "Chile", "秘鲁": "Peru",
        "哥伦比亚": "Colombia", "委内瑞拉": "Venezuela", "古巴": "Cuba",
        "澳大利亚": "Australia", "新西兰": "New Zealand",
        "埃及": "Egypt", "南非": "South Africa", "肯尼亚": "Kenya", "尼日利亚": "Nigeria",
        "摩洛哥": "Morocco", "阿尔及利亚": "Algeria", "突尼斯": "Tunisia",
        "埃塞俄比亚": "Ethiopia", "坦桑尼亚": "Tanzania", "加纳": "Ghana",
    }
    _CN_EN_TYPE = {
        "园林": "Garden", "民居": "Residence", "寺庙": "Temple", "宫殿": "Palace",
        "衙署": "Government Office", "塔": "Pagoda", "桥": "Bridge",
        "祠堂": "Ancestral Hall", "书院": "Academy", "现代中式": "Modern Chinese Style",
        "其他": "Other", "会馆": "Guild Hall", "庙堂": "Temple Hall",
        "博物馆": "Museum", "宗教朝拜场所": "Religious Site",
        "旅游景点": "Tourist Attraction", "中餐厅": "Chinese Restaurant",
        "同乡会馆": "Native-Place Guild Hall", "清真寺": "Mosque",
        "图书馆": "Library", "医院": "Hospital",
        "政府行政机关": "Government Office",
    }
    _CN_EN_CONDITION = {
        "良好": "Good", "破损": "Damaged", "倒塌": "Collapsed", "已毁": "Destroyed", "状态良好": "Good Condition",
        "破损待修复": "Damaged, Needs Repair", "损毁严重": "Severely Damaged", "拆除": "Demolished",
        "未知": "Unknown",
        "现存": "Existing", "历史遗迹": "Historic Site", "幼儿园": "Kindergarten",
        "已衰败": "Declined", "废弃": "Abandoned", "空置": "Vacant",
        "状态": "Condition",
    }
    _CN_EN_SPREAD = {
        "华人": "Chinese Diaspora", "传教士": "Missionary",
        "政治决策": "Political Decision", "商业贸易": "Trade",
        "文化交流": "Cultural Exchange", "殖民建设": "Colonial Construction",
        "藏传佛教传播": "Tibetan Buddhist Spread", "其他": "Other",
        "华人移民": "Chinese Diaspora", "移民": "Immigrant",
    }

    _CN_EN_ARCHITECT = {
        "当地华人工匠": "Local Chinese Craftsmen",
        "当地华人": "Local Chinese",
        "华人工匠": "Chinese Craftsmen",
        "工匠": "Craftsmen",
        "华人": "Chinese",
    }

    _CN_EN_GENERAL = {
        "纪念郑和的文化遗产": "Zheng He Cultural Heritage",
        "首都": "Capital", "包括": "including", "等": "etc.",
        "但": "but", "曾设": "formerly housed",
        "当地": "local", "工匠": "craftsmen",
        "广肇": "Guangzhao", "阿格拉": "Agra",
    }

    _CN_EN_CITY = {
        "加尔各答": "Kolkata", "河内": "Hanoi", "胡志明市": "Ho Chi Minh City",
        "三宝垄": "Semarang", "摩拉路": "Jalan Muara", "额尔登特": "Erdenet",
        "会安": "Hoi An",
    }

    # Reverse maps: EN → CN for Chinese export
    _EN_CN_COUNTRY = {v: k for k, v in _CN_EN_COUNTRY.items()}
    _EN_CN_TYPE = {v: k for k, v in _CN_EN_TYPE.items()}
    _EN_CN_CONDITION = {v: k for k, v in _CN_EN_CONDITION.items()}
    _EN_CN_SPREAD = {v: k for k, v in _CN_EN_SPREAD.items()}
    _EN_CN_CITY = {v: k for k, v in _CN_EN_CITY.items()}
    _EN_CN_GENERAL = {v: k for k, v in _CN_EN_GENERAL.items()}
    _EN_CN_ARCHITECT = {v: k for k, v in _CN_EN_ARCHITECT.items()}

    @staticmethod
    def _translate_text(text, to_en):
        """Translate known terms in text using all available maps."""
        if not text or not isinstance(text, str):
            return text or ""
        if to_en:
            maps = [DatabasePanel._CN_EN_COUNTRY, DatabasePanel._CN_EN_TYPE,
                    DatabasePanel._CN_EN_CONDITION, DatabasePanel._CN_EN_SPREAD,
                    DatabasePanel._CN_EN_CITY, DatabasePanel._CN_EN_GENERAL,
                    DatabasePanel._CN_EN_ARCHITECT]
        else:
            maps = [DatabasePanel._EN_CN_COUNTRY, DatabasePanel._EN_CN_TYPE,
                    DatabasePanel._EN_CN_CONDITION, DatabasePanel._EN_CN_SPREAD,
                    DatabasePanel._EN_CN_CITY, DatabasePanel._EN_CN_GENERAL,
                    DatabasePanel._EN_CN_ARCHITECT]
        # Collect all pairs, sort by key length descending to avoid partial replacements
        pairs = []
        for m in maps:
            for k, v in m.items():
                pairs.append((k, v))
        pairs.sort(key=lambda x: len(x[0]), reverse=True)
        result = text
        for k, v in pairs:
            if k in result:
                result = result.replace(k, v)
            elif k.lower() in result.lower():
                import re as _re
                result = _re.sub(_re.escape(k), v, result, flags=_re.IGNORECASE)
        return result

    def _translate_val(self, val, mapping):
        """Translate a string value using mapping; if multi-value comma-separated,
        translate each part. Return original if no mapping found."""
        if not val or not isinstance(val, str):
            return val or ""
        parts = [p.strip() for p in val.replace("、", ",").replace("，", ",").split(",")]
        translated = []
        for p in parts:
            # Try exact match first (case-insensitive)
            found_en = None
            for k, v in mapping.items():
                if p.lower() == k.lower():
                    found_en = v
                    break
            if found_en:
                translated.append(found_en)
            else:
                # Try substring match (case-insensitive), replace only matched part
                found = False
                for k, v in mapping.items():
                    if k.lower() in p.lower():
                        translated.append(re.sub(re.escape(k), v, p, flags=re.IGNORECASE))
                        found = True
                        break
                if not found:
                    translated.append(p)
        return ", ".join(translated)

    def _export_data_cn(self):
        self._do_export("cn")

    def _export_data_en(self):
        self._do_export("en")

    def _do_export(self, lang):
        import io
        from openpyxl import Workbook
        from openpyxl.styles import Alignment
        from openpyxl.utils import get_column_letter
        from openpyxl.drawing.image import Image as XLImage
        from PIL import Image as PILImage
        from ui.image_components import fetch_image_bytes

        # Use same sort order as the current UI list
        is_cn = (lang == "cn")
        order_by = None
        sort_field = None
        sort_desc = True
        if hasattr(self, '_sort_field_var'):
            sort_label = self._sort_field_var.get()
            sort_field = self._sort_field_map.get(sort_label)
            if sort_field and sort_field != "updated_at":
                sort_desc = self._sort_desc_var.get()
                order_by = (sort_field, "DESC" if sort_desc else "ASC")

        cases = db.get_all_cases(order_by=order_by)

        if sort_field == "country":
            country_counts = {}
            for c in cases:
                co = (c.get("country") or "").strip()
                if co:
                    country_counts[co] = country_counts.get(co, 0) + 1
            cases.sort(key=lambda c: (
                -country_counts.get((c.get("country") or "").strip(), 0) * (1 if sort_desc else -1),
                0 if (c.get("name") or "").strip() else 1,
            ))

        if not cases:
            messagebox.showinfo("提示", "没有数据可导出。")
            return

        # --- Column definitions ---
        db_cols = ["id", "name", "name_en", "name_official", "country", "city",
                    "year_built", "demolition_year", "restoration_year", "quantity",
                    "building_type", "architect", "condition", "spread_method",
                    "latitude", "longitude"]
        cn_headers = ["ID", "建筑名称", "英文名称", "官方语言名称", "国家", "城市",
                      "建造时间", "拆除时间", "修复时间", "数量",
                      "功能", "设计师", "建筑状态", "传播方式",
                      "纬度", "经度"]
        en_headers = ["ID", "Name", "English Name", "Official Name", "Country", "City",
                      "Year Built", "Demolition Year", "Restoration Year", "Quantity",
                      "Building Type", "Architect", "Condition", "Spread Method",
                      "Latitude", "Longitude"]

        cn_extra = ["名称核验结果", "其他字段核验结果", "参考链接"]
        en_extra = ["Name Verification", "Field Verification", "Reference Links"]
        max_img_cols = 5
        for i in range(1, max_img_cols + 1):
            cn_extra.append(f"图片{i}")
            en_extra.append(f"Image {i}")

        headers = cn_headers if is_cn else en_headers
        extra_headers = cn_extra if is_cn else en_extra
        all_headers = headers + extra_headers
        BASE_COLS = len(db_cols)
        VERIFY_COLS = 3  # 名称核验结果, 字段分析结果, 参考链接

        # --- Extract verification details + images + name fields for each case ---
        name_results_list = []
        field_verify_list = []
        ref_links_list = []
        image_rows = []
        vfy_names_cn = []
        vfy_names_en = []
        vfy_names_off = []
        for case in cases:
            # Get name fields from the same function used by the verification details dialog
            name_info = _build_name_info(case)
            vfy_names_cn.append(name_info["chinese_name"] or "")
            vfy_names_en.append(name_info["english_name"] or "")
            vfy_names_off.append(name_info["official_language_name"] or "")

            # Get verification display data from the same logic as the dialog
            name_text, field_text, ref_text = _build_verification_display_data(case, lang)
            name_results_list.append(name_text)
            field_verify_list.append(field_text)
            ref_links_list.append(ref_text)

            # Collect images from the curated image_path / image_urls fields
            row_images = []
            img_sources = parse_image_urls(
                case.get("image_urls", "") or "",
                case.get("image_path", "") or "")
            for src in img_sources:
                try:
                    data, _ = fetch_image_bytes(src, timeout=10)
                    if data:
                        row_images.append(PILImage.open(io.BytesIO(data)))
                except Exception:
                    pass
            image_rows.append(row_images)

        # --- Build Excel with openpyxl ---
        wb = Workbook()
        ws = wb.active
        ws.title = "中文导出" if is_cn else "English Export"

        # Headers
        for col_idx, h in enumerate(all_headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Data rows
        IMG_START_COL = BASE_COLS + VERIFY_COLS + 1  # first image column (1-based)
        IMG_MAX_W, IMG_MAX_H = 120, 90

        for row_idx, case in enumerate(cases):
            excel_row = row_idx + 2
            for col_idx, db_col in enumerate(db_cols):
                val = case.get(db_col, "")
                if val is None:
                    val = ""
                # Use verification-extracted name fields for name columns
                if db_col == "name":
                    if is_cn:
                        val = vfy_names_cn[row_idx] or val or ""
                    else:
                        val = vfy_names_en[row_idx] or vfy_names_cn[row_idx] or val or ""
                elif db_col == "name_en":
                    val = vfy_names_en[row_idx] or val or ""
                elif db_col == "name_official":
                    val = vfy_names_off[row_idx] or val or ""
                if val is None:
                    val = ""
                # Translate for English export
                if not is_cn and isinstance(val, str) and val:
                    if db_col == "country":
                        val = self._translate_val(val, self._CN_EN_COUNTRY)
                    elif db_col == "city":
                        val = self._translate_val(val, self._CN_EN_CITY)
                    elif db_col == "building_type":
                        val = self._translate_val(val, self._CN_EN_TYPE)
                    elif db_col == "condition":
                        val = self._translate_val(val, self._CN_EN_CONDITION)
                    elif db_col == "spread_method":
                        val = self._translate_val(val, self._CN_EN_SPREAD)
                    elif db_col == "architect":
                        val = self._translate_val(val, self._CN_EN_ARCHITECT)
                # Translate for Chinese export (EN → CN)
                if is_cn and isinstance(val, str) and val:
                    if db_col == "country":
                        val = self._translate_val(val, self._EN_CN_COUNTRY)
                    elif db_col == "city":
                        val = self._translate_val(val, self._EN_CN_CITY)
                    elif db_col == "building_type":
                        val = self._translate_val(val, self._EN_CN_TYPE)
                    elif db_col == "condition":
                        val = self._translate_val(val, self._EN_CN_CONDITION)
                    elif db_col == "spread_method":
                        val = self._translate_val(val, self._EN_CN_SPREAD)
                    elif db_col == "architect":
                        val = self._translate_val(val, self._EN_CN_ARCHITECT)
                ws.cell(row=excel_row, column=col_idx + 1, value=str(val))

            # Verification columns
            ws.cell(row=excel_row, column=BASE_COLS + 1, value=name_results_list[row_idx])
            ws.cell(row=excel_row, column=BASE_COLS + 2, value=field_verify_list[row_idx])
            ws.cell(row=excel_row, column=BASE_COLS + 3, value=ref_links_list[row_idx])

            # Image columns: one image per column, up to max_img_cols
            imgs = image_rows[row_idx]
            if imgs:
                ws.row_dimensions[excel_row].height = IMG_MAX_H * 0.8
            for i, pil_img in enumerate(imgs[:max_img_cols]):
                try:
                    thumb = pil_img.copy()
                    thumb.thumbnail((IMG_MAX_W, IMG_MAX_H), PILImage.LANCZOS)
                    buf = io.BytesIO()
                    thumb.save(buf, format="PNG")
                    buf.seek(0)
                    xl_img = XLImage(buf)
                    xl_img.width = thumb.width
                    xl_img.height = thumb.height
                    img_col = IMG_START_COL + i
                    xl_img.anchor = f"{get_column_letter(img_col)}{excel_row}"
                    ws.add_image(xl_img)
                except Exception:
                    pass

        # Column widths
        widths = {
            1: 6, 2: 22, 3: 18, 4: 18, 5: 10, 6: 10,
            7: 10, 8: 10, 9: 10, 10: 6, 11: 14, 12: 12, 13: 8, 14: 10,
            BASE_COLS + 1: 32,
            BASE_COLS + 2: 35,
            BASE_COLS + 3: 28,
        }
        for c, w in widths.items():
            ws.column_dimensions[get_column_letter(c)].width = w
        for i in range(max_img_cols):
            ws.column_dimensions[get_column_letter(IMG_START_COL + i)].width = IMG_MAX_W * 1.2 / 7

        # Freeze header row
        ws.freeze_panes = "A2"

        # --- Save ---
        title = "导出中文数据" if is_cn else "Export English Data"
        default_name = "export_cn.xlsx" if is_cn else "export_en.xlsx"
        filepath = filedialog.asksaveasfilename(
            title=title,
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not filepath:
            return
        try:
            wb.save(filepath)
            messagebox.showinfo("完成", f"数据已导出到:\n{filepath}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _show_duplicates(self):
        import sqlite3
        conn = sqlite3.connect(db.DB_PATH)
        rows = conn.execute("""
            SELECT name, country, COUNT(*) as cnt, GROUP_CONCAT(id) as ids
            FROM cases
            GROUP BY name, country
            HAVING cnt > 1
        """).fetchall()
        conn.close()
        dupes = [{"name": r[0], "country": r[1], "count": r[2],
                  "ids": [int(x) for x in r[3].split(",")]} for r in rows]

        if not dupes:
            messagebox.showinfo("去重检查", "未发现重复记录。")
            return

        win = tk.Toplevel(self.frame)
        win.title(f"发现 {len(dupes)} 组重复记录")
        win.geometry("700x400")
        win.transient(self.frame)

        ttk.Label(win, text="以下记录按「建筑名称+国家」判断为可能重复，请手动确认并删除：",
                  font=("", 10)).pack(padx=10, pady=5)

        tree = ttk.Treeview(win, columns=("name", "country", "count"), show="headings")
        tree.heading("name", text="建筑名称")
        tree.heading("country", text="国家")
        tree.heading("count", text="重复数量")
        tree.column("name", width=250)
        tree.column("country", width=150)
        tree.column("count", width=80)
        tree.pack(fill="both", expand=True, padx=10, pady=5)

        for d in dupes:
            tree.insert("", "end", values=(d["name"], d.get("country", ""), d["count"]))

        def delete_dupes():
            sel = tree.selection()
            if not sel:
                return
            for item in sel:
                idx = tree.index(item)
                ids = dupes[idx]["ids"]
                for dup_id in ids[1:]:
                    db.delete_case(dup_id)
            messagebox.showinfo("完成", "已处理选中的重复记录（保留每组第一条）。")
            win.destroy()
            self.refresh()

        ttk.Button(win, text="删除选中重复（保留第一条）", command=delete_dupes).pack(pady=8)

    def _clean_data(self):
        if not messagebox.askyesno("数据清洗", "将执行以下操作：\n- 去除所有文本字段首尾空白\n- 统一空值为空字符串\n\n确定继续？"):
            return

        cases = db.get_all_cases()
        cleaned = 0
        for c in cases:
            updates = {}
            for field, label, *_ in db.FIELDS:
                val = c.get(field)
                if isinstance(val, str) and val.strip() != val:
                    updates[field] = val.strip()
            if updates:
                db.update_case(c["id"], updates)
                cleaned += 1

        messagebox.showinfo("清洗完成", f"已清洗 {cleaned} 条记录。")
        self.refresh()

    def _check_missing(self):
        cases = db.get_all_cases()
        missing = []
        key_fields = ["name", "country", "city", "year_built", "building_type", "architect"]
        key_labels = {"name": "建筑名称", "country": "国家", "city": "城市",
                     "year_built": "建造时间", "building_type": "功能", "architect": "设计师"}
        for c in cases:
            for f in key_fields:
                if not c.get(f):
                    missing.append({"id": c["id"], "name": c.get("name", ""),
                                   "missing_label": key_labels[f]})
        if not missing:
            messagebox.showinfo("缺失检查", "所有记录的关键字段都已填写完整。")
            return

        win = tk.Toplevel(self.frame)
        win.title(f"缺失信息 - {len(missing)} 处")
        win.geometry("600x400")
        win.transient(self.frame)

        tree = ttk.Treeview(win, columns=("id", "name", "field"), show="headings")
        tree.heading("id", text="ID")
        tree.heading("name", text="建筑名称")
        tree.heading("field", text="缺失字段")
        tree.column("id", width=50)
        tree.column("name", width=250)
        tree.column("field", width=150)
        tree.pack(fill="both", expand=True, padx=10, pady=5)

        for m in missing:
            tree.insert("", "end", values=(m["id"], m["name"], m["missing_label"]))

        def goto_case():
            sel = tree.selection()
            if not sel:
                return
            item = tree.item(sel[0])
            case_id = int(item["values"][0])
            case = db.get_case_by_id(case_id)
            if case:
                dlg = _EditDialog(win, case)
                win.wait_window(dlg.win)
                win.destroy()
                self.refresh()

        ttk.Button(win, text="编辑选中案例", command=goto_case).pack(pady=8)

    def _add_case(self):
        empty_case = {"id": None, "name": "", "country": "", "city": "",
                      "year_built": "", "demolition_year": "",
                      "restoration_year": "", "quantity": "",
                      "building_type": "", "architect": "", "condition": "",
                      "latitude": "", "longitude": "", "image_path": "", "image_urls": "",
                      "search_time": "", "search_keywords": "", "notes": "",
                      "name_en": "", "name_official": "",
                      "material": "", "area_sqm": "", "protection_level": "",
                      "description": ""}
        dlg = _EditDialog(self.frame, empty_case)
        self.frame.wait_window(dlg.win)
        self.refresh()

    def _on_db_tree_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        case_id = _parse_tree_id(self.tree.item(sel[0], "values")[0])
        case = db.get_case_by_id(case_id)
        if case:
            _VerificationResultDialog(self.frame, case, is_formal=True)

    def _edit_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择一条记录。")
            return
        case_id = _parse_tree_id(self.tree.item(sel[0], "values")[0])
        case = db.get_case_by_id(case_id)
        if case:
            dlg = _EditDialog(self.frame, case)
            self.frame.wait_window(dlg.win)
            self.refresh()

    def _delete_selected(self):
        items = self.tree.selection()
        if not items:
            messagebox.showinfo("提示", "请先选择要删除的记录。")
            return
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(items)} 条记录吗？"):
            for item in items:
                case_id = self.tree.item(item, "values")[0]
                db.delete_case(_parse_tree_id(case_id))
            self.refresh()

    # -- AI completion --


# ============================================================
#  Add Raw Case Dialog
# ============================================================

_ADD_FIELDS = [
    ("name", "建筑名称"), ("name_en", "英文名称"), ("name_official", "官方语言名称"),
    ("country", "国家"), ("city", "城市"),
    ("construction_year", "建造时间"), ("function", "功能"),
    ("designer", "设计师"), ("spread_method", "传播方式"),
    ("building_status", "建筑状态"), ("quantity", "数量规模"),
    ("search_keywords", "检索词"), ("authenticity_notes", "备注"),
]


class _AddRawCaseDialog:
    def __init__(self, parent, on_done=None):
        self.on_done = on_done
        self.win = tk.Toplevel(parent)
        self.win.title("添加案例")
        self.win.geometry("500x550")
        self.win.transient(parent)

        canvas = tk.Canvas(self.win, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", width=480)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y", pady=10)

        self._entries = {}
        for field, label in _ADD_FIELDS:
            row = ttk.Frame(inner)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=label, width=12).pack(side="left")
            var = tk.StringVar()
            e = ttk.Entry(row, textvariable=var, width=40)
            e.pack(side="left", padx=5)
            self._entries[field] = var

        btn_row = ttk.Frame(inner)
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="保存", command=self._save).pack(side="right", padx=3)
        ttk.Button(btn_row, text="取消", command=self.win.destroy).pack(side="right", padx=3)

    def _save(self):
        data = {}
        for field, var in self._entries.items():
            val = var.get().strip()
            if val:
                data[field] = val
        if not data.get("name") and not data.get("country"):
            messagebox.showwarning("提示", "建筑名称和国家至少填一个。", parent=self.win)
            return
        if not data.get("name"):
            data["name"] = ""
            if "quantity" not in data:
                # Auto-count: use existing case count for this country, or 0
                country = data.get("country", "")
                data["quantity"] = db.count_raw_cases_by_country(country)
        db.insert_raw_case(data)
        self.win.destroy()
        if self.on_done:
            self.on_done()


# ============================================================
#  Merge Dialog
# ============================================================

_MERGE_FIELDS = [
    ("name", "建筑名称"), ("name_en", "英文名称"), ("name_official", "官方语言名称"),
    ("country", "国家"), ("city", "城市"),
    ("construction_year", "建造时间"), ("demolition_year", "拆除时间"),
    ("restoration_year", "修复时间"), ("quantity", "数量规模"),
    ("function", "功能"), ("designer", "设计师"),
    ("latitude", "纬度"), ("longitude", "经度"),
    ("search_time", "检索时间"), ("search_keywords", "检索词"),
    ("spread_method", "传播方式"), ("building_status", "建筑状态"),
    ("image_urls", "图片链接"), ("source_urls", "来源链接"),
    ("authenticity_notes", "备注"),
]

_SKIP_MERGE_FIELDS = {"id", "created_at", "updated_at", "authenticity_status", "authenticity_score", "raw_case_id"}


def _build_merge_fields(cases):
    """Build full field list: predefined fields + any custom fields present in any case."""
    seen = set()
    result = []
    for f, l in _MERGE_FIELDS:
        result.append((f, l))
        seen.add(f)
    for case in cases:
        for key in case:
            if key not in seen and key not in _SKIP_MERGE_FIELDS:
                result.append((key, key))
                seen.add(key)
    return result


class _MergeDialog:
    def __init__(self, parent, cases, on_done=None, smart_defaults=False, table="raw"):
        self.cases = cases
        self.on_done = on_done
        self._smart_defaults = smart_defaults
        self._table = table

        n = len(cases)
        win_width = min(300 + n * 210, 1500)

        self.win = tk.Toplevel(parent)
        self.win.title(f"合并案例 ({n}条记录)")
        self.win.geometry(f"{win_width}x700")
        self.win.transient(parent)

        ttk.Label(self.win, text="勾选每个字段要保留的值（可多选），保留的值将用 \"; \" 拼接。第一条为主记录。",
                  font=("", 10)).pack(padx=15, pady=(15, 5))

        # Header
        hdr = ttk.Frame(self.win)
        hdr.pack(fill="x", padx=15, pady=5)
        ttk.Label(hdr, text="字段", width=14, font=("", 9, "bold")).pack(side="left")

        self._hdr_labels = []
        for i, case in enumerate(cases):
            name = case.get("name", "?")[:22]
            prefix = "★主" if i == 0 else f"#{i + 1}"
            color = "#1565c0" if i == 0 else "#555"
            lbl = ttk.Label(hdr, text=f"{prefix} {name}", width=26, font=("", 9, "bold"),
                            foreground=color)
            lbl.pack(side="left", padx=(0, 3))
            self._hdr_labels.append(lbl)

        ttk.Label(hdr, text="状态", width=6, font=("", 9, "bold")).pack(side="left")

        # Scrollable field list
        scroll_frame = ttk.Frame(self.win)
        scroll_frame.pack(fill="both", expand=True, padx=15, pady=5)

        canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        self._field_frame = ttk.Frame(canvas)

        def _on_canvas_configure(event):
            canvas.itemconfig("inner", width=event.width)

        self._field_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.create_window((0, 0), window=self._field_frame, anchor="nw", tags="inner")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self._check_vars = {}  # field -> [BooleanVar, ...]
        self._indicators = {}
        self._field_rows = {}
        self._populate()

        # Bottom buttons
        act_row = ttk.Frame(self.win)
        act_row.pack(fill="x", padx=15, pady=(10, 15), side="bottom")
        ttk.Button(act_row, text="仅选主记录", command=self._select_primary_only).pack(side="left", padx=2)
        ttk.Button(act_row, text="全选有值", command=self._select_all_nonempty).pack(side="left", padx=2)
        ttk.Button(act_row, text="全部取消", command=self._select_none).pack(side="left", padx=2)
        ttk.Button(act_row, text="执行合并", command=self._do_merge).pack(side="right", padx=3)
        ttk.Button(act_row, text="取消", command=self.win.destroy).pack(side="right", padx=3)

    def _case_val_str(self, case_idx, field):
        val = self.cases[case_idx].get(field)
        if val is not None and val != "" and val != "None":
            return str(val)
        return ""

    def _populate(self):
        for w in self._field_frame.winfo_children():
            w.destroy()
        self._check_vars.clear()
        self._indicators.clear()
        self._field_rows.clear()

        for field, label in _build_merge_fields(self.cases):
            row = ttk.Frame(self._field_frame)
            row.pack(fill="x", pady=1)
            self._field_rows[field] = row

            ttk.Label(row, text=label, width=14, font=("", 9)).pack(side="left")

            vars_list = []
            # Compute smart defaults
            if self._smart_defaults:
                # Collect which cases have non-empty values for this field
                has_val = [bool(self._case_val_str(i, field)) for i in range(len(self.cases))]
                n_with_val = sum(has_val)
                # Links: always keep all
                is_link = field in ("source_urls", "image_urls")
                if is_link:
                    defaults = has_val  # check all that have values
                elif n_with_val == 0:
                    defaults = [False] * len(self.cases)
                elif n_with_val == 1:
                    defaults = has_val  # check the only one that has value
                else:
                    # Check if all values are the same
                    vals = [self._case_val_str(i, field) for i in range(len(self.cases))]
                    non_empty = [v for v in vals if v]
                    if len(set(non_empty)) == 1:
                        first_idx = has_val.index(True)
                        defaults = [i == first_idx for i in range(len(self.cases))]
                    else:
                        defaults = [False] * len(self.cases)  # different values — user picks
            else:
                primary_has = bool(self._case_val_str(0, field))
                defaults = [(i == 0 and primary_has) for i in range(len(self.cases))]

            for i, case in enumerate(self.cases):
                val_str = self._case_val_str(i, field)
                var = tk.BooleanVar(value=defaults[i])
                vars_list.append(var)
                cb = ttk.Checkbutton(row, text=val_str[:48] or "(空)", variable=var,
                                     command=lambda f=field: self._on_choice(f))
                cb.pack(side="left", padx=(0, 6))

            self._check_vars[field] = vars_list

            ind = ttk.Label(row, text="", width=6, font=("", 8))
            ind.pack(side="left", padx=(5, 0))
            self._indicators[field] = ind

        self._update_indicators()

    def _on_choice(self, field):
        self._update_indicators()

    def _update_indicators(self):
        for field, vars_list in self._check_vars.items():
            indicator = self._indicators.get(field)
            if not indicator:
                continue
            checked = sum(1 for v in vars_list if v.get())
            if checked == 0:
                indicator.config(text="未选", foreground="#e65100")
            elif checked == 1:
                indicator.config(text="单选", foreground="#1565c0")
            else:
                indicator.config(text=f"多选{checked}", foreground="#2e7d32")

    def _select_primary_only(self):
        for field, vars_list in self._check_vars.items():
            for i, var in enumerate(vars_list):
                var.set(i == 0 and bool(self._case_val_str(0, field)))
        self._update_indicators()

    def _select_all_nonempty(self):
        for field, vars_list in self._check_vars.items():
            for i, var in enumerate(vars_list):
                var.set(bool(self._case_val_str(i, field)))
        self._update_indicators()

    def _select_none(self):
        for vars_list in self._check_vars.values():
            for var in vars_list:
                var.set(False)
        self._update_indicators()

    def _do_merge(self):
        try:
            self._do_merge_impl()
        except Exception as e:
            import traceback
            messagebox.showerror("合并失败", f"{e}\n\n{traceback.format_exc()}")

    def _do_merge_impl(self):
        primary_id = self.cases[0]["id"]
        secondary_ids = [c["id"] for c in self.cases[1:]]

        field_updates = {}
        for field, vars_list in self._check_vars.items():
            selected = []
            seen = set()
            for i, var in enumerate(vars_list):
                if var.get():
                    val = self._case_val_str(i, field)
                    if val and val not in seen:
                        selected.append(val)
                        seen.add(val)

            if selected:
                new_val = "; ".join(selected)
                if new_val != self._case_val_str(0, field):
                    field_updates[field] = new_val
            else:
                if self._case_val_str(0, field):
                    field_updates[field] = ""

        if self._table == "verified":
            filled = db.merge_verified_cases(primary_id, secondary_ids, field_updates)
        else:
            filled = db.merge_multiple_raw_cases(primary_id, secondary_ids, field_updates)

        msg = f"合并完成。\n主记录 ID={primary_id}"
        if filled:
            msg += f"\n更新了 {len(filled)} 个字段"
        messagebox.showinfo("合并结果", msg)
        self.win.destroy()
        if self.on_done:
            self.on_done()


# ============================================================
#  Shared Dialogs
class _RawCaseDetailDialog:
    """Show full details of a raw case with image thumbnails."""

    def __init__(self, parent, case):
        self.win = tk.Toplevel(parent)
        self.win.title(f"案例详情 - {case.get('name', '')[:30]}")
        self.win.geometry("700x650")
        self.win.transient(parent)

        # Image section
        _img_urls, _merged_path, _urls = merged_image_sources(case)
        img_section = ttk.LabelFrame(self.win, text=" 建筑图片（点击放大） ", padding=5)
        img_section.pack(fill="x", padx=10, pady=(10, 0))
        ImageThumbnailStrip(img_section, _img_urls, _merged_path).pack(fill="x")

        # Text details
        text = tk.Text(self.win, wrap="word", font=("", 11))
        text.pack(fill="both", expand=True, padx=10, pady=(8, 10))

        lines = []
        for f in db.RAW_FIELD_NAMES:
            val = case.get(f)
            label = {name: lbl for name, lbl, *_ in db.RAW_FIELDS}.get(f, f)
            if val is not None:
                lines.append(f"{label}: {val}")
        text.insert("1.0", "\n".join(lines))
        text.config(state="disabled")

        ttk.Button(self.win, text="关闭", command=self.win.destroy).pack(pady=(0, 10))


class _CaseDetailDialog:
    """Show full case details with verification status and source traceability."""

    def __init__(self, parent, case):
        self.case = case
        case_id = case.get("id")

        # Load field sources and build lookup: field_name -> latest source info
        self._field_sources = db.get_field_sources(case_id=case_id)
        self._source_lookup = {}
        for fs in self._field_sources:
            fn = fs.get("field_name", "")
            if fn:
                self._source_lookup[fn] = fs  # overwrite with latest

        self.win = tk.Toplevel(parent)
        name_short = (case.get("name", "") or "")[:30]
        self.win.title(f"案例详情 - {name_short}")
        self.win.geometry("820x900")
        self.win.transient(parent)

        canvas = tk.Canvas(self.win)
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw", width=800)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ---- 1. Image thumbnails ----
        _img_urls, _merged_path, _urls = merged_image_sources(case)
        img_section = ttk.LabelFrame(self.scroll_frame, text=" 建筑图片（点击放大） ", padding=5)
        img_section.pack(fill="x", padx=10, pady=(10, 0))
        ImageThumbnailStrip(img_section, _img_urls, _merged_path).pack(fill="x")

        # ---- 2. Field values with per-field source tags ----
        fields_frame = ttk.LabelFrame(self.scroll_frame, text=" 基本信息与来源追溯 ", padding=5)
        fields_frame.pack(fill="x", padx=10, pady=(8, 0))

        field_labels_dict = {f[0]: f[1] for f in db.FIELDS}

        source_type_labels = {
            "manual": "手动录入", "ai_research": "AI补全", "ai": "AI提取",
            "web": "网络搜索", "unknown": "未知来源",
        }

        for field, label, *_t in db.FIELDS:
            row = ttk.Frame(fields_frame)
            row.pack(fill="x", pady=1)

            ttk.Label(row, text=f"{label}：", font=("", 9, "bold"), width=14,
                     anchor="e").pack(side="left", padx=(2, 5))

            val = case.get(field)
            display = str(val) if val is not None else "—"

            # Source tag
            src = self._source_lookup.get(field)
            if src:
                stype = src.get("source_type", "unknown")
                st_label = source_type_labels.get(stype, stype)
                tag_color = {"manual": "#2e7d32", "ai_research": "#1565c0",
                            "ai": "#6a1b9a", "web": "#e65100"}.get(stype, "#757575")
                tag_text = f"[{st_label}]"
            else:
                tag_text = "[无记录]"
                tag_color = "#9e9e9e"

            val_frame = ttk.Frame(row)
            val_frame.pack(side="left", fill="x", expand=True)

            ttk.Label(val_frame, text=display, wraplength=420, justify="left",
                     font=("", 9)).pack(side="left", anchor="w")
            ttk.Label(val_frame, text=tag_text, font=("", 8), foreground=tag_color,
                     cursor="hand2").pack(side="left", padx=(8, 0))

        # ---- 3. Case-level verification summary ----
        trace_frame = ttk.LabelFrame(self.scroll_frame, text=" 案例核验摘要 ", padding=5)
        trace_frame.pack(fill="x", padx=10, pady=(8, 0))

        status = case.get("verification_status", "unverified")
        status_labels = {
            "verified": "✓ 已验证", "unverified": "○ 未验证",
            "needs_review": "⚠ 需复核", "low_confidence": "✗ 可信度低"
        }
        ttk.Label(trace_frame, text=f"验证状态：{status_labels.get(status, status)}",
                 font=("", 10, "bold")).pack(anchor="w")

        cred = case.get("source_credibility", "unknown")
        cred_labels = {"academic": "学术", "government": "政府", "news": "新闻",
                      "blog": "博客", "unknown": "未知"}
        ttk.Label(trace_frame, text=f"来源可信度：{cred_labels.get(cred, cred)}",
                 font=("", 9)).pack(anchor="w", pady=(2, 0))

        score = case.get("completeness_score")
        score_display = f"{float(score):.0%}" if score else "—"
        ttk.Label(trace_frame, text=f"完整度：{score_display}",
                 font=("", 9)).pack(anchor="w", pady=(2, 0))

        raw_ids = case.get("raw_case_ids", "")
        if raw_ids:
            ttk.Label(trace_frame, text=f"来源原始案例ID：{raw_ids}",
                     font=("", 9), foreground="#555").pack(anchor="w", pady=(2, 0))

        # ---- 4. Field sources detail table ----
        if self._field_sources:
            src_frame = ttk.LabelFrame(self.scroll_frame, text=" 字段来源详细记录 ", padding=5)
            src_frame.pack(fill="x", padx=10, pady=(8, 0))
            fs_tree = ttk.Treeview(src_frame,
                columns=("field", "value", "source", "type", "time"),
                show="headings", height=min(len(self._field_sources), 6))
            fs_tree.heading("field", text="字段")
            fs_tree.heading("value", text="值")
            fs_tree.heading("source", text="来源URL（可复制）")
            fs_tree.heading("type", text="类型")
            fs_tree.heading("time", text="记录时间")
            fs_tree.column("field", width=80)
            fs_tree.column("value", width=130)
            fs_tree.column("source", width=300)
            fs_tree.column("type", width=65)
            fs_tree.column("time", width=120)
            fs_tree.pack(fill="x")
            for fs in self._field_sources:
                stype = source_type_labels.get(fs.get("source_type", ""), fs.get("source_type", ""))
                source_url = fs.get("source_url") or ""
                fs_tree.insert("", "end", values=(
                    field_labels_dict.get(fs.get("field_name", ""), fs.get("field_name", "")),
                    (fs.get("field_value") or "")[:50],
                    source_url,
                    stype,
                    fs.get("created_at", ""),
                ))

        # ---- 5. Verification logs ----
        if raw_ids:
            try:
                raw_id_list = [int(x.strip()) for x in raw_ids.split(",") if x.strip()]
            except ValueError:
                raw_id_list = []
            all_logs = []
            for rid in raw_id_list:
                logs = db.get_verification_logs(rid)
                all_logs.extend(logs)
            if all_logs:
                log_frame = ttk.LabelFrame(self.scroll_frame, text=" 验证日志 ", padding=5)
                log_frame.pack(fill="x", padx=10, pady=(8, 0))
                log_text = tk.Text(log_frame, wrap="word", font=("", 9), height=5)
                log_text.pack(fill="both", expand=True)
                log_lines = []
                for log in all_logs:
                    log_lines.append(
                        f"[{log.get('created_at','')}] raw#{log.get('raw_case_id','')} "
                        f"{log.get('action','')}: {log.get('result','') or log.get('details','')}"
                    )
                log_text.insert("1.0", "\n".join(log_lines))
                log_text.config(state="disabled")

        # ---- Buttons ----
        btn_bar = ttk.Frame(self.win)
        btn_bar.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_bar, text="编辑", command=self._edit).pack(side="left", padx=3)
        ttk.Button(btn_bar, text="关闭", command=self.win.destroy).pack(side="right", padx=3)

    def _edit(self):
        dlg = _EditDialog(self.win, self.case)
        self.win.wait_window(dlg.win)
        updated = db.get_case_by_id(self.case.get("id"))
        if updated:
            self.win.destroy()
            _CaseDetailDialog(self.win.master, updated)



class _EditDialog:
    """Simple case editor (reused from search_tab.EditDialog pattern)."""

    def __init__(self, parent, case_data):
        self.case_data = case_data
        self.result = None

        # Detect whether this is a raw case or a regular case
        self._is_raw = "construction_year" in case_data or "function" in case_data
        fields_def = db.RAW_FIELDS if self._is_raw else db.FIELDS

        self.win = tk.Toplevel(parent)
        is_new = not case_data.get("id")
        self.win.title("新增案例" if is_new else f"编辑案例 #{case_data['id']}")
        self.win.geometry("700x850")
        self.win.transient(parent)

        canvas = tk.Canvas(self.win)
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)

        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw", width=680)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Image section
        _img_urls, _merged_path, _urls = merged_image_sources(case_data)
        img_section = ttk.LabelFrame(self.scroll_frame, text=" 建筑图片（点击放大） ", padding=5)
        img_section.pack(fill="x", padx=10, pady=(10, 0))
        ImageThumbnailStrip(img_section, _img_urls, _merged_path).pack(fill="x")

        self._entries = {}
        for field, label, type_def in fields_def:
            row = ttk.Frame(self.scroll_frame)
            row.pack(fill="x", padx=10, pady=2)
            ttk.Label(row, text=label, width=14, anchor="e").pack(side="left", padx=3)
            if field in ("description", "notes", "authenticity_notes"):
                entry = tk.Text(row, height=3, width=40)
                val = case_data.get(field)
                if val:
                    entry.insert("1.0", str(val))
                entry.pack(side="left", fill="x", expand=True)
                self._entries[field] = entry
            elif field == "building_type":
                entry = ttk.Combobox(row, values=db.BUILDING_TYPES, state="readonly")
                entry.set(case_data.get(field, "") or "")
                entry.pack(side="left", fill="x", expand=True)
                self._entries[field] = entry
            elif field == "condition":
                entry = ttk.Combobox(row, values=db.CONDITION_OPTIONS, state="readonly")
                entry.set(case_data.get(field, "") or "")
                entry.pack(side="left", fill="x", expand=True)
                self._entries[field] = entry
            else:
                var = tk.StringVar(value=case_data.get(field, "") or "")
                entry = ttk.Entry(row, textvariable=var)
                entry.pack(side="left", fill="x", expand=True)
                self._entries[field] = var

        btn_bar = ttk.Frame(self.scroll_frame)
        btn_bar.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_bar, text="保存", command=self._save).pack(side="right", padx=3)
        ttk.Button(btn_bar, text="取消", command=self.win.destroy).pack(side="right", padx=3)

    def _save(self):
        data = {}
        for field, widget in self._entries.items():
            if isinstance(widget, tk.Text):
                val = widget.get("1.0", "end-1c").strip()
            elif isinstance(widget, ttk.Combobox):
                val = widget.get().strip()
            else:
                val = widget.get().strip()
            data[field] = val if val else None

        if not data.get("name"):
            messagebox.showwarning("提示", "建筑名称不能为空。")
            return

        num_fields = ["area_sqm", "latitude", "longitude", "authenticity_score"]
        for num_field in num_fields:
            try:
                if data.get(num_field):
                    data[num_field] = float(data[num_field])
                else:
                    data[num_field] = None
            except (ValueError, TypeError):
                messagebox.showerror("错误", f"{num_field} 请输入数字。")
                return

        try:
            if data.get("quantity"):
                data["quantity"] = int(float(data["quantity"]))
        except (ValueError, TypeError):
            messagebox.showerror("错误", "数量请输入整数。")
            return

        case_id = self.case_data.get("id")
        if case_id:
            if self._is_raw:
                db.update_raw_case(case_id, data)
                db.sync_raw_to_case(case_id)
            else:
                # Record field sources for changed fields
                for field, new_val in data.items():
                    old_val = self.case_data.get(field)
                    old_str = str(old_val) if old_val is not None else ""
                    new_str = str(new_val) if new_val is not None else ""
                    if old_str != new_str:
                        db.insert_field_source(case_id=case_id, field_name=field,
                                               field_value=new_str, source_url="",
                                               source_type="manual", confidence=1.0)
                db.update_case(case_id, data)
        else:
            if self._is_raw:
                db.insert_raw_case(data)
            else:
                new_id = db.insert_case(data)
                # Record all non-empty fields as manual entry
                for field, val in data.items():
                    if val:
                        db.insert_field_source(case_id=new_id, field_name=field,
                                               field_value=str(val), source_url="",
                                               source_type="manual", confidence=1.0)

        self.win.destroy()


def _parse_name_details(text, name_info):
    """Extract structured name fields from 【名称详情】 section in conclusion text."""
    import re
    detail_match = re.search(r'【名称详情】\s*\n(.*?)(?=\n\n|\Z)', text, re.DOTALL)
    if not detail_match:
        return
    detail_text = detail_match.group(1)
    patterns = {
        "chinese_name": [r'中文[名称]?\s*[：:]\s*(.+)'],
        "english_name": [r'English\s*[：:]\s*(.+)', r'英文[名称]?\s*[：:]\s*(.+)'],
        "official_language_name": [r'官方语言[名称]?\s*[：:]\s*(.+)'],
        "common_names": [r'常用名[称]?\s*[：:]\s*(.+)'],
        "other_names": [r'其他名[称]?\s*[：:]\s*(.+)'],
    }
    for field, pat_list in patterns.items():
        for pat in pat_list:
            m = re.search(pat, detail_text)
            if m:
                val = m.group(1).strip()
                if field in ("common_names", "other_names"):
                    names = [n.strip() for n in re.split(r'[,，、]', val) if n.strip()]
                    if names and not name_info.get(field):
                        name_info[field] = names
                else:
                    if val and not name_info.get(field):
                        name_info[field] = val
                break


def _build_name_info(case):
    """Build name_info dict for a case using the same logic as the verification details dialog."""
    import json as _json

    # Resolve raw_case_id
    raw_case_id = case.get("raw_case_id", case.get("id"))
    if "year_built" in case or "building_type" in case:
        raw_ids_str = case.get("raw_case_ids", "") or ""
        if raw_ids_str:
            first_id = raw_ids_str.split(",")[0]
            try:
                raw_case_id = int(first_id)
            except ValueError:
                pass

    name_info = {
        "chinese_name": case.get("name", ""),
        "english_name": case.get("name_en", ""),
        "official_language_name": case.get("name_official", ""),
        "common_names": [],
        "other_names": [],
    }

    field_labels_map_full = {
        "name": "建筑名称", "country": "国家", "city": "城市",
        "construction_year": "建造时间", "demolition_year": "拆除时间",
        "restoration_year": "修复时间", "building_status": "建筑状态",
        "spread_method": "传播方式",
        "quantity": "数量", "function": "功能", "designer": "设计师",
        "coordinates": "经纬度", "image_urls": "图片",
    }

    try:
        logs = db.get_verification_logs(raw_case_id)
    except Exception:
        logs = []

    if logs:
        for log in reversed(logs):
            action = log.get("action", "")
            details = log.get("details", "")
            result = log.get("result", "")

            if action == "manual_result":
                try:
                    r = _json.loads(result)
                    conclusion = r.get("conclusion", "")
                    parts = conclusion.split("[其他字段]")
                    name_part = parts[0].replace("[名称字段]", "").strip() if parts else ""
                    if name_part:
                        _parse_name_details(name_part, name_info)
                except (_json.JSONDecodeError, TypeError):
                    pass

            elif action in ("field_verify_ai", "field_verify_gemini"):
                try:
                    r = _json.loads(result)
                    if isinstance(r, dict) and "results" in r:
                        results_list = r["results"]
                    elif isinstance(r, list):
                        results_list = r
                    else:
                        results_list = [r]
                    is_name = False
                    for fn, fl in field_labels_map_full.items():
                        if fl in details or fn in details.lower():
                            if fn == "name":
                                is_name = True
                            break
                    if is_name:
                        for item in results_list:
                            if not isinstance(item, dict):
                                continue
                            if item.get("chinese_name"):
                                name_info["chinese_name"] = item["chinese_name"]
                            if item.get("english_name"):
                                name_info["english_name"] = item["english_name"]
                            if item.get("official_language_name"):
                                name_info["official_language_name"] = item["official_language_name"]
                            if item.get("common_names"):
                                name_info["common_names"] = item["common_names"]
                            if item.get("other_names"):
                                name_info["other_names"] = item["other_names"]
                except (_json.JSONDecodeError, TypeError, AttributeError):
                    pass

            elif action == "info_analysis":
                try:
                    r = _json.loads(result)
                    completions = r.get("completions", {})
                    if not completions and any(k in field_labels_map_full for k in r):
                        completions = r
                    for fn, info in completions.items():
                        if isinstance(info, dict) and info.get("value"):
                            if fn == "name" and info.get("value"):
                                name_info["chinese_name"] = info["value"]
                except (_json.JSONDecodeError, TypeError):
                    pass

            elif action == "auto_complete":
                try:
                    r = _json.loads(result)
                    completions = r.get("completions", {})
                    if not completions and any(k in field_labels_map_full for k in r):
                        completions = r
                    for fn, info in completions.items():
                        if isinstance(info, dict) and info.get("value"):
                            if fn == "name" and info.get("value"):
                                name_info["chinese_name"] = info["value"]
                except (_json.JSONDecodeError, TypeError):
                    pass

    # Fill empty values with defaults
    if not name_info["chinese_name"]:
        name_info["chinese_name"] = case.get("name", "")
    if not name_info["english_name"]:
        name_info["english_name"] = case.get("name_en", "") or ""
    if not name_info["official_language_name"]:
        name_info["official_language_name"] = case.get("name_official", "") or ""

    return name_info


def _build_verification_display_data(case, lang="cn"):
    """Return (name_results_text, field_verify_text, ref_links_text) for export.

    Mirrors the formal verification details dialog display exactly.
    """
    import json as _json

    # Resolve raw_case_id
    raw_case_id = case.get("raw_case_id", case.get("id"))
    if "year_built" in case or "building_type" in case:
        raw_ids_str = case.get("raw_case_ids", "") or ""
        if raw_ids_str:
            first_id = raw_ids_str.split(",")[0]
            try:
                raw_case_id = int(first_id)
            except ValueError:
                pass

    field_labels_map_full = {
        "name": "建筑名称", "country": "国家", "city": "城市",
        "construction_year": "建造时间", "demolition_year": "拆除时间",
        "restoration_year": "修复时间", "building_status": "建筑状态",
        "spread_method": "传播方式",
        "quantity": "数量", "function": "功能", "designer": "设计师",
        "coordinates": "经纬度", "image_urls": "图片",
    }

    try:
        logs = db.get_verification_logs(raw_case_id)
    except Exception:
        logs = []

    all_name_result_texts = []
    all_other_result_texts = []

    if logs:
        for log in reversed(logs):
            action = log.get("action", "")
            details = log.get("details", "")
            result = log.get("result", "")

            if action == "manual_result":
                try:
                    r = _json.loads(result)
                    conclusion = r.get("conclusion", "")
                    if conclusion:
                        name_m = re.search(r'\[名称字段\]\s*\n(.*?)(?=\[其他字段\]|\Z)', conclusion, re.DOTALL)
                        other_m = re.search(r'\[其他字段\]\s*\n(.*?)(?=\Z)', conclusion, re.DOTALL)
                        name_part = name_m.group(1).strip() if name_m else ""
                        other_part = other_m.group(1).strip() if other_m else ""
                        if name_part:
                            all_name_result_texts.append((details, name_part))
                        if other_part:
                            all_other_result_texts.append((details, other_part))
                        if not name_part and not other_part:
                            all_other_result_texts.append((details, conclusion))
                except (_json.JSONDecodeError, TypeError):
                    pass

            elif action == "field_verify_text":
                if result and result.strip():
                    if details.startswith("[名称]"):
                        all_name_result_texts.append((details[4:], result))
                    else:
                        all_other_result_texts.append((details, result))

    is_en = (lang == "en")
    name_info = _build_name_info(case)

    field_labels_en = {
        "建筑名称": "Name", "国家": "Country", "城市": "City",
        "建造时间": "Year Built", "拆除时间": "Demolition Year",
        "修复时间": "Restoration Year", "建筑状态": "Condition",
        "传播方式": "Spread Method", "数量": "Quantity",
        "功能": "Building Type", "设计师": "Architect",
        "经纬度": "Coordinates", "图片": "Images",
    }

    # --- 名称核验结果: 结论 + 参考链接 (same as formal dialog) ---
    name_parts = []
    name_parts.append("【Conclusion】" if is_en else "【结论】")
    cn_name = name_info.get("chinese_name", "")
    en_name = name_info.get("english_name", "")
    off_name = name_info.get("official_language_name", "")
    common = name_info.get("common_names", [])
    other = name_info.get("other_names", [])
    conclusion_lines = []
    if cn_name:
        conclusion_lines.append(f"Chinese: {cn_name}" if is_en else f"中文名：{cn_name}")
    if en_name:
        conclusion_lines.append(f"English: {en_name}" if is_en else f"英文名：{en_name}")
    if off_name:
        conclusion_lines.append(f"Official: {off_name}" if is_en else f"官方语言名：{off_name}")
    if common:
        cn_str = common if isinstance(common, str) else ("、".join(common) if not is_en else ", ".join(common))
        conclusion_lines.append(f"Common Names: {cn_str}" if is_en else f"常用名：{cn_str}")
    if other:
        on_str = other if isinstance(other, str) else ("、".join(other) if not is_en else ", ".join(other))
        conclusion_lines.append(f"Other Names: {on_str}" if is_en else f"其他名：{on_str}")
    name_parts.append("\n".join(conclusion_lines) if conclusion_lines else "—")
    name_parts.append("")
    name_parts.append("【Reference Links】" if is_en else "【参考链接】")
    name_refs = []
    if all_name_result_texts:
        for src_info, text in all_name_result_texts:
            for ref in _VerificationResultDialog._parse_name_references(text):
                if ref not in name_refs:
                    name_refs.append(ref)
    if name_refs:
        for ref in name_refs:
            name_parts.append(f"  • {ref}")
    else:
        name_parts.append("—")
    name_results_text = "\n".join(name_parts)

    # --- 其他字段核验结果: 值 + 参考链接 (same as formal dialog) ---
    field_parts = []
    if all_other_result_texts:
        other_field_data = _VerificationResultDialog._parse_other_field_data(all_other_result_texts, field_labels_map_full)
        if other_field_data:
            field_order = [fl for fk, fl in field_labels_map_full.items() if fk != "name"]
            for fl_cn in [f for f in field_order if f in other_field_data]:
                entries = other_field_data[fl_cn]
                values = []
                refs_all = []
                for entry in entries:
                    val = entry.get("value", "")
                    if val and val not in values:
                        val = DatabasePanel._translate_text(val, to_en=is_en)
                        values.append(val)
                    for ref in entry.get("references", []):
                        normalized = ref.rstrip('/')
                        if not any(r.rstrip('/') == normalized for r in refs_all):
                            refs_all.append(ref)
                fl_display = field_labels_en.get(fl_cn, fl_cn) if is_en else fl_cn
                field_parts.append(f"【{fl_display}】")
                field_parts.append(f"  Value: {' / '.join(values)}" if is_en else f"  值：{' / '.join(values)}")
                if refs_all:
                    field_parts.append(f"  Reference Links:" if is_en else f"  参考链接：")
                    for ref in refs_all:
                        field_parts.append(f"    • {ref}")
                field_parts.append("")
        else:
            for i, (src_info, text) in enumerate(all_other_result_texts):
                if i > 0:
                    field_parts.append("─" * 40)
                    field_parts.append("")
                field_parts.append(f"── [{src_info}] ──")
                field_parts.append(DatabasePanel._translate_text(text, to_en=is_en))
                field_parts.append("")
    field_verify_text = "\n".join(field_parts).strip()

    # --- 参考链接: 来源列表 ---
    try:
        sources = db.get_sources_for_raw_case(raw_case_id)
    except Exception:
        sources = []
    ref_urls = [s.get("url", "") for s in sources if s.get("url")]
    ref_links_text = "\n".join(ref_urls)

    return name_results_text, field_verify_text, ref_links_text


class _VerificationResultDialog:
    """Dialog showing case details, name info, verification results and logs."""

    def __init__(self, parent, case, is_formal=False):
        self.win = tk.Toplevel(parent)
        self._parent = parent
        self._case = case
        self._is_formal = is_formal
        name_short = (case.get("name", "") or "")[:30]
        if is_formal:
            self.win.title(f"正式案例核验详情 - {name_short}")
        else:
            self.win.title(f"核验详情 - {name_short}")
        self.win.geometry("750x850")
        self.win.transient(parent)

        # --- Refresh button ---
        top_btn_frame = ttk.Frame(self.win)
        top_btn_frame.pack(pady=(5, 0))
        refresh_btn = ttk.Button(top_btn_frame, text="刷新", command=self._refresh)
        refresh_btn.pack(side="left", padx=2)

        # --- Build name_info from shared function ---
        name_info = _build_name_info(case)

        # --- Collect verification logs and extract display data ---
        raw_case_id = case.get("raw_case_id", case.get("id"))
        if "year_built" in case or "building_type" in case:
            raw_ids_str = case.get("raw_case_ids", "") or ""
            if raw_ids_str:
                first_id = raw_ids_str.split(",")[0]
                try:
                    raw_case_id = int(first_id)
                except ValueError:
                    pass
        logs = db.get_verification_logs(raw_case_id)
        all_name_result_texts = []  # all name result texts (appended in order)
        all_other_result_texts = []  # all other-field result texts (appended in order)
        name_operations = []  # [{source, time, search_content, result}] name field ops in order
        search_contents = []  # (action_desc, search_content) pairs for display

        # Field labels for display
        field_labels_map_full = {
            "name": "建筑名称", "country": "国家", "city": "城市",
            "construction_year": "建造时间", "demolition_year": "拆除时间",
            "restoration_year": "修复时间", "building_status": "建筑状态",
            "spread_method": "传播方式",
            "quantity": "数量", "function": "功能", "designer": "设计师",
            "coordinates": "经纬度", "image_urls": "图片",
        }

        if logs:
            for log in reversed(logs):
                action = log.get("action", "")
                details = log.get("details", "")
                result = log.get("result", "")

                if action in ("field_verify_ai", "field_verify_deepseek", "field_verify_gemini"):
                    try:
                        r = json.loads(result)
                        # Check if this is about the name field
                        is_name = False
                        for fn, fl in field_labels_map_full.items():
                            if fl in details or fn in details.lower():
                                if fn == "name":
                                    is_name = True
                                break
                        if is_name:
                            # Record full operation context for name field
                            sc = ""
                            if isinstance(r, dict):
                                sc = r.get("search_content", "")
                            name_operations.append({
                                "_source": details,
                                "_time": log.get("created_at", "")[:19],
                                "_search_content": sc,
                                "_result": r,
                            })
                        else:
                            # Other fields: extract search content
                            if isinstance(r, dict) and "results" in r:
                                sc = r.get("search_content", "")
                                if sc:
                                    search_contents.append((details, sc))
                    except (json.JSONDecodeError, TypeError, AttributeError):
                        pass

                elif action == "name":
                    # Batch verification name step: details="current:xxx", result=JSON of full name result
                    try:
                        r = json.loads(result)
                        if isinstance(r, dict):
                            current_val = ""
                            if details and details.startswith("current:"):
                                current_val = details[8:]
                            name_text = f"核验时间: {log.get('created_at', '')[:19]}\n"
                            name_text += f"当前值: {current_val}\n"
                            name_text += f"正确值: {r.get('correct_value', '')}\n"
                            name_text += f"中文名: {r.get('chinese_name', '')}\n"
                            name_text += f"英文名: {r.get('english_name', '')}\n"
                            name_text += f"官方语言名: {r.get('official_language_name', '')}\n"
                            name_text += f"常用名: {r.get('common_names', '')}\n"
                            name_text += f"置信度: {r.get('confidence', '')}\n"
                            name_text += f"结论: {r.get('conclusion', '')}\n"
                            name_text += f"推理: {r.get('reasoning', '')}\n"
                            refs = r.get('references', [])
                            if refs:
                                name_text += "参考来源:\n"
                                for ref in refs:
                                    name_text += f"  - {ref}\n"
                            all_name_result_texts.append((details, name_text))
                    except (json.JSONDecodeError, TypeError):
                        pass

                elif action == "info_analysis":
                    try:
                        r = json.loads(result)
                        sc = r.get("search_content", "")
                        if sc:
                            search_contents.append((details, sc))
                    except (json.JSONDecodeError, TypeError):
                        pass

                elif action == "auto_complete":
                    pass

                elif action == "search_result":
                    # Browser search result content saved from Chrome extension
                    try:
                        r = json.loads(result)
                        sc = r.get("search_content", "")
                        if sc:
                            search_contents.append((details, sc))
                    except (json.JSONDecodeError, TypeError):
                        pass

                elif action == "manual_result":
                    try:
                        r = json.loads(result)
                        conclusion = r.get("conclusion", "")
                        if conclusion:
                            name_m = re.search(r'\[名称字段\]\s*\n(.*?)(?=\[其他字段\]|\Z)', conclusion, re.DOTALL)
                            other_m = re.search(r'\[其他字段\]\s*\n(.*?)(?=\Z)', conclusion, re.DOTALL)
                            name_part = name_m.group(1).strip() if name_m else ""
                            other_part = other_m.group(1).strip() if other_m else ""
                            if name_part:
                                all_name_result_texts.append((details, name_part))
                            if other_part:
                                all_other_result_texts.append((details, other_part))
                            if not name_part and not other_part:
                                all_other_result_texts.append((details, conclusion))
                    except (json.JSONDecodeError, TypeError):
                        pass

                elif action == "field_verify_text":
                    if result and result.strip():
                        if details.startswith("[名称]"):
                            all_name_result_texts.append((details[4:].lstrip(), result))
                        else:
                            all_other_result_texts.append((details, result))

        # --- Scrollable content area ---
        canvas = tk.Canvas(self.win, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", tags="inner")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_scroll(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_scroll(event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_scroll)
        canvas.bind("<Leave>", _unbind_scroll)

        def _on_canvas_configure(event):
            canvas.itemconfig("inner", width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- Name details ---
        name_frame = ttk.LabelFrame(inner, text=" 名称详情 ", padding=8)
        name_frame.pack(fill="x", padx=10, pady=(10, 5))

        name_rows = [
            ("中文名", name_info["chinese_name"] or "—"),
            ("英文名", name_info["english_name"] or "—"),
            ("官方语言名", name_info["official_language_name"] or "—"),
        ]
        for label, val in name_rows:
            row = ttk.Frame(name_frame)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f"{label}：", font=("", 9, "bold"), width=12,
                     anchor="e").pack(side="left", padx=(2, 5))
            ttk.Label(row, text=val, font=("", 9)).pack(side="left")

        if name_info["common_names"]:
            row = ttk.Frame(name_frame)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text="常用名：", font=("", 9, "bold"), width=12,
                     anchor="e").pack(side="left", padx=(2, 5))
            ttk.Label(row, text=", ".join(name_info["common_names"]), font=("", 9)).pack(side="left")

        if name_info["other_names"]:
            row = ttk.Frame(name_frame)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text="其他名：", font=("", 9, "bold"), width=12,
                     anchor="e").pack(side="left", padx=(2, 5))
            ttk.Label(row, text=", ".join(name_info["other_names"]), font=("", 9)).pack(side="left")

        # --- Image thumbnails ---
        img_urls, merged_path, _all_urls = merged_image_sources(case)
        img_section = ttk.LabelFrame(inner, text=" 建筑图片（点击放大） ", padding=5)
        img_section.pack(fill="x", padx=10, pady=5)
        if img_urls or merged_path:
            ImageThumbnailStrip(img_section, img_urls, merged_path).pack(fill="x")
        else:
            ttk.Label(img_section, text="暂无图片", foreground="#999").pack(pady=5)

        # --- 名称字段核验结果 ---
        # Show name-related operations in chronological order (oldest first)
        name_ops_display = list(reversed(name_operations))  # collected in reverse, flip to chronological
        res_frame = ttk.LabelFrame(inner, text=" 名称字段核验结果 ", padding=8)
        res_frame.pack(fill="x", padx=10, pady=5)

        if is_formal:
            # New format: 结论 + 依据 + 参考链接
            name_text = tk.Text(res_frame, wrap="word", font=("", 9), height=12)
            name_scroll = ttk.Scrollbar(res_frame, orient="vertical", command=name_text.yview)
            name_text.configure(yscrollcommand=name_scroll.set)
            name_text.pack(side="left", fill="x", expand=True)
            name_scroll.pack(side="right", fill="y")

            name_text.tag_configure("bold", font=("", 9, "bold"))
            name_text.insert("end", "【结论】\n", "bold")
            cn_name = name_info.get("chinese_name", "")
            en_name = name_info.get("english_name", "")
            off_name = name_info.get("official_language_name", "")
            common = name_info.get("common_names", [])
            other = name_info.get("other_names", [])
            if cn_name:
                name_text.insert("end", f"中文名：{cn_name}\n")
            if en_name:
                name_text.insert("end", f"英文名：{en_name}\n")
            if off_name:
                name_text.insert("end", f"官方语言名：{off_name}\n")
            if common:
                cn_str = common if isinstance(common, str) else "、".join(common)
                name_text.insert("end", f"常用名：{cn_str}\n")
            if other:
                on_str = other if isinstance(other, str) else "、".join(other)
                name_text.insert("end", f"其他名：{on_str}\n")
            if not any([cn_name, en_name, off_name, common, other]):
                name_text.insert("end", "—\n")

            name_text.insert("end", "\n【参考链接】\n", "bold")
            name_refs = []
            if all_name_result_texts:
                for src_info, text in all_name_result_texts:
                    for ref in self._parse_name_references(text):
                        if ref not in name_refs:
                            name_refs.append(ref)
            if name_refs:
                for ref in name_refs:
                    name_text.insert("end", f"  • {ref}\n")
            else:
                name_text.insert("end", "—\n")

            name_text.config(state="disabled")
        else:
            if all_name_result_texts:
                total_lines = sum(len(text.split("\n")) for _, text in all_name_result_texts)
                name_text = tk.Text(res_frame, wrap="word", font=("", 9),
                                    height=min(max(total_lines + 2, 6), 18))
                name_scroll = ttk.Scrollbar(res_frame, orient="vertical", command=name_text.yview)
                name_text.configure(yscrollcommand=name_scroll.set)
                name_text.pack(side="left", fill="x", expand=True)
                name_scroll.pack(side="right", fill="y")

                for i, (src_info, text) in enumerate(all_name_result_texts):
                    if i > 0:
                        name_text.insert("end", "\n" + "─" * 40 + "\n\n")
                    name_text.insert("end", f"── [{src_info}] ──\n{text}\n\n")
                name_text.config(state="disabled")
            else:
                name_text = tk.Text(res_frame, wrap="word", font=("", 9), height=2)
                name_text.pack(fill="x", expand=True)
                name_text.insert("end", "暂无名称核验记录")
                name_text.config(state="disabled")

        # --- 其他字段核验结果 ---
        other_frame = ttk.LabelFrame(inner, text=" 其他字段核验结果 ", padding=8)
        other_frame.pack(fill="x", padx=10, pady=5)

        if all_other_result_texts:
            if is_formal:
                # Parse structured field data from raw text
                other_field_data = self._parse_other_field_data(all_other_result_texts, field_labels_map_full)
                if other_field_data:
                    lines = []
                    field_order = [fl for fk, fl in field_labels_map_full.items() if fk != "name"]
                    for fl_cn in [f for f in field_order if f in other_field_data]:
                        entries = other_field_data[fl_cn]
                        values = []
                        refs_all = []
                        for entry in entries:
                            val = entry.get("value", "")
                            if val and val not in values:
                                values.append(val)
                            for ref in entry.get("references", []):
                                normalized = ref.rstrip('/')
                                if not any(r.rstrip('/') == normalized for r in refs_all):
                                    refs_all.append(ref)
                        lines.append(f"【{fl_cn}】")
                        lines.append(f"  值：{' / '.join(values)}")
                        if refs_all:
                            lines.append(f"  参考链接：")
                            for ref in refs_all:
                                lines.append(f"    • {ref}")
                        lines.append("")
                    display_text = "\n".join(lines)
                else:
                    display_text = ""
                    for i, (src_info, text) in enumerate(all_other_result_texts):
                        if i > 0:
                            display_text += "\n" + "─" * 40 + "\n\n"
                        display_text += f"── [{src_info}] ──\n{text}\n\n"
            else:
                display_text = ""
                for i, (src_info, text) in enumerate(all_other_result_texts):
                    if i > 0:
                        display_text += "\n" + "─" * 40 + "\n\n"
                    display_text += f"── [{src_info}] ──\n{text}\n\n"

            total_lines = len(display_text.split("\n"))
            other_text = tk.Text(other_frame, wrap="word", font=("", 9),
                                 height=min(max(total_lines + 2, 6), 18))
            other_scroll = ttk.Scrollbar(other_frame, orient="vertical", command=other_text.yview)
            other_text.configure(yscrollcommand=other_scroll.set)
            other_text.pack(side="left", fill="x", expand=True)
            other_scroll.pack(side="right", fill="y")
            other_text.insert("1.0", display_text)
            other_text.config(state="disabled")
        else:
            other_text = tk.Text(other_frame, wrap="word", font=("", 9), height=2)
            other_text.pack(fill="x", expand=True)
            other_text.insert("end", "暂无其他字段核验记录")
            other_text.config(state="disabled")


        # --- Modules only shown for authenticity verification ---
        if not is_formal:
            # --- Case full details ---
            detail_frame = ttk.LabelFrame(inner, text=" 案例详情 ", padding=8)
            detail_frame.pack(fill="x", padx=10, pady=5)

            is_cases_table = "year_built" in case or "building_type" in case
            if is_cases_table:
                fields_show = ["name", "name_en", "name_official", "country", "city",
                               "year_built", "demolition_year", "restoration_year",
                               "quantity", "building_type", "architect",
                               "latitude", "longitude", "search_time", "search_keywords",
                               "material", "area_sqm", "protection_level",
                               "condition", "description"]
            else:
                fields_show = ["name", "name_en", "name_official", "country", "city",
                               "construction_year", "demolition_year", "restoration_year",
                               "quantity", "function", "designer",
                               "latitude", "longitude", "search_time", "search_keywords",
                               "authenticity_status", "authenticity_score", "authenticity_notes"]
            field_label_map = {
                "name": "建筑名称", "name_en": "英文名称", "name_official": "官方语言名称",
                "country": "国家", "city": "城市",
                "construction_year": "建造时间", "demolition_year": "拆除时间",
                "restoration_year": "修复时间", "quantity": "数量规模",
                "function": "功能", "designer": "设计师",
                "latitude": "纬度", "longitude": "经度",
                "search_time": "检索时间", "search_keywords": "检索词",
                "authenticity_status": "真实性状态", "authenticity_score": "可信度",
                "authenticity_notes": "评估备注", "notes": "备注",
                "year_built": "建造时间", "building_type": "功能",
                "architect": "设计师", "material": "结构材料",
                "area_sqm": "建筑面积(㎡)", "protection_level": "保护级别",
                "condition": "建筑状态", "description": "特征描述",
                "building_status": "建筑状态", "image_urls": "图片链接",
                "source_urls": "来源链接", "spread_method": "传播方式",
            }
            case_text = tk.Text(detail_frame, wrap="word", font=("", 9), height=6)
            lines = []
            for f in fields_show:
                val = case.get(f)
                if val is not None and val != "":
                    label = field_label_map.get(f, f)
                    lines.append(f"{label}: {val}")
            case_text.insert("1.0", "\n".join(lines))
            case_text.config(state="disabled")
            case_text.pack(fill="x")

            # --- Editable notes field ---
            notes_frame = ttk.Frame(detail_frame)
            notes_frame.pack(fill="x", pady=(5, 0))
            ttk.Label(notes_frame, text="备注：", font=("", 9)).pack(side="left", anchor="n")
            notes_text = tk.Text(notes_frame, wrap="word", font=("", 9), height=3)
            notes_text.pack(side="left", fill="x", expand=True)
            notes_val = case.get("notes") or ""
            notes_text.insert("1.0", notes_val)

            case_id = case.get("id")
            _notes_original = notes_val
            def _auto_save_notes(event=None):
                nonlocal _notes_original
                try:
                    new_val = notes_text.get("1.0", "end-1c").strip()
                except tk.TclError:
                    return
                if new_val == _notes_original:
                    return
                _notes_original = new_val
                db.update_raw_case(case_id, {"notes": new_val})
            notes_text.bind("<FocusOut>", _auto_save_notes)

            # --- Edit history ---
            edit_logs = [l for l in logs if l.get("action") == "manual_edit"]
            edit_frame = ttk.LabelFrame(inner, text=" 修改记录 ", padding=8)
            edit_frame.pack(fill="x", padx=10, pady=5)

            if edit_logs:
                edit_text = tk.Text(edit_frame, wrap="word", font=("", 9),
                                    height=min(len(edit_logs) + 1, 6))
                edit_scroll = ttk.Scrollbar(edit_frame, orient="vertical", command=edit_text.yview)
                edit_text.configure(yscrollcommand=edit_scroll.set)
                edit_text.pack(side="left", fill="x", expand=True)
                edit_scroll.pack(side="right", fill="y")

                edit_lines = []
                for log in edit_logs:
                    ts = log.get("created_at", "")[:19]
                    details = log.get("details", "")
                    edit_lines.append(f"[{ts}] {details}")
                edit_text.insert("1.0", "\n".join(edit_lines))
                edit_text.config(state="disabled")
            else:
                ttk.Label(edit_frame, text="暂无修改记录", foreground="#999").pack()

            # --- 搜索结果内容 ---
            sc_frame = ttk.LabelFrame(inner, text=" 搜索结果内容 ", padding=8)
            sc_frame.pack(fill="x", padx=10, pady=5)
            if search_contents:
                sc_text = tk.Text(sc_frame, wrap="word", font=("", 9), height=16)
                sc_scroll = ttk.Scrollbar(sc_frame, orient="vertical", command=sc_text.yview)
                sc_text.configure(yscrollcommand=sc_scroll.set)
                sc_text.pack(side="left", fill="both", expand=True)
                sc_scroll.pack(side="right", fill="y")

                for i, (desc, content) in enumerate(search_contents):
                    if i > 0:
                        sc_text.insert("end", "\n" + "─" * 40 + "\n\n")
                    sc_text.insert("end", f"── [{desc}] ──\n{content}\n\n")
                sc_text.config(state="disabled")
            else:
                ttk.Label(sc_frame, text="暂无搜索结果 — 在真实性核验页面使用百度/谷歌/AI分析搜索字段后，搜索结果会记录在此处。",
                          foreground="#999").pack(pady=5)

            # --- Verification logs (non-edit) ---
            other_logs = [l for l in logs if l.get("action") not in ("manual_edit", "search_result")]
            log_frame = ttk.LabelFrame(inner, text=" 核验日志 ", padding=8)
            log_frame.pack(fill="x", padx=10, pady=(5, 10))

            if other_logs:
                log_text = tk.Text(log_frame, wrap="word", font=("", 9),
                                   height=min(len(other_logs) + 2, 6))
                log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
                log_text.configure(yscrollcommand=log_scroll.set)
                log_text.pack(side="left", fill="x", expand=True)
                log_scroll.pack(side="right", fill="y")

                lines = []
                for log in other_logs:
                    ts = log.get("created_at", "")[:19]
                    action = log.get("action", "")
                    details = log.get("details", "")
                    result = log.get("result", "")
                    lines.append(f"[{ts}] {action}")
                    if details:
                        lines.append(f"  {details}")
                    if result:
                        lines.append(f"  结果: {result[:300]}")
                        lines.append("")
                    lines.append("")
                log_text.insert("1.0", "\n".join(lines))
                log_text.config(state="disabled")
            else:
                ttk.Label(log_frame, text="暂无核验记录", foreground="#999").pack()


        # --- Sources list (editable) ---
        src_frame = ttk.LabelFrame(inner, text=" 来源列表 ", padding=8)
        src_frame.pack(fill="x", padx=10, pady=5)

        sources = db.get_sources_for_raw_case(raw_case_id)
        # Store source data by tree item id for lookup
        self._src_data = {}  # item_id -> source dict
        self._src_raw_case_id = raw_case_id

        src_tree = ttk.Treeview(src_frame, columns=("url", "cred"), show="headings", height=max(len(sources), 3))
        src_tree.heading("url", text="URL")
        src_tree.heading("cred", text="可信度")
        src_tree.column("url", width=420)
        src_tree.column("cred", width=80)
        src_tree.pack(fill="x")

        def _refresh_src_tree():
            src_tree.delete(*src_tree.get_children())
            self._src_data.clear()
            for s in db.get_sources_for_raw_case(self._src_raw_case_id):
                item = src_tree.insert("", "end", values=(s.get("url", ""), s.get("credibility", "unknown")))
                self._src_data[item] = s

        _refresh_src_tree()

        # --- Inline edit on double-click ---
        def _on_src_double_click(event):
            region = src_tree.identify_region(event.x, event.y)
            if region != "cell":
                return
            item = src_tree.identify_row(event.y)
            col = src_tree.identify_column(event.x)
            if not item or not col:
                return
            col_idx = int(col.replace("#", "")) - 1
            col_names = ["url", "credibility"]
            if col_idx >= len(col_names):
                return
            field = col_names[col_idx]
            old_val = src_tree.item(item, "values")[col_idx]

            # Get cell bbox for placing entry
            bbox = src_tree.bbox(item, col)
            if not bbox:
                return

            entry = tk.Entry(src_frame, font=("", 9))
            entry.insert(0, old_val)
            entry.select_range(0, "end")
            entry.place(x=bbox.x, y=bbox.y, width=bbox.width, height=bbox.height)
            entry.focus_set()

            def _save_edit(e=None):
                new_val = entry.get().strip()
                entry.destroy()
                if new_val != old_val and item in self._src_data:
                    src = self._src_data[item]
                    try:
                        db.update_case_source(src["id"], {field: new_val})
                        # Also update credibility in case_sources if editing cred
                        db.insert_verification_log(self._src_raw_case_id, "source_edit",
                                                   f"来源编辑: {field} {old_val} -> {new_val}")
                        _refresh_src_tree()
                    except Exception as ex:
                        messagebox.showerror("保存失败", str(ex))

            def _cancel_edit(e=None):
                entry.destroy()

            entry.bind("<Return>", _save_edit)
            entry.bind("<Escape>", _cancel_edit)
            entry.bind("<FocusOut>", lambda e: entry.destroy())

        src_tree.bind("<Double-1>", _on_src_double_click)

        # --- Right-click menu ---
        src_menu = tk.Menu(src_tree, tearoff=0)
        src_menu.add_command(label="复制链接", command=lambda: self._copy_source_url(src_tree))
        src_menu.add_command(label="在新窗口打开链接", command=lambda: self._open_source_url(src_tree))
        src_menu.add_separator()
        src_menu.add_command(label="编辑", command=lambda: _edit_selected())
        src_menu.add_command(label="删除", command=lambda: _delete_selected())

        def _on_src_right_click(event):
            item = src_tree.identify_row(event.y)
            if item:
                src_tree.selection_set(item)
                src_menu.post(event.x_root, event.y_root)

        src_tree.bind("<Button-2>" if sys.platform == "darwin" else "<Button-3>", _on_src_right_click)

        # --- Edit selected source via popup ---
        def _edit_selected():
            sel = src_tree.selection()
            if not sel:
                messagebox.showinfo("提示", "请先选中一条来源。")
                return
            item = sel[0]
            src = self._src_data.get(item)
            if not src:
                return
            _open_source_edit_dialog(src_frame, src, self._src_raw_case_id, _refresh_src_tree)

        # --- Delete selected source ---
        def _delete_selected():
            sel = src_tree.selection()
            if not sel:
                messagebox.showinfo("提示", "请先选中一条来源。")
                return
            item = sel[0]
            src = self._src_data.get(item)
            if not src:
                return
            url_preview = src.get("url", "")[:50]
            if not messagebox.askyesno("确认删除", f"确定要删除此来源吗？\n\n{url_preview}"):
                return
            try:
                db.delete_case_source(src["id"])
                db.insert_verification_log(self._src_raw_case_id, "source_delete",
                                           f"删除来源: {url_preview}")
                _refresh_src_tree()
            except Exception as ex:
                messagebox.showerror("删除失败", str(ex))

        # --- Buttons ---
        src_btn_bar = ttk.Frame(src_frame)
        src_btn_bar.pack(fill="x", pady=(5, 0))
        ttk.Button(src_btn_bar, text="添加来源", command=lambda: _open_source_add_dialog(src_frame, self._src_raw_case_id, _refresh_src_tree)).pack(side="left", padx=2)
        ttk.Button(src_btn_bar, text="编辑选中", command=_edit_selected).pack(side="left", padx=2)
        ttk.Button(src_btn_bar, text="删除选中", command=_delete_selected).pack(side="left", padx=2)

        ttk.Button(inner, text="关闭", command=self.win.destroy).pack(pady=(10, 10))

    @staticmethod
    def _write_name_result_item(text_widget, item):
        """Write a single name result item dict into a Text widget."""
        has_content = False
        correct = item.get("correct_value", "")
        if correct:
            text_widget.insert("end", f"正确值：{correct}\n")
            has_content = True
        if item.get("chinese_name") and item["chinese_name"] != correct:
            text_widget.insert("end", f"中文名：{item['chinese_name']}\n")
            has_content = True
        if item.get("english_name"):
            text_widget.insert("end", f"英文名：{item['english_name']}\n")
            has_content = True
        if item.get("official_language_name"):
            text_widget.insert("end", f"官方语言名：{item['official_language_name']}\n")
            has_content = True
        if item.get("common_names"):
            cn = item["common_names"]
            cn_str = cn if isinstance(cn, str) else ", ".join(cn)
            text_widget.insert("end", f"常用名：{cn_str}\n")
            has_content = True
        if item.get("other_names"):
            on = item["other_names"]
            on_str = on if isinstance(on, str) else ", ".join(on)
            text_widget.insert("end", f"其他名：{on_str}\n")
            has_content = True
        conclusion = item.get("conclusion", "")
        if conclusion:
            text_widget.insert("end", f"结论：{conclusion}\n")
            has_content = True
        reasoning = item.get("reasoning", "")
        if reasoning:
            text_widget.insert("end", f"依据：{reasoning}\n")
            has_content = True
        refs = item.get("references", [])
        if refs:
            text_widget.insert("end", "参考链接：\n")
            for ref in refs:
                text_widget.insert("end", f"  • {ref}\n")
            has_content = True
        sc = item.get("search_content", "")
        if sc and not has_content:
            text_widget.insert("end", sc)

    def _copy_source_url(self, src_tree):
        """Copy the URL of the selected source row to clipboard."""
        sel = src_tree.selection()
        if sel:
            values = src_tree.item(sel[0], "values")
            if values:
                self.win.clipboard_clear()
                self.win.clipboard_append(values[0])

    def _open_source_url(self, src_tree):
        """Open the selected source URL in the default browser."""
        sel = src_tree.selection()
        if sel:
            values = src_tree.item(sel[0], "values")
            if values:
                import webbrowser
                webbrowser.open(values[0])

    def _refresh(self):
        """Close and reopen the dialog with fresh data."""
        self.win.destroy()
        _VerificationResultDialog(self._parent, self._case, is_formal=self._is_formal)

    @staticmethod
    def _parse_name_references(text):
        """Extract reference URLs from name field verification raw text."""
        import re as _re
        refs = []
        # Lines starting with • after 参考 section
        in_refs = False
        for line in text.split('\n'):
            if '参考' in line and ('链接' in line or '来源' in line or '参考' in line):
                in_refs = True
                continue
            if in_refs and line.strip().startswith('•'):
                url_match = _re.search(r'(https?://\S+)', line)
                if url_match:
                    refs.append(url_match.group(1))
            elif in_refs and not line.strip().startswith('•') and line.strip():
                in_refs = False
        # Also find any http URLs
        urls = _re.findall(r'(https?://[^\s\n\)）]+)', text)
        for u in urls:
            if u not in refs:
                refs.append(u)
        return refs

    @staticmethod
    def _parse_other_field_data(all_other_result_texts, field_labels_map):
        """Parse other field data from raw text entries.
        Returns: {field_label: [{value, references}, ...]}
        """
        import re as _re
        result = {}

        for src_info, text in all_other_result_texts:
            # Collect all field matches with positions
            field_matches = []
            for fl_key, fl_label in field_labels_map.items():
                if fl_key == "name":
                    continue
                for m in _re.finditer(
                    rf'(?:^|\n)\s*[✓]\s*{_re.escape(fl_label)}[：:]\s*(.+?)(?:\n|$)',
                    text, _re.MULTILINE
                ):
                    val = m.group(1).strip()
                    if val and val not in ("未找到", "（空）", "not_found"):
                        field_matches.append((m.start(), m.end(), fl_label, val))

            if not field_matches:
                continue

            # Sort by position in text
            field_matches.sort(key=lambda x: x[0])

            for i, (start, end, fl_label, val) in enumerate(field_matches):
                val = _re.sub(r'[，,\s]*[(（]?原值[：:][^，,;；。\n]*(?:[)）])?', '', val).strip()
                val = _re.sub(r'[，,\s]*[(（]?提取自[：:][^，,;；。\n]*(?:[)）])?', '', val).strip()
                val = _re.sub(r'[，,\s]*[(（]?置信度[：:][^，,;；。\n]*(?:[)）])?', '', val).strip()

                # Extract text segment between this field line and the next field line
                if i + 1 < len(field_matches):
                    segment = text[end:field_matches[i + 1][0]]
                else:
                    segment = text[end:]

                # Find URLs only in this field's segment
                refs = _re.findall(r'https?://[^\s\n]+', segment)
                refs = [r.rstrip('.,;:!?）)》〕】') for r in refs]

                if fl_label not in result:
                    result[fl_label] = []
                result[fl_label].append({
                    "value": val,
                    "references": refs,
                })

        return result


def _open_source_edit_dialog(parent, src, raw_case_id, refresh_callback):
    """Popup dialog to edit a case source (URL and credibility)."""
    win = tk.Toplevel(parent)
    win.title("编辑来源")
    win.geometry("550x180")
    win.transient(parent)
    win.resizable(False, False)

    f = ttk.Frame(win, padding=10)
    f.pack(fill="both", expand=True)

    ttk.Label(f, text="URL：").grid(row=0, column=0, sticky="w", pady=5)
    url_var = tk.StringVar(value=src.get("url", ""))
    url_entry = ttk.Entry(f, textvariable=url_var, width=60)
    url_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5)

    ttk.Label(f, text="可信度：").grid(row=1, column=0, sticky="w", pady=5)
    cred_var = tk.StringVar(value=src.get("credibility", "unknown"))
    cred_combo = ttk.Combobox(f, textvariable=cred_var, values=["high", "medium", "low", "unknown"], width=15)
    cred_combo.grid(row=1, column=1, sticky="w", pady=5, padx=5)

    def _save():
        new_url = url_var.get().strip()
        new_cred = cred_var.get().strip()
        if not new_url:
            messagebox.showwarning("提示", "URL 不能为空。")
            return
        try:
            db.update_case_source(src["id"], {"url": new_url, "credibility": new_cred})
            db.insert_verification_log(raw_case_id, "source_edit",
                                       f"编辑来源: url={new_url[:60]}, cred={new_cred}")
            refresh_callback()
            win.destroy()
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    btn_bar = ttk.Frame(f)
    btn_bar.grid(row=2, column=0, columnspan=2, pady=(15, 0))
    ttk.Button(btn_bar, text="保存", command=_save).pack(side="left", padx=5)
    ttk.Button(btn_bar, text="取消", command=win.destroy).pack(side="left", padx=5)


def _open_source_add_dialog(parent, raw_case_id, refresh_callback):
    """Popup dialog to add a new case source."""
    win = tk.Toplevel(parent)
    win.title("添加来源")
    win.geometry("550x180")
    win.transient(parent)
    win.resizable(False, False)

    f = ttk.Frame(win, padding=10)
    f.pack(fill="both", expand=True)

    ttk.Label(f, text="URL：").grid(row=0, column=0, sticky="w", pady=5)
    url_var = tk.StringVar()
    url_entry = ttk.Entry(f, textvariable=url_var, width=60)
    url_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5)

    ttk.Label(f, text="可信度：").grid(row=1, column=0, sticky="w", pady=5)
    cred_var = tk.StringVar(value="unknown")
    cred_combo = ttk.Combobox(f, textvariable=cred_var, values=["high", "medium", "low", "unknown"], width=15)
    cred_combo.grid(row=1, column=1, sticky="w", pady=5, padx=5)

    def _save():
        new_url = url_var.get().strip()
        new_cred = cred_var.get().strip()
        if not new_url:
            messagebox.showwarning("提示", "URL 不能为空。")
            return
        try:
            db.insert_case_source(raw_case_id, new_url, credibility=new_cred)
            db.insert_verification_log(raw_case_id, "source_add",
                                       f"添加来源: url={new_url[:60]}, cred={new_cred}")
            refresh_callback()
            win.destroy()
        except Exception as e:
            messagebox.showerror("添加失败", str(e))

    btn_bar = ttk.Frame(f)
    btn_bar.grid(row=2, column=0, columnspan=2, pady=(15, 0))
    ttk.Button(btn_bar, text="添加", command=_save).pack(side="left", padx=5)
    ttk.Button(btn_bar, text="取消", command=win.destroy).pack(side="left", padx=5)

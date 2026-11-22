import pandas as pd
import os
from tkinter import filedialog, messagebox


def import_excel(db_module):
    filepath = filedialog.askopenfilename(
        title="选择要导入的文件",
        filetypes=[("Excel 文件", "*.xlsx *.xls"), ("CSV 文件", "*.csv"), ("所有文件", "*.*")]
    )
    if not filepath:
        return 0

    try:
        if filepath.endswith(".csv"):
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)
    except Exception as e:
        messagebox.showerror("导入失败", f"无法读取文件：{e}")
        return 0

    if df.empty:
        messagebox.showwarning("导入失败", "文件中没有数据。")
        return 0

    count = db_module.import_from_dataframe(df)
    return count


def export_to_excel(dataframe, title="导出数据"):
    filepath = filedialog.asksaveasfilename(
        title=title,
        defaultextension=".xlsx",
        filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")]
    )
    if not filepath:
        return False

    try:
        if filepath.endswith(".csv"):
            dataframe.to_csv(filepath, index=False, encoding="utf-8-sig")
        else:
            dataframe.to_excel(filepath, index=False, engine="openpyxl")
        return True
    except Exception as e:
        messagebox.showerror("导出失败", f"无法保存文件：{e}")
        return False

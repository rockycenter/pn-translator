#!/usr/bin/env python3
"""
料号翻译工具 - 桌面版
双击运行，拖入周报，一键翻译。无需网络、无需浏览器。
"""

import os
import sys
import re
import shutil
import tempfile
import zipfile
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET

import openpyxl
import xlrd

# ── 内置 Mapping 路径 ─────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).parent
MAPPING_PATH = BASE_DIR / "mapping.xlsx"


def load_mapping(path):
    """读取 mapping，构建双向映射：无论输入新料号还是旧料号，都能翻译。
    重复项只保留第一条记录，避免被错误覆盖。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    mapping = {}  # {part_number: (old, new)}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) < 11:
            continue
        old_v = row[9]   # J 列
        new_v = row[10]  # K 列
        if old_v and new_v:
            old_v = str(old_v).strip()
            new_v = str(new_v).strip()
            if old_v and new_v and new_v.lower() != "nan":
                # 重复 key 只保留第一条（先到先得）
                if old_v not in mapping:
                    mapping[old_v] = (old_v, new_v)
                if new_v not in mapping:
                    mapping[new_v] = (old_v, new_v)
    wb.close()
    return mapping


def is_part_number(value):
    if not value:
        return False
    return bool(re.match(r"^[AB][A-Z0-9]{3,}(\.[ABT])?$", value.strip()))


def translate_xlsx(report_path, output_path, progress_callback=None):
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ET.register_namespace('', NS)
    ET.register_namespace('r',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    if progress_callback:
        progress_callback("加载对照表...")
    mapping = load_mapping(MAPPING_PATH)

    if progress_callback:
        progress_callback("正在翻译...")
    tmpdir = Path(tempfile.mkdtemp(prefix='pn_trans_'))

    try:
        with zipfile.ZipFile(report_path, 'r') as z:
            z.extractall(tmpdir)

        ss_path = tmpdir / 'xl' / 'sharedStrings.xml'
        if not ss_path.exists():
            shutil.copy2(report_path, output_path)
            return {"translated": 0, "unchanged": 0, "not_found": 0}

        tree = ET.parse(ss_path)
        root = tree.getroot()
        stats = {"translated": 0, "unchanged": 0, "not_found": 0}

        for si in root.findall(f'{{{NS}}}si'):
            t_elem = si.find(f'{{{NS}}}t')
            if t_elem is None or t_elem.text is None:
                continue
            original = t_elem.text.strip()
            if not is_part_number(original):
                continue
            if original in mapping:
                old_val, new_val = mapping[original]
                if old_val != new_val:
                    # 统一输出格式：旧料号-新料号
                    t_elem.text = f"{old_val}-{new_val}"
                    t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    stats["translated"] += 1
                else:
                    stats["unchanged"] += 1
            else:
                stats["not_found"] += 1

        tree.write(ss_path, xml_declaration=True, encoding='UTF-8')

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fpath in sorted(tmpdir.rglob('*')):
                if fpath.is_file():
                    arcname = str(fpath.relative_to(tmpdir))
                    zout.write(fpath, arcname)

        return stats
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def translate_xls(report_path, output_path, progress_callback=None):
    """翻译 .xls 文件（Excel 97-2003 格式），输出为 .xlsx"""

    if progress_callback:
        progress_callback("加载对照表...")
    mapping = load_mapping(MAPPING_PATH)

    if progress_callback:
        progress_callback("正在读取文件...")

    # 用 xlrd 读取 .xls
    wb_in = xlrd.open_workbook(report_path)

    # 用 openpyxl 创建输出 .xlsx
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    stats = {"translated": 0, "unchanged": 0, "not_found": 0}
    total_sheets = wb_in.nsheets

    for si in range(total_sheets):
        ws_in = wb_in.sheet_by_index(si)

        if progress_callback:
            progress_callback(f"翻译中... ({si+1}/{total_sheets})")

        ws_out = wb_out.create_sheet(title=ws_in.name[:31])  # Excel 限制 31 字符

        for r in range(ws_in.nrows):
            for c in range(ws_in.ncols):
                cell_value = ws_in.cell_value(r, c)
                cell_type = ws_in.cell_type(r, c)

                if cell_type == xlrd.XL_CELL_TEXT and cell_value:
                    original = str(cell_value).strip()
                    if is_part_number(original) and original in mapping:
                        old_val, new_val = mapping[original]
                        if old_val != new_val:
                            ws_out.cell(row=r+1, column=c+1).value = f"{old_val}-{new_val}"
                            stats["translated"] += 1
                        else:
                            ws_out.cell(row=r+1, column=c+1).value = original
                            stats["unchanged"] += 1
                    else:
                        ws_out.cell(row=r+1, column=c+1).value = original
                elif cell_type == xlrd.XL_CELL_NUMBER:
                    ws_out.cell(row=r+1, column=c+1).value = cell_value
                elif cell_type == xlrd.XL_CELL_DATE:
                    import datetime
                    dt = xlrd.xldate_as_datetime(cell_value, wb_in.datemode)
                    ws_out.cell(row=r+1, column=c+1).value = dt
                elif cell_type == xlrd.XL_CELL_BOOLEAN:
                    ws_out.cell(row=r+1, column=c+1).value = bool(cell_value)
                elif cell_type == xlrd.XL_CELL_EMPTY:
                    pass  # 空单元格，不写入
                else:
                    # 其他类型转字符串
                    ws_out.cell(row=r+1, column=c+1).value = str(cell_value) if cell_value else None

    wb_out.save(output_path)
    return stats


# ── GUI ────────────────────────────────────────────────────────────
class TranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("料号翻译工具")
        self.root.geometry("520x450")
        self.root.resizable(False, False)

        # 设置窗口图标
        icon_path = BASE_DIR / "icons" / "icon.ico"
        try:
            self.root.iconbitmap(default=str(icon_path))
        except Exception:
            pass

        self.root.update_idletasks()
        w, h = 520, 450
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        self.bg = "#f5f5f7"
        self.card = "#ffffff"
        self.accent = "#0071e3"
        self.text = "#1d1d1f"
        self.sub = "#6e6e73"

        self.root.configure(bg=self.bg)
        self.file_path = None
        self._build_ui()

    def _build_ui(self):
        title = tk.Label(self.root, text="📦 料号翻译工具",
                         font=("Microsoft YaHei UI", 20, "bold"),
                         fg=self.text, bg=self.bg)
        title.pack(pady=(30, 4))

        version = tk.Label(self.root, text="V1.2",
                           font=("Microsoft YaHei UI", 9),
                           fg=self.sub, bg=self.bg)
        version.pack(pady=(0, 6))

        subtitle = tk.Label(self.root,
                            text="拖入或选择你需要的文件，一键翻译为【旧料号-新料号】格式\n靶向治疗，拒绝内耗",
                            font=("Microsoft YaHei UI", 11),
                            fg=self.sub, bg=self.bg)
        subtitle.pack(pady=(0, 20))

        card = tk.Frame(self.root, bg=self.card, bd=0,
                        highlightbackground="#d2d2d7",
                        highlightthickness=1, padx=24, pady=24)
        card.pack(fill="both", padx=30, pady=(0, 10))

        self.drop_frame = tk.Frame(card, bg=self.card, bd=2, relief="groove")
        self.drop_frame.pack(fill="both", expand=True, ipady=30)

        self.drop_label = tk.Label(self.drop_frame,
                                   text="点击选择你需要的文件",
                                   font=("Microsoft YaHei UI", 12),
                                   fg=self.sub, bg=self.card, justify="center")
        self.drop_label.pack(expand=True)

        self.file_label = tk.Label(self.drop_frame, text="",
                                   font=("Microsoft YaHei UI", 11, "bold"),
                                   fg=self.accent, bg=self.card)
        self.file_label.pack()

        self.drop_frame.bind("<Button-1>", lambda e: self._select_file())
        self.drop_label.bind("<Button-1>", lambda e: self._select_file())
        # 拖拽功能仅 macOS 支持，Windows 自动降级为按钮选择
        pass

        btn_frame = tk.Frame(card, bg=self.card)
        btn_frame.pack(fill="x", pady=(16, 0))

        self.select_btn = tk.Button(btn_frame, text="选择文件",
                                    font=("Microsoft YaHei UI", 11),
                                    bg="#e8e8ed", fg=self.text,
                                    relief="flat", padx=20, pady=6,
                                    command=self._select_file)
        self.select_btn.pack(side="left")

        # 蓝底白字按钮
        self.translate_btn = tk.Button(btn_frame, text="开始翻译",
                                       font=("Microsoft YaHei UI", 12, "bold"),
                                       bg=self.accent, fg="white",
                                       activebackground="#0062c4",
                                       activeforeground="white",
                                       disabledforeground="white",
                                       relief="flat", padx=24, pady=8,
                                       borderwidth=0,
                                       highlightthickness=0,
                                       state="disabled",
                                       command=self._translate)
        self.translate_btn.pack(side="right")

        # 署名行
        signature = tk.Label(card,
                            text="ROCKYCENTER PRODUCTION",
                            font=("Microsoft YaHei UI", 8),
                            fg="#c7c7cc", bg=self.card,
                            justify="center")
        signature.pack(side="bottom", pady=(12, 0))

        self.progress = ttk.Progressbar(card, mode="indeterminate", length=400)

        self.status_label = tk.Label(card, text="",
                                     font=("Microsoft YaHei UI", 10),
                                     fg=self.sub, bg=self.card)

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="选择周报文件",
            filetypes=[("Excel 文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.file_path = path
            self.file_label.config(text=f"✅  {os.path.basename(path)}")
            self.translate_btn.config(state="normal")

    def _on_drop(self, event):
        data = event.data
        if data:
            path = data.strip().strip('{').strip('}')
            if os.path.isfile(path) and (path.lower().endswith('.xlsx') or path.lower().endswith('.xls')):
                self.file_path = path
                self.file_label.config(text=f"✅  {os.path.basename(path)}")
                self.translate_btn.config(state="normal")

    def _translate(self):
        if not self.file_path:
            return
                if self.file_path.lower().endswith('.xls') and not self.file_path.lower().endswith('.xlsx'):
            base = self.file_path.rsplit('.', 1)[0]
            default_out = base + '_translated.xlsx'
        else:
            default_out = self.file_path.replace('.xlsx', '_translated.xlsx')
        output_path = filedialog.asksaveasfilename(
            title="保存翻译结果",
            defaultextension=".xlsx",
            initialfile=os.path.basename(default_out),
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not output_path:
            return

        self.translate_btn.config(state="disabled", text="翻译中...")
        self.select_btn.config(state="disabled")
        self.progress.pack(pady=(12, 0))
        self.progress.start()

        def done(stats):
            self.root.after(0, lambda: self._on_done(stats, output_path))

        def run():
            try:
                is_xls = self.file_path.lower().endswith('.xls') and not self.file_path.lower().endswith('.xlsx')
                if is_xls:
                    stats = translate_xls(self.file_path, output_path, progress_callback=lambda msg: None)
                else:
                    stats = translate_xlsx(self.file_path, output_path)
                done(stats)
            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=run, daemon=True).start()

    def _on_done(self, stats, output_path):
        self.progress.stop()
        self.progress.pack_forget()
        self.translate_btn.config(state="normal", text="开始翻译")
        self.select_btn.config(state="normal")

        t = stats.get("translated", 0)
        nf = stats.get("not_found", 0)

        msg = f"翻译完成！\n\n✅ 已翻译：{t} 个\n"
        if nf > 0:
            msg += f"⚠️ 未找到映射：{nf} 个（其他辅助 Sheet 中的旧编号，不影响阅读）\n"
        msg += f"\n文件已保存至：\n{output_path}"
        messagebox.showinfo("翻译完成", msg)

    def _on_error(self, err_msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.translate_btn.config(state="normal", text="开始翻译")
        self.select_btn.config(state="normal")
        messagebox.showerror("翻译失败", f"发生错误：\n{err_msg}")


def main():
    root = tk.Tk()
    TranslatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

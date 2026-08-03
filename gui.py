#!/usr/bin/env python3
"""
料号翻译工具 - 桌面版（tkinter GUI）
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

import pandas as pd

# ── 内置 Mapping 路径 ─────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).parent
MAPPING_PATH = BASE_DIR / "mapping.xlsx"

# ── 核心翻译逻辑 ──────────────────────────────────────────────────
def load_mapping(path):
    df = pd.read_excel(path, header=None, skiprows=2)
    df.columns = range(df.shape[1])
    mapping = df[[9, 10]].dropna()
    new_to_old = {}
    for _, row in mapping.iterrows():
        old_v = str(row[9]).strip()
        new_v = str(row[10]).strip()
        if new_v.lower() != "nan" and old_v.lower() != "nan":
            new_to_old[new_v] = old_v
    return new_to_old

def is_part_number(value):
    if not value: return False
    return bool(re.match(r"^[AB][A-Z0-9]{3,}(\.[ABT])?$", value.strip()))

def translate_xlsx(report_path, output_path, progress_callback=None):
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ET.register_namespace('', NS)
    ET.register_namespace('r',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    if progress_callback:
        progress_callback("加载对照表...")
    new_to_old = load_mapping(MAPPING_PATH)

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
        total = 0

        for si in root.findall(f'{{{NS}}}si'):
            t_elem = si.find(f'{{{NS}}}t')
            if t_elem is None or t_elem.text is None:
                continue
            original = t_elem.text.strip()
            if not is_part_number(original):
                continue
            total += 1
            if original in new_to_old:
                old_val = new_to_old[original]
                if old_val != original:
                    t_elem.text = f"{old_val}-{original}"
                    t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    stats["translated"] += 1
                else:
                    stats["unchanged"] += 1
            else:
                stats["not_found"] += 1

            if progress_callback and total % 50 == 0:
                progress_callback(f"翻译中... ({total})")

        tree.write(ss_path, xml_declaration=True, encoding='UTF-8')

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fpath in sorted(tmpdir.rglob('*')):
                if fpath.is_file():
                    arcname = str(fpath.relative_to(tmpdir))
                    zout.write(fpath, arcname)

        return stats
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ── GUI ────────────────────────────────────────────────────────────
class TranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("料号翻译工具")
        self.root.geometry("520x420")
        self.root.resizable(False, False)

        # 居中
        self.root.update_idletasks()
        w, h = 520, 420
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        # 样式
        self.bg = "#f5f5f7"
        self.card = "#ffffff"
        self.accent = "#0071e3"
        self.text = "#1d1d1f"
        self.sub = "#6e6e73"
        self.green = "#34c759"
        self.border = "#d2d2d7"

        self.root.configure(bg=self.bg)
        self.file_path = None

        self._build_ui()

    def _build_ui(self):
        # 标题
        title = tk.Label(self.root, text="📦 料号翻译工具",
                         font=("Microsoft YaHei UI", 20, "bold"),
                         fg=self.text, bg=self.bg)
        title.pack(pady=(30, 4))

        subtitle = tk.Label(self.root,
                            text="拖入或选择周报文件，一键翻译为「旧料号-新料号」格式",
                            font=("Microsoft YaHei UI", 11),
                            fg=self.sub, bg=self.bg)
        subtitle.pack(pady=(0, 20))

        # 卡片区域
        card = tk.Frame(self.root, bg=self.card, bd=0,
                        highlightbackground=self.border,
                        highlightthickness=1,
                        padx=24, pady=24)
        card.pack(fill="both", padx=30, pady=(0, 10))

        # 拖放区域
        self.drop_frame = tk.Frame(card, bg=self.card, bd=2, relief="groove")
        self.drop_frame.pack(fill="both", expand=True, ipady=30)

        self.drop_label = tk.Label(self.drop_frame,
                                   text="📊  拖拽周报文件到此处\n或点击下方按钮选择",
                                   font=("Microsoft YaHei UI", 12),
                                   fg=self.sub, bg=self.card,
                                   justify="center")
        self.drop_label.pack(expand=True)

        self.file_label = tk.Label(self.drop_frame, text="",
                                   font=("Microsoft YaHei UI", 11, "bold"),
                                   fg=self.accent, bg=self.card)
        self.file_label.pack()

        # 拖放绑定
        self.drop_frame.bind("<Button-1>", lambda e: self._select_file())
        self.drop_label.bind("<Button-1>", lambda e: self._select_file())
        self.root.drop_target_register("*")
        self.root.dnd_bind('<<Drop>>', self._on_drop)

        # 按钮
        btn_frame = tk.Frame(card, bg=self.card)
        btn_frame.pack(fill="x", pady=(16, 0))

        self.select_btn = tk.Button(btn_frame, text="选择文件",
                                    font=("Microsoft YaHei UI", 11),
                                    bg="#e8e8ed", fg=self.text,
                                    activebackground="#d2d2d7",
                                    relief="flat", padx=20, pady=6,
                                    command=self._select_file)
        self.select_btn.pack(side="left")

        self.translate_btn = tk.Button(btn_frame, text="开始翻译",
                                       font=("Microsoft YaHei UI", 12, "bold"),
                                       bg=self.accent, fg="white",
                                       activebackground="#0062c4",
                                       relief="flat", padx=24, pady=8,
                                       state="disabled",
                                       command=self._translate)
        self.translate_btn.pack(side="right")

        # 进度条
        self.progress = ttk.Progressbar(card, mode="indeterminate", length=400)
        self.progress.pack(pady=(12, 0))
        self.progress.pack_forget()

        self.status_label = tk.Label(card, text="",
                                     font=("Microsoft YaHei UI", 10),
                                     fg=self.sub, bg=self.card)
        self.status_label.pack(pady=(4, 0))

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="选择周报文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.file_path = path
            self.file_label.config(text=f"✅  {os.path.basename(path)}")
            self.translate_btn.config(state="normal")

    def _on_drop(self, event):
        # 解析拖放的文件路径
        data = event.data
        if data:
            # Windows 拖放格式: {path} 或带空格的路径
            path = data.strip().strip('{').strip('}')
            if os.path.isfile(path) and path.endswith('.xlsx'):
                self.file_path = path
                self.file_label.config(text=f"✅  {os.path.basename(path)}")
                self.translate_btn.config(state="normal")
            else:
                messagebox.showwarning("不支持", "请拖入 .xlsx 格式的周报文件")

    def _translate(self):
        if not self.file_path:
            return

        # 选择输出路径
        default_out = self.file_path.replace('.xlsx', '_translated.xlsx')
        output_path = filedialog.asksaveasfilename(
            title="保存翻译结果",
            defaultextension=".xlsx",
            initialfile=os.path.basename(default_out),
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not output_path:
            return

        # 开始翻译
        self.translate_btn.config(state="disabled", text="翻译中...")
        self.select_btn.config(state="disabled")
        self.progress.pack(pady=(12, 0))
        self.progress.start()

        def update_status(msg):
            self.root.after(0, lambda: self.status_label.config(text=msg))

        def done(stats):
            self.root.after(0, lambda: self._on_done(stats, output_path))

        def run():
            try:
                stats = translate_xlsx(self.file_path, output_path,
                                       progress_callback=update_status)
                done(stats)
            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=run, daemon=True).start()

    def _on_done(self, stats, output_path):
        self.progress.stop()
        self.progress.pack_forget()
        self.translate_btn.config(state="normal", text="开始翻译")
        self.select_btn.config(state="normal")
        self.status_label.config(text="")

        t = stats.get("translated", 0)
        u = stats.get("unchanged", 0)
        nf = stats.get("not_found", 0)

        msg = f"翻译完成！\n\n"
        msg += f"✅ 已翻译：{t} 个\n"
        if u > 0:
            msg += f"➖ 无需翻译：{u} 个\n"
        if nf > 0:
            msg += f"⚠️ 未找到映射：{nf} 个（其他辅助 Sheet 中的旧编号，不影响阅读）\n"
        msg += f"\n文件已保存至：\n{output_path}"

        messagebox.showinfo("翻译完成", msg)

    def _on_error(self, err_msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.translate_btn.config(state="normal", text="开始翻译")
        self.select_btn.config(state="normal")
        self.status_label.config(text="")
        messagebox.showerror("翻译失败", f"发生错误：\n{err_msg}")


# ── 入口 ──────────────────────────────────────────────────────────
def main():
    root = tk.Tk()
    app = TranslatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
按 Sheet 名合并多个 Excel —— 可选字段 · 保留原格式

由独立脚本 DuoExcelUI2.py 移植为工具箱功能模块。
合并逻辑（do_merge 及以下工具函数）未作任何改动，只把主窗口改成了可嵌入的面板。

【合并方案】
  · 以「第一个含该 sheet 的文件」为格式母版，整份保留其格式
  · 其他文件的同名 sheet 逐行追加；列按【表头名】对应，不新增列、不丢列
  · 追加行样式从母版的数据行复制 → 字体/边框/数字格式不丢
  · 可勾选只保留哪些字段；不删空行空列；源文件不改动

依赖：只用 openpyxl（不需要 pandas / numpy）
"""
import os
import shutil
import copy as _copy
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

try:
    from .base import ToolFrame
except ImportError:                      # 允许直接运行本文件调试：python tools/duo_excel_ui.py
    import os as _os
    import sys as _sys
    _sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
    from tools.base import ToolFrame     # noqa: F401

# ============ 可调参数 ============
HEADER_ROW = 1                    # 表头所在行（1 起）
EXCEL_EXT = ('.xlsx', '.xlsm')    # 支持的扩展名（.xls 老格式 openpyxl 读不了）
# =================================


def _s(v):
    """单元格值 → 去空白的字符串"""
    return '' if v is None else str(v).strip()


def list_excel_files(folder):
    """列出文件夹内的 Excel 文件；跳过 ~$ 临时文件与隐藏文件"""
    out = []
    for fn in sorted(os.listdir(folder)):
        if fn.startswith('~$') or fn.startswith('.'):
            continue                      # ← 旧版这里写反了，会把 ~$ 临时文件当数据读
        if fn.lower().endswith(EXCEL_EXT):
            out.append(os.path.join(folder, fn))
    return out


def scan_folder(folder, on_log=None):
    """
    扫描文件夹，返回 (文件列表, sheet 信息字典)

    sheet 信息字典结构：
        { sheet名: { 'files': [文件路径...],   # 含该 sheet 的文件（按文件名排序）
                     'base':  格式母版文件,      # 第一个含该 sheet 的文件
                     'columns': [表头名...] } }  # 母版的列（作为字段勾选依据）
    """
    def log(m):
        if on_log:
            on_log(m)

    files = list_excel_files(folder)
    log('找到 %d 个 Excel 文件' % len(files))

    sheets = {}
    for p in files:
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
        except Exception as e:
            log('  [跳过] %s 打不开：%s' % (os.path.basename(p), e))
            continue
        try:
            names = list(wb.sheetnames)
            for sn in names:
                ws = wb[sn]
                hdr = []
                for row in ws.iter_rows(min_row=HEADER_ROW, max_row=HEADER_ROW, values_only=True):
                    hdr = [_s(v) for v in row]
                    break
                if not any(hdr):
                    log('  [跳过] %s / %s：表头为空' % (os.path.basename(p), sn))
                    continue
                d = sheets.setdefault(sn, {'files': [], 'base': p, 'columns': None})
                d['files'].append(p)
                if d['columns'] is None:
                    d['columns'] = hdr
                    d['base'] = p
        finally:
            wb.close()

    for sn, d in sheets.items():
        log('  工作表【%s】: %d 个文件, %d 个字段' % (sn, len(d['files']), len(d['columns'])))
    return files, sheets


# ══════════════════════════════════════════════════════════════
#  二、样式工具
# ══════════════════════════════════════════════════════════════
def copy_cell_style(src, dst):
    """把 src 单元格的样式复制到 dst（含数字格式）"""
    dst.font = _copy.copy(src.font)
    dst.border = _copy.copy(src.border)
    dst.fill = _copy.copy(src.fill)
    dst.alignment = _copy.copy(src.alignment)
    dst.protection = _copy.copy(src.protection)
    dst.number_format = src.number_format


def snapshot_style(src):
    """把单元格样式取成快照（脱离原工作簿也能复用）"""
    return {
        'font': _copy.copy(src.font),
        'border': _copy.copy(src.border),
        'fill': _copy.copy(src.fill),
        'alignment': _copy.copy(src.alignment),
        'protection': _copy.copy(src.protection),
        'number_format': src.number_format,
    }


def apply_snapshot(dst, snap):
    """把样式快照套到目标单元格"""
    dst.font = snap['font']
    dst.border = snap['border']
    dst.fill = snap['fill']
    dst.alignment = snap['alignment']
    dst.protection = snap['protection']
    dst.number_format = snap['number_format']


def copy_sheet_layout(src_path, sheet_name, ws_out, keep_idx):
    """把模板 sheet 的版式复制到新 sheet：列宽 + 冻结窗格（keep_idx 为保留列的 0 基索引）"""
    wb = load_workbook(src_path, data_only=False)
    try:
        ws = wb[sheet_name]
        for pos, src_i in enumerate(keep_idx):
            src_letter = get_column_letter(src_i + 1)
            dst_letter = get_column_letter(pos + 1)
            dim = ws.column_dimensions.get(src_letter)
            if dim is not None and dim.width:
                ws_out.column_dimensions[dst_letter].width = dim.width
            elif dim is not None and dim.hidden:
                ws_out.column_dimensions[dst_letter].hidden = True
        try:
            if ws.freeze_panes:
                ws_out.freeze_panes = ws.freeze_panes
        except Exception:
            pass
    finally:
        wb.close()


def read_sheet_rows(path, sheet_name, keep_names, src_col_index):
    """
    读一个源 sheet 的数据行（表头行之后），只取 keep_names 指定的列。
    返回 [[值...], ...]，每行长度 = len(keep_names)；整行为空则跳过。
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return []
        ws = wb[sheet_name]
        out = []
        for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
            if row is None:
                continue
            vals = []
            for nm in keep_names:
                i = src_col_index.get(nm)
                vals.append(row[i] if (i is not None and i < len(row)) else None)
            if any(v is not None and _s(v) != '' for v in vals):
                out.append(vals)
        return out
    finally:
        wb.close()


# ══════════════════════════════════════════════════════════════
#  三、合并主流程
# ══════════════════════════════════════════════════════════════
def do_merge(folder, save_path, sheets, selection, log):
    """
    sheets    : scan_folder 返回的 sheet 信息
    selection : { sheet名: [要保留的列名...] }   （按母版列顺序给出；空列表=跳过该表）
    """
    # 选中的 sheet（字段不为空）
    targets = {sn: cols for sn, cols in selection.items() if cols}
    if not targets:
        log('没有勾选任何字段，未执行合并')
        return False

    # 选一个「输出母版文件」：包含最多目标 sheet 的那个文件
    files = []
    for sn in targets:
        for p in sheets[sn]['files']:
            if p not in files:
                files.append(p)
    score = {}
    for p in files:
        try:
            wb = load_workbook(p, read_only=True)
            owned = sum(1 for sn in targets if sn in wb.sheetnames)
            wb.close()
        except Exception:
            owned = -1
        score[p] = owned
    out_base = max(score, key=lambda k: (score[k], -files.index(k)))
    log('输出母版文件：%s（覆盖 %d 个目标工作表）' % (os.path.basename(out_base), score[out_base]))

    # 1) 复制母版文件 → 输出文件（完整保留其所有格式）
    if os.path.abspath(out_base) == os.path.abspath(save_path):
        log('[X] 输出文件不能覆盖母版文件，请换个保存路径')
        return False
    if os.path.exists(save_path):
        try:
            os.remove(save_path)
        except Exception:
            pass
    shutil.copyfile(out_base, save_path)

    wb_out = load_workbook(save_path)
    total_rows = 0

    for sn, keep_cols in targets.items():
        info = sheets[sn]
        all_cols = info['columns']

        # 母版列次序里保留的位置（0 基）
        keep_set = set(keep_cols)
        keep_idx = [i for i, c in enumerate(all_cols) if c in keep_set]
        if not keep_idx:
            continue
        keep_names = [all_cols[i] for i in keep_idx]
        ncol = len(keep_names)

        native = sn in wb_out.sheetnames          # 输出母版文件里就有这个表 → 格式原样保留
        if native:
            ws = wb_out[sn]
            tpl_path = out_base
            # 删掉未勾选的列（从右往左删，避免索引错位）
            drop = [i + 1 for i in range(len(all_cols)) if i not in keep_idx]
            for ci in sorted(drop, reverse=True):
                if ci <= ws.max_column:
                    ws.delete_cols(ci)
            header_row = HEADER_ROW
            # 母版自身的数据行直接保留 → 只追加其它文件
            append_files = [p for p in info['files'] if os.path.abspath(p) != os.path.abspath(out_base)]
        else:
            ws = wb_out.create_sheet(sn)
            tpl_path = info['base']
            copy_sheet_layout(tpl_path, sn, ws, keep_idx)
            # 写表头：样式从模板表头单元格复制
            wb_tpl = load_workbook(tpl_path, data_only=False)
            try:
                ws_tpl = wb_tpl[sn]
                for pos, src_i in enumerate(keep_idx):
                    src_cell = ws_tpl.cell(row=HEADER_ROW, column=src_i + 1)
                    dst_cell = ws.cell(row=HEADER_ROW, column=pos + 1)
                    dst_cell.value = all_cols[src_i]
                    copy_cell_style(src_cell, dst_cell)
            finally:
                wb_tpl.close()
            header_row = HEADER_ROW
            append_files = list(info['files'])        # 该表不在输出母版里 → 所有文件都要写

        # 追加行的样式来源 = 模板表的「第一条数据行」(row = HEADER_ROW+1)，逐列取。
        # 注意：新建表那条分支此时输出表里还没有数据行，所以必须从【模板文件】读取，
        #       不能从输出表读（否则追加行会丢掉边框/数字格式）。
        style_src = {}
        row_height = None
        try:
            wb_st = load_workbook(tpl_path, data_only=False)
            ws_st = wb_st[sn]
            if ws_st.max_row >= HEADER_ROW + 1:
                for pos, src_i in enumerate(keep_idx):
                    style_src[pos + 1] = snapshot_style(
                        ws_st.cell(row=HEADER_ROW + 1, column=src_i + 1))
                rd = ws_st.row_dimensions.get(HEADER_ROW + 1)
                if rd is not None:
                    row_height = rd.height
            else:
                log('  [!] 模板表【%s】只有表头、没有数据行，追加行将使用默认样式' % sn)
            wb_st.close()
        except Exception as e:
            log('  [!] 读取模板样式失败，追加行用默认样式：%s' % e)

        added = 0
        warned = set()
        for p in append_files:
            # 取源文件的表头 → 列名到索引的映射（重名列只取第一次出现）
            try:
                wb_s = load_workbook(p, read_only=True, data_only=True)
                if sn not in wb_s.sheetnames:
                    wb_s.close()
                    continue
                ws_s = wb_s[sn]
                raw_hdr = []
                for row in ws_s.iter_rows(min_row=HEADER_ROW, max_row=HEADER_ROW, values_only=True):
                    raw_hdr = [_s(v) for v in row]
                    break
                wb_s.close()
            except Exception as e:
                log('  [跳过] %s / %s 读取失败：%s' % (os.path.basename(p), sn, e))
                continue

            src_index = {}
            for i, h in enumerate(raw_hdr):
                if h and h not in src_index:
                    src_index[h] = i

            missing = [nm for nm in keep_names if nm not in src_index]
            for nm in missing:
                if nm not in warned:
                    warned.add(nm)
                    log('  [!] %s 里缺少字段「%s」，该列留空' % (os.path.basename(p), nm))

            rows = read_sheet_rows(p, sn, keep_names, src_index)
            if not rows:
                continue
            r0 = ws.max_row + 1
            for ri, vals in enumerate(rows):
                rr = r0 + ri
                for ci, v in enumerate(vals, start=1):
                    cell = ws.cell(row=rr, column=ci)
                    cell.value = v
                    if ci in style_src:
                        apply_snapshot(cell, style_src[ci])
                if row_height:
                    ws.row_dimensions[rr].height = row_height
            added += len(rows)
            log('  + %s：追加 %d 行' % (os.path.basename(p), len(rows)))

        total_rows += added
        log('工作表【%s】完成：保留 %d 个字段，追加 %d 行' % (sn, ncol, added))

    # 删掉母版里没用到的多余空表（例如母版文件有导出模板页）
    keep_sheets = set(targets.keys())
    for sn in list(wb_out.sheetnames):
        if sn not in keep_sheets:
            try:
                del wb_out[sn]
                log('删除多余工作表：%s' % sn)
            except Exception:
                pass

    wb_out.save(save_path)
    log('全部完成：共追加 %d 行，已保存到 %s' % (total_rows, save_path))
    return True


# ══════════════════════════════════════════════════════════════
#  四、界面
# ══════════════════════════════════════════════════════════════
class DuoExcelUI(ToolFrame):
    NAME = "按Sheet名合并"
    DESC = "合并多个 Excel 中同名工作表，可选字段、保留原格式"
    def __init__(self, master):
        super().__init__(master)

        self.folder_var = tk.StringVar()
        self.save_var = tk.StringVar()
        self.sheets = {}          # scan 结果
        self.col_vars = {}        # {sheet: {列名: BooleanVar}}
        self.cur_sheet = None     # 当前显示的工作表

        self._build_top()
        self._build_mid()
        self._build_bottom()

    # ---------- 顶部：选文件夹 ----------
    def _build_top(self):
        f = tk.LabelFrame(self, text='1. 选择包含待合并 Excel 的文件夹')
        f.pack(fill=tk.X, padx=10, pady=(10, 4))
        tk.Entry(f, textvariable=self.folder_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=6)
        tk.Button(f, text='浏览', width=7, command=self.pick_folder).pack(side=tk.LEFT, padx=4)
        tk.Button(f, text='扫描', width=7, bg='#2E86AB', fg='white', command=self.do_scan).pack(side=tk.LEFT, padx=(0, 6))

    # ---------- 中部：左表名 + 右字段 ----------
    def _build_mid(self):
        mid = tk.Frame(self)
        mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        left = tk.LabelFrame(mid, text='2. 工作表（点选查看字段）')
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))
        self.lb = tk.Listbox(left, width=22, height=12, exportselection=False)
        self.lb.pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=6)
        sb = tk.Scrollbar(left, orient='vertical', command=self.lb.yview)
        sb.pack(side=tk.LEFT, fill=tk.Y, pady=6)
        self.lb.config(yscrollcommand=sb.set)
        self.lb.bind('<<ListboxSelect>>', lambda e: self.on_pick_sheet())

        right = tk.LabelFrame(mid, text='3. 勾选要合并的字段（取消全部勾选 = 跳过该表）')
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 可滚动的勾选区
        wrap = tk.Frame(right)
        wrap.pack(fill=tk.BOTH, expand=True, padx=6, pady=(6, 2))
        self.canvas = tk.Canvas(wrap, highlightthickness=0)
        vsb = tk.Scrollbar(wrap, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.col_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.col_frame, anchor='nw')
        self.col_frame.bind('<Configure>',
                            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        # 滚轮只在鼠标位于本区域时生效（不用 bind_all，否则会抢走其他面板的滚轮）
        self.canvas.bind('<Enter>', lambda e: self.canvas.bind_all('<MouseWheel>', self._on_wheel))
        self.canvas.bind('<Leave>', lambda e: self.canvas.unbind_all('<MouseWheel>'))

        bar = tk.Frame(right)
        bar.pack(fill=tk.X, padx=6, pady=(0, 6))
        tk.Button(bar, text='全选', width=8, command=lambda: self.set_all(True)).pack(side=tk.LEFT, padx=3)
        tk.Button(bar, text='全不选', width=8, command=lambda: self.set_all(False)).pack(side=tk.LEFT, padx=3)
        tk.Button(bar, text='反选', width=8, command=self.invert).pack(side=tk.LEFT, padx=3)
        tk.Button(bar, text='应用到所有工作表', command=self.apply_to_all).pack(side=tk.LEFT, padx=12)

    # ---------- 底部：保存 + 按钮 + 日志 ----------
    def _build_bottom(self):
        f = tk.LabelFrame(self, text='4. 保存合并结果')
        f.pack(fill=tk.X, padx=10, pady=4)
        tk.Entry(f, textvariable=self.save_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=6)
        tk.Button(f, text='另存为', width=7, command=self.pick_save).pack(side=tk.LEFT, padx=4)
        tk.Button(f, text='▶ 开始合并', width=12, bg='#E67E22', fg='white',
                  font=('微软雅黑', 10, 'bold'), command=self.run).pack(side=tk.LEFT, padx=(0, 6))

        lf = tk.LabelFrame(self, text='运行日志')
        lf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))
        self.log_text = tk.Text(lf, height=7, wrap='word')
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0), pady=6)
        lsb = tk.Scrollbar(lf, orient='vertical', command=self.log_text.yview)
        lsb.pack(side=tk.LEFT, fill=tk.Y, pady=6, padx=(0, 6))
        self.log_text.config(yscrollcommand=lsb.set)
        self.log('就绪。请选择文件夹后点「扫描」。')

    def _on_wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 120), 'units')

    # ---------- 日志 ----------
    def log(self, msg):
        self.log_text.insert(tk.END, str(msg) + '\n')
        self.log_text.see(tk.END)
        self.update_idletasks()

    # ---------- 事件 ----------
    def pick_folder(self):
        p = filedialog.askdirectory(title='选择包含待合并 Excel 的文件夹')
        if p:
            self.folder_var.set(p)

    def pick_save(self):
        p = filedialog.asksaveasfilename(title='保存合并结果',
                                         defaultextension='.xlsx',
                                         filetypes=[('Excel 文件', '*.xlsx')],
                                         initialfile='合并结果.xlsx')
        if p:
            self.save_var.set(p)

    def do_scan(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning('提示', '请先选择有效的文件夹')
            return
        self.log_text.delete('1.0', tk.END)
        self.sheets = {}
        self.col_vars = {}
        self.lb.delete(0, tk.END)
        for w in self.col_frame.winfo_children():
            w.destroy()

        files, sheets = scan_folder(folder, self.log)
        self.sheets = sheets
        if not sheets:
            self.log('没有扫描到任何工作表。')
            messagebox.showwarning('提示', '没有扫描到可用工作表（检查文件是否为 .xlsx/.xlsm）')
            return
        for i, sn in enumerate(sheets):
            n = len(sheets[sn]['files'])
            self.lb.insert(tk.END, '%s  (%d文件)' % (sn, n))
            self.col_vars[sn] = {c: tk.BooleanVar(value=True) for c in sheets[sn]['columns']}
        self.lb.selection_set(0)
        self.on_pick_sheet()
        self.log('扫描完成：%d 个工作表名。默认全选字段，按需取消。' % len(sheets))

    def on_pick_sheet(self):
        sel = self.lb.curselection()
        if not sel:
            return
        sn = list(self.sheets.keys())[sel[0]]
        self.cur_sheet = sn
        for w in self.col_frame.winfo_children():
            w.destroy()
        cols = self.sheets[sn]['columns']
        self.log('当前工作表【%s】共 %d 个字段' % (sn, len(cols)))
        for i, c in enumerate(cols):
            cb = tk.Checkbutton(self.col_frame, text=c, variable=self.col_vars[sn][c],
                                anchor='w', width=30)
            cb.grid(row=i // 3, column=i % 3, sticky='w', padx=6, pady=2)
        self.canvas.yview_moveto(0)

    def set_all(self, val):
        if not self.cur_sheet:
            return
        for v in self.col_vars[self.cur_sheet].values():
            v.set(val)

    def invert(self):
        if not self.cur_sheet:
            return
        for v in self.col_vars[self.cur_sheet].values():
            v.set(not v.get())

    def apply_to_all(self):
        """把当前表的勾选情况按字段名套用到其他工作表"""
        if not self.cur_sheet:
            return
        cur = {c: v.get() for c, v in self.col_vars[self.cur_sheet].items()}
        n = 0
        for sn, d in self.col_vars.items():
            if sn == self.cur_sheet:
                continue
            for c, v in d.items():
                if c in cur:
                    v.set(cur[c])
            n += 1
        self.log('已把当前字段勾选按【字段名】套用到另外 %d 个工作表' % n)

    def run(self):
        folder = self.folder_var.get().strip()
        save = self.save_var.get().strip()
        if not self.sheets:
            messagebox.showwarning('提示', '请先点「扫描」')
            return
        if not save:
            messagebox.showwarning('提示', '请先指定保存路径')
            return
        selection = {}
        for sn, d in self.col_vars.items():
            cols = [c for c in self.sheets[sn]['columns'] if d[c].get()]
            if cols:
                selection[sn] = cols
        if not selection:
            messagebox.showwarning('提示', '至少要在一个工作表里勾选字段')
            return
        self.log('─' * 60)
        ok = do_merge(folder, save, self.sheets, selection, self.log)
        if ok:
            messagebox.showinfo('完成', '合并完成！\n保存到：\n%s' % save)
        else:
            messagebox.showerror('失败', '合并未完成，请看下方日志')


if __name__ == '__main__':
    # 单独调试用：在本目录的上一级（ExcelToolbox 根目录）执行
    #     python -m tools.duo_excel_ui
    import sys
    _root = tk.Tk()
    _root.title('按 Sheet 名合并 Excel（单独调试窗口）')
    _root.geometry('1020x760')
    DuoExcelUI(_root).pack(fill=tk.BOTH, expand=True)
    if len(sys.argv) > 1 and sys.argv[1] == '--smoke':
        _root.update()          # 只构建界面，立刻退出（用于自动化自检）
        print('SMOKE OK: %s' % DuoExcelUI.NAME)
        _root.destroy()
    else:
        _root.mainloop()

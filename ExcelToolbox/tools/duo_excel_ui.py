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
import queue
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.formatting.formatting import ConditionalFormatting, ConditionalFormattingList

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
    log('找到 %d 个 Excel 文件：' % len(files))
    for p in files:
        try:
            log('    · %s（%.1f MB）'
                % (os.path.basename(p), os.path.getsize(p) / 1048576.0))
        except Exception:
            log('    · %s' % os.path.basename(p))
    if len(files) > 1:
        log('    ↑ 请确认这里没有【以前生成的合并结果】。'
            '有的话它也会被当成输入合并进来，导致数据重复；'
            '建议把结果存到单独的文件夹。')

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


def copy_sheet_layout(wb_src, sheet_name, ws_out, keep_idx):
    """把模板 sheet 的版式复制到新 sheet：列宽 + 冻结窗格（keep_idx 为保留列的 0 基索引）

    wb_src 收【已经打开的工作簿】。以前这里自己 load_workbook，
    一个大文件要全量加载十几秒，而 do_merge 里同一份文件会被反复打开。
    """
    ws = wb_src[sheet_name]
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


def delete_unkept_cols(ws, keep_idx, log=None):
    """删掉未勾选的列 —— 必须按【连续区间】合并成几次调用，绝不能逐列删。

    为什么：openpyxl 的 delete_cols(idx, amount) 每调用一次都要重建整张表的
    单元格索引，所以【调用次数】才是成本，单次删多少列几乎不影响耗时。
    实测（2026年8月计划表.xlsx / 已完成：16681 行 x 266 列，139 万单元格）：
        逐列删 262 列  ->  0.78 秒/列 x 262 = 204 秒
        合并区间删除    ->  几次调用，十几秒
    必须从右往左删，否则前面的列号会错位。

    另外：delete_cols 只搬单元格，【列宽】和 autoFilter/条件格式/数据验证
    这些范围引用它一概不管，必须自己补（见 remap_after_delete）。
    """
    max_col = ws.max_column
    keep = sorted({i + 1 for i in keep_idx if 1 <= i + 1 <= max_col})

    # 删之前先记下保留列的列宽/隐藏状态，删完按新位置重设
    keep_dim = []
    for k in keep:
        dim = ws.column_dimensions.get(get_column_letter(k))
        keep_dim.append((getattr(dim, 'width', None),
                         bool(getattr(dim, 'hidden', False)) if dim is not None else False))

    ranges = []
    prev = max_col + 1                     # 从最右边界的下一列开始往左扫
    for k in reversed(keep):
        if prev - 1 >= k + 1:              # k 右边还有要删的列
            ranges.append((k + 1, prev - 1 - k))
        prev = k
    if prev - 1 >= 1:                      # 最左边那一段
        ranges.append((1, prev - 1))
    dropped = sum(a for _, a in ranges)
    for start, amount in ranges:
        ws.delete_cols(start, amount)

    # 重设列宽：不这么做的话，数据左移了列宽还留在原地，结果的列宽和母版对不上
    for key in list(ws.column_dimensions.keys()):
        try:
            del ws.column_dimensions[key]
        except Exception:
            pass
    for pos, (w, hid) in enumerate(keep_dim):
        letter = get_column_letter(pos + 1)
        if w:
            ws.column_dimensions[letter].width = w
        if hid:
            ws.column_dimensions[letter].hidden = True

    if log and dropped:
        log('    删除未勾选列 %d 个，合并成 %d 次区间删除（逐列删要 %d 次）'
            % (dropped, len(ranges), dropped))


def _map_range(rng, colmap):
    """把一个范围按【原列号→新列号】映射成若干连续范围；全落在被删列上则返回 []"""
    news = sorted({colmap[c] for c in range(rng.min_col, rng.max_col + 1) if c in colmap})
    if not news:
        return []
    runs, s, p = [], news[0], news[0]
    for c in news[1:]:
        if c == p + 1:
            p = c
        else:
            runs.append((s, p))
            s = p = c
    runs.append((s, p))
    return [CellRange(min_col=a, max_col=b, min_row=rng.min_row, max_row=rng.max_row)
            for a, b in runs]


def remap_after_delete(ws, keep_idx, log=None):
    """把 autoFilter / 条件格式 / 数据验证 的范围改到删列后的新列上。

    openpyxl 的 delete_cols 一个都不管。不修的话实测会出现：
      · autoFilter 还是 A1:JF16681，而表里只剩 3 列 → 筛选范围跑到表外
      · 条件格式原本挂在 H/I 列（正是保留的那两列），删列后还指着 H/I，
        于是【在保留列上彻底失效】；也有反过来挂错到别的列上的风险
    映射不上的（整段都在被删列里）就丢掉，符合"这些列本来就不合并了"的语义。
    """
    colmap = {src_i + 1: pos + 1 for pos, src_i in enumerate(keep_idx)}
    done = []

    # 条件格式
    try:
        old = list(ws.conditional_formatting)
        if old:
            new_cf = ConditionalFormattingList()
            dropped = 0
            for cf in old:
                rngs = []
                for r in cf.sqref.ranges:
                    rngs += _map_range(r, colmap)
                if not rngs:
                    dropped += 1
                    continue
                # ConditionalFormattingList.add(范围或CF对象, 单条规则)
                key = ConditionalFormatting(sqref=MultiCellRange(rngs))
                for rule in cf.rules:
                    new_cf.add(key, rule)
            new_cf.max_priority = max(getattr(old, 'max_priority', 0),
                                      new_cf.max_priority)
            ws.conditional_formatting = new_cf
            done.append('条件格式 %d 条（丢弃 %d 条只作用于被删列的）'
                        % (len(new_cf), dropped))
    except Exception as e:
        if log:
            log('    [!] 条件格式范围重映射失败：%s' % e)

    # 数据验证
    try:
        keep_dv = []
        n_drop = 0
        for dv in ws.data_validations.dataValidation:
            rngs = []
            for r in dv.sqref.ranges:
                rngs += _map_range(r, colmap)
            if rngs:
                dv.sqref = MultiCellRange(rngs)
                keep_dv.append(dv)
            else:
                n_drop += 1
        ws.data_validations.dataValidation = keep_dv
        done.append('数据验证 %d 条（丢弃 %d 条）' % (len(keep_dv), n_drop))
    except Exception as e:
        if log:
            log('    [!] 数据验证范围重映射失败：%s' % e)

    if log and done:
        log('    范围引用已重映射：' + '；'.join(done))


def fix_autofilter(ws, ncol, log=None):
    """筛选范围改成覆盖整表（要在追加完所有行之后调）"""
    try:
        if ws.auto_filter is not None and ws.auto_filter.ref:
            old = ws.auto_filter.ref
            ws.auto_filter.ref = 'A1:%s%d' % (get_column_letter(ncol), ws.max_row)
            if log and old != ws.auto_filter.ref:
                log('    筛选范围已修正：%s -> %s' % (old, ws.auto_filter.ref))
    except Exception as e:
        if log:
            log('    [!] 筛选范围修正失败：%s' % e)


def read_sheet_block(path, sheet_name, keep_names):
    """
    一次打开，返回 (表头名列表, {表头: 0基索引}, 数据行列表)

    表头与数据在同一次 open 里读完。以前是先开一次读表头、read_sheet_rows
    再开一次读数据，同一个大文件白读两遍。
    返回 (None, None, None) 表示该文件里没有这个工作表。
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return None, None, None
        ws = wb[sheet_name]
        it = ws.iter_rows(values_only=True)
        try:
            first = next(it)
        except StopIteration:
            first = ()
        hdr = [_s(v) for v in (first or ())]
        src_index = {}
        for i, h in enumerate(hdr):
            if h and h not in src_index:
                src_index[h] = i
        idxs = [src_index.get(nm) for nm in keep_names]
        out = []
        for row in it:
            if row is None:
                continue
            vals = [row[i] if (i is not None and i < len(row)) else None for i in idxs]
            if any(v is not None and _s(v) != '' for v in vals):
                out.append(vals)
        return hdr, src_index, out
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

    # ★ 必须先把【输出文件本身】从输入里剔除。
    #   结果一般就存在待合并文件夹里，下次再点合并时它会被当成输入文件扫进来，
    #   于是"上次的结果 + 原始文件"一起合并 → 行数翻倍、数据重复。
    skip = os.path.abspath(save_path)
    kept_sheets = {}
    for sn in targets:
        info = sheets[sn]
        fl = [p for p in info['files'] if os.path.abspath(p) != skip]
        if len(fl) != len(info['files']):
            log('  [!] 输出文件和某个待合并文件同名，已从输入里剔除：%s'
                % os.path.basename(skip))
        if not fl:
            log('  [!] 工作表【%s】剔除输出文件后已无输入文件，跳过' % sn)
            continue
        kept_sheets[sn] = dict(info, files=fl)
    if not kept_sheets:
        log('剔除输出文件后没有可合并的输入，未执行')
        return False
    sheets = kept_sheets
    targets = {sn: targets[sn] for sn in kept_sheets}

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
        # 模板工作簿只加载一次：native 时母版就是输出工作簿本身，连开都不用开
        wb_tpl = wb_out if native else load_workbook(info['base'], data_only=False)
        wb_tpl_owned = not native
        style_src = {}
        row_height = None
        try:
            if native:
                ws = wb_out[sn]
                delete_unkept_cols(ws, keep_idx, log)
                remap_after_delete(ws, keep_idx, log)
                # 母版自身的数据行直接保留 → 只追加其它文件
                append_files = [p for p in info['files']
                                if os.path.abspath(p) != os.path.abspath(out_base)]
            else:
                ws = wb_out.create_sheet(sn)
                copy_sheet_layout(wb_tpl, sn, ws, keep_idx)
                ws_tpl = wb_tpl[sn]
                for pos, src_i in enumerate(keep_idx):
                    src_cell = ws_tpl.cell(row=HEADER_ROW, column=src_i + 1)
                    dst_cell = ws.cell(row=HEADER_ROW, column=pos + 1)
                    dst_cell.value = all_cols[src_i]
                    copy_cell_style(src_cell, dst_cell)
                append_files = list(info['files'])   # 该表不在输出母版里 → 所有文件都要写

            # 追加行的样式来源 = 模板表的「第一条数据行」(row = HEADER_ROW+1)，逐列取。
            # native 时列已经删好了，直接从输出表读就行（母版格式原样，无需再开文件）；
            # 新建表时输出表里还没有数据行，必须从模板工作簿读，
            # 否则追加行会丢掉边框/数字格式。
            ws_st = wb_tpl[sn]
            if ws_st.max_row >= HEADER_ROW + 1:
                for pos, src_i in enumerate(keep_idx):
                    col = (pos + 1) if native else (src_i + 1)
                    style_src[pos + 1] = snapshot_style(
                        ws_st.cell(row=HEADER_ROW + 1, column=col))
                rd = ws_st.row_dimensions.get(HEADER_ROW + 1)
                if rd is not None:
                    row_height = rd.height
            else:
                log('  [!] 模板表【%s】只有表头、没有数据行，追加行将使用默认样式' % sn)
        except Exception as e:
            log('  [!] 读取模板样式失败，追加行用默认样式：%s' % e)
        finally:
            if wb_tpl_owned:
                wb_tpl.close()

        style_list = [style_src.get(i + 1) for i in range(ncol)]

        added = 0
        warned = set()
        missing_map = {}                  # {字段名: [缺该字段的文件名...]}
        for p in append_files:
            # 一次打开，表头和数据一起读（旧版在这里开了两次同一个文件）
            try:
                hdr, src_index, rows = read_sheet_block(p, sn, keep_names)
            except Exception as e:
                log('  [跳过] %s / %s 读取失败：%s' % (os.path.basename(p), sn, e))
                continue
            if hdr is None:
                continue

            for nm in keep_names:
                if nm not in src_index:
                    missing_map.setdefault(nm, []).append(os.path.basename(p))
                    if nm not in warned:
                        warned.add(nm)
                        log('  [!] %s 里缺少字段「%s」，该列留空'
                            % (os.path.basename(p), nm.replace('\n', ' ')))

            if not rows:
                log('  = %s：没有非空数据行' % os.path.basename(p))
                continue

            r0 = ws.max_row + 1
            base = os.path.basename(p)
            for ri, vals in enumerate(rows):
                rr = r0 + ri
                for ci, v in enumerate(vals):
                    cell = ws.cell(row=rr, column=ci + 1)
                    cell.value = v
                    snap = style_list[ci]
                    if snap is not None:
                        apply_snapshot(cell, snap)
                if row_height:
                    ws.row_dimensions[rr].height = row_height
                if (ri + 1) % 3000 == 0:          # 大表给点进度反馈，避免看着像死了
                    log('    … %s 已追加 %d / %d 行' % (base, ri + 1, len(rows)))
            added += len(rows)
            log('  + %s：追加 %d 行' % (base, len(rows)))

        total_rows += added
        fix_autofilter(ws, ncol, log)     # 筛选范围要覆盖追加后的全部行
        log('工作表【%s】完成：保留 %d 个字段，追加 %d 行' % (sn, ncol, added))
        if missing_map:
            # 这个汇总很重要：字段名只要有一点不一样（多语言后缀、换行、型号/产品编号），
            # 就要靠【表头名完全相同】来对齐，对不上的文件那几列就是空的。
            log('  [!] 本表有 %d 个字段在部分文件里没找到（那些文件的对应列留空）：'
                % len(missing_map))
            for nm, fl in missing_map.items():
                log('        「%s」在 %d 个文件里没有：%s'
                    % (nm.replace('\n', ' '), len(fl), '、'.join(sorted(set(fl)))))

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
    DESC = "合并多个 Excel 中同名工作表，可选字段、保留原格式（大表在后台跑，界面不卡）"

    def __init__(self, master):
        super().__init__(master)

        self.folder_var = tk.StringVar()
        self.save_var = tk.StringVar()
        self.status_var = tk.StringVar(value='就绪')
        self.sheets = {}          # scan 结果
        self.col_vars = {}        # {sheet: {列名: BooleanVar}}
        self.cur_sheet = None     # 当前显示的工作表

        # 后台线程 → 界面 的消息队列。
        # 扫描/合并都放到子线程里跑，子线程【只往队列里放消息】，绝不碰任何控件；
        # 界面线程用 after() 定时取消息来刷新。这样窗口永远能响应，
        # 不会出现"未响应"（那是耗时代码跑在界面线程里造成的）。
        self._q = queue.Queue()
        self._busy = False
        self._buttons = []

        self._build_top()
        self._build_mid()
        self._build_bottom()
        self._poll()

    # ---------- 顶部：选文件夹 ----------
    def _build_top(self):
        f = tk.LabelFrame(self, text='1. 选择包含待合并 Excel 的文件夹')
        f.pack(fill=tk.X, padx=10, pady=(10, 4))
        tk.Entry(f, textvariable=self.folder_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=6)
        b1 = tk.Button(f, text='浏览', width=7, command=self.pick_folder)
        b1.pack(side=tk.LEFT, padx=4)
        b2 = tk.Button(f, text='扫描', width=7, bg='#2E86AB', fg='white', command=self.do_scan)
        b2.pack(side=tk.LEFT, padx=(0, 6))
        self._buttons += [b1, b2]

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

    # ---------- 底部：保存 + 按钮 + 状态 + 日志 ----------
    def _build_bottom(self):
        f = tk.LabelFrame(self, text='4. 保存合并结果')
        f.pack(fill=tk.X, padx=10, pady=4)
        tk.Entry(f, textvariable=self.save_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=6)
        b1 = tk.Button(f, text='另存为', width=7, command=self.pick_save)
        b1.pack(side=tk.LEFT, padx=4)
        b2 = tk.Button(f, text='▶ 开始合并', width=12, bg='#E67E22', fg='white',
                       font=('微软雅黑', 10, 'bold'), command=self.run)
        b2.pack(side=tk.LEFT, padx=(0, 6))
        self._buttons += [b1, b2]

        tk.Label(self, textvariable=self.status_var, anchor='w',
                 fg='#2E86AB', font=('微软雅黑', 9)).pack(fill=tk.X, padx=14, pady=(0, 2))

        lf = tk.LabelFrame(self, text='运行日志')
        lf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(2, 10))
        self.log_text = tk.Text(lf, height=7, wrap='word')
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0), pady=6)
        lsb = tk.Scrollbar(lf, orient='vertical', command=self.log_text.yview)
        lsb.pack(side=tk.LEFT, fill=tk.Y, pady=6, padx=(0, 6))
        self.log_text.config(yscrollcommand=lsb.set)
        self._write_log('就绪。请选择文件夹后点「扫描」。')

    def _on_wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 120), 'units')

    # ---------- 后台任务 ↔ 界面 的桥 ----------
    def _poll(self):
        """界面线程定时取后台消息 —— 所有控件操作都只在这里发生"""
        buf = []

        def flush():
            """日志攒成一批插一次。一行一插在消息密集时会把界面拖住。"""
            if buf:
                self.log_text.insert(tk.END, '\n'.join(buf) + '\n')
                self.log_text.see(tk.END)
                del buf[:]

        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == 'log':
                    buf.append(str(payload))
                    if len(buf) >= 200:
                        flush()
                    continue
                flush()
                if kind == 'status':
                    self.status_var.set(payload)
                elif kind == 'scan_done':
                    self._after_scan(payload)
                elif kind == 'merge_done':
                    self._after_merge(payload)
                elif kind == 'error':
                    self._write_log('出错：\n' + payload)
                    self._set_busy(False)
                    self.status_var.set('出错')
                    messagebox.showerror('出错', '详见下方日志')
        except queue.Empty:
            pass
        flush()
        self.after(120, self._poll)

    def _write_log(self, msg):
        self.log_text.insert(tk.END, str(msg) + '\n')
        self.log_text.see(tk.END)

    def log(self, msg):
        """界面线程记日志"""
        self._write_log(msg)

    def _worker_log(self, msg):
        """后台线程记日志：只入队，绝不碰控件"""
        self._q.put(('log', str(msg)))

    def _set_busy(self, busy):
        self._busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        for b in self._buttons:
            b.config(state=state)
        try:
            self.config(cursor='watch' if busy else '')
        except Exception:
            pass

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

    @staticmethod
    def _field_names(cols):
        """给界面用的字段名列表：去掉空列名、去重、把换行压成空格便于显示"""
        out, seen = [], set()
        for c in cols:
            if c and c not in seen:
                seen.add(c)
                out.append(c)
        return out

    def do_scan(self):
        if self._busy:
            return
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning('提示', '请先选择有效的文件夹')
            return
        self.log_text.delete('1.0', tk.END)
        self.sheets = {}
        self.col_vars = {}
        self.cur_sheet = None
        self.lb.delete(0, tk.END)
        for w in self.col_frame.winfo_children():
            w.destroy()
        self._set_busy(True)
        self.status_var.set('正在扫描…（大文件要几秒，界面不会卡）')
        self._write_log('开始扫描：%s' % folder)
        threading.Thread(target=self._scan_job, args=(folder,), daemon=True).start()

    def _scan_job(self, folder):
        try:
            _files, sheets = scan_folder(folder, self._worker_log)
            self._q.put(('scan_done', sheets))
        except Exception:
            self._q.put(('error', traceback.format_exc()))

    def _after_scan(self, sheets):
        self.sheets = sheets
        self._set_busy(False)
        if not sheets:
            self.status_var.set('没有扫描到工作表')
            self._write_log('没有扫描到任何工作表。')
            messagebox.showwarning('提示', '没有扫描到可用工作表（检查文件是否为 .xlsx/.xlsm）')
            return
        for sn in sheets:
            n = len(sheets[sn]['files'])
            self.lb.insert(tk.END, '%s  (%d文件)' % (sn, n))
            self.col_vars[sn] = {c: tk.BooleanVar(value=True)
                                 for c in self._field_names(sheets[sn]['columns'])}
        self.lb.selection_set(0)
        self.on_pick_sheet()
        self.status_var.set('扫描完成：%d 个工作表名' % len(sheets))
        self._write_log('扫描完成：%d 个工作表名。默认全选字段，按需取消。' % len(sheets))

    def on_pick_sheet(self):
        sel = self.lb.curselection()
        if not sel:
            return
        keys = list(self.sheets.keys())
        if sel[0] >= len(keys):
            return
        sn = keys[sel[0]]
        self.cur_sheet = sn
        for w in self.col_frame.winfo_children():
            w.destroy()
        names = self._field_names(self.sheets[sn]['columns'])
        self._write_log('当前工作表【%s】%d 个字段（原始表头 %d 列）'
                        % (sn, len(names), len(self.sheets[sn]['columns'])))
        for i, c in enumerate(names):
            cb = tk.Checkbutton(self.col_frame,
                                text=c.replace('\r', ' ').replace('\n', ' '),
                                variable=self.col_vars[sn][c], anchor='w', width=34)
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
        if self._busy:
            return
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
            cols = [c for c in self.sheets[sn]['columns'] if c in d and d[c].get()]
            if cols:
                selection[sn] = cols
        if not selection:
            messagebox.showwarning('提示', '至少要在一个工作表里勾选字段')
            return
        # 结果存进待合并文件夹的话，下次扫描会把它也算进去 —— 提醒一下并说明已自动剔除
        try:
            if os.path.abspath(os.path.dirname(save)) == os.path.abspath(folder):
                self._write_log('[!] 输出文件保存在待合并文件夹里。本次已自动把输出文件'
                                '从输入中剔除，但建议改存到别的文件夹，更不容易搞混。')
        except Exception:
            pass
        self._write_log('─' * 60)
        self._set_busy(True)
        self.status_var.set('正在合并…（大表要几分钟，界面不会卡，可看日志进度）')
        threading.Thread(target=self._merge_job,
                         args=(folder, save, dict(self.sheets), selection),
                         daemon=True).start()

    def _merge_job(self, folder, save, sheets, selection):
        try:
            ok = do_merge(folder, save, sheets, selection, self._worker_log)
            self._q.put(('merge_done', (ok, save)))
        except Exception:
            self._q.put(('error', traceback.format_exc()))

    def _after_merge(self, payload):
        ok, save = payload
        self._set_busy(False)
        if ok:
            self.status_var.set('合并完成')
            messagebox.showinfo('完成', '合并完成！\n保存到：\n%s' % save)
        else:
            self.status_var.set('合并未完成')
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

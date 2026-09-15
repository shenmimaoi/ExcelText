# -*- coding: utf-8 -*-
"""把写法不规范的编号自动匹配到规范写法（10级派生匹配）

由原工具源码移植。逻辑未改动，仅把主窗口改成可嵌入的功能面板。
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from openpyxl import load_workbook

from .base import ToolFrame
from .matching import normalize, build_index, lookup, to_text


class CleanMatchTool(ToolFrame):
    NAME = "编号清洗匹配"
    DESC = "把写法不规范的编号自动匹配到规范写法（10级派生匹配）"
    def __init__(self, master):
        super().__init__(master)

        self.src_path = tk.StringVar()
        self.map_path = tk.StringVar()

        self.col_target = tk.StringVar(value="产品编号")
        self.col_clean = tk.StringVar(value="清洗后编号")
        self.col_new = tk.StringVar(value="正确成品编号")
        self.col_map_key = tk.StringVar(value="原始编号")
        self.col_map_val = tk.StringVar(value="正确编号")

        # 可选开关
        self.opt_strip_bracket = tk.BooleanVar(value=False)   # 删除括号及内容
        self.opt_status_col = tk.BooleanVar(value=True)       # 输出"匹配状态"列

        # ========== 文件选择区 ==========
        frame_file = tk.LabelFrame(self, text="文件选择")
        frame_file.pack(padx=10, pady=5, fill=tk.X)

        tk.Label(frame_file, text="原数据文件：").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(frame_file, textvariable=self.src_path, width=60).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame_file, text="选择文件", command=self.select_src).grid(row=0, column=2, padx=5)

        tk.Label(frame_file, text="对照匹配表：").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(frame_file, textvariable=self.map_path, width=60).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame_file, text="选择文件", command=self.select_map).grid(row=1, column=2, padx=5)

        # ========== 列名配置区 ==========
        frame_col = tk.LabelFrame(self, text="列名配置（和Excel表头完全一致）")
        frame_col.pack(padx=10, pady=5, fill=tk.X)

        tk.Label(frame_col, text="【原数据表】", fg="#3080E0").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        tk.Label(frame_col, text="待处理编号列：").grid(row=0, column=1, padx=5, pady=3, sticky="e")
        tk.Entry(frame_col, textvariable=self.col_target, width=16).grid(row=0, column=2, padx=5, pady=3)
        tk.Label(frame_col, text="清洗后列名：").grid(row=0, column=3, padx=5, pady=3, sticky="e")
        tk.Entry(frame_col, textvariable=self.col_clean, width=16).grid(row=0, column=4, padx=5, pady=3)
        tk.Label(frame_col, text="正确结果列名：").grid(row=0, column=5, padx=5, pady=3, sticky="e")
        tk.Entry(frame_col, textvariable=self.col_new, width=16).grid(row=0, column=6, padx=5, pady=3)

        tk.Label(frame_col, text="【对照表】", fg="#28A745").grid(row=1, column=0, padx=5, pady=3, sticky="w")
        tk.Label(frame_col, text="原始编号列：").grid(row=1, column=1, padx=5, pady=3, sticky="e")
        tk.Entry(frame_col, textvariable=self.col_map_key, width=16).grid(row=1, column=2, padx=5, pady=3)
        tk.Label(frame_col, text="正确编号列：").grid(row=1, column=3, padx=5, pady=3, sticky="e")
        tk.Entry(frame_col, textvariable=self.col_map_val, width=16).grid(row=1, column=4, padx=5, pady=3)

        # ========== 规则说明 ==========
        frame_rule = tk.LabelFrame(self, text="清洗 + 匹配规则（自动执行）")
        frame_rule.pack(padx=10, pady=5, fill=tk.X)
        rule_text = """清洗规则：去所有空白(NBSP/零宽字符) → NFKC全角转半角 → 英寸号 ”″ → " → 各种横杠 –—−－ → - → 中文括号转英文
匹配规则：① 精确  ② 英寸→in（7”Plate = 7inPlate）  ③ 英寸去除  ④ 去尾部序号（123-1 = 123，双向）
          ⑤ 去全部尾部序号  ⑥ 去横杠  ⑦ 数字去前导零
输出：原列 → 清洗后编号 → 正确成品编号 →（可选）匹配状态"""
        tk.Label(frame_rule, text=rule_text, fg="#666", justify="left", anchor="w").pack(padx=5, pady=5, fill=tk.X)

        frame_opt = tk.Frame(self)
        frame_opt.pack(padx=10, pady=2, fill=tk.X)
        tk.Checkbutton(frame_opt, text="删除括号及括号内内容（默认关闭，保持原行为）",
                       variable=self.opt_strip_bracket).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(frame_opt, text="输出「匹配状态」列",
                       variable=self.opt_status_col).pack(side=tk.LEFT, padx=15)

        # ========== 日志区 ==========
        frame_log = tk.LabelFrame(self, text="运行日志")
        frame_log.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(frame_log, height=14)
        self.log_text.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        # ========== 按钮区 ==========
        frame_btn = tk.Frame(self)
        frame_btn.pack(pady=8)
        tk.Button(frame_btn, text="开始处理", command=self.run_task, bg="#3080E0", fg="white", width=12).pack(side=tk.LEFT, padx=10)
        tk.Button(frame_btn, text="清空日志", command=self.clear_log, width=10).pack(side=tk.LEFT)

    # ---------------- 基础 ----------------
    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def select_src(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xlsm")])
        if path:
            self.src_path.set(path)

    def select_map(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xlsm")])
        if path:
            self.map_path.set(path)

    # ---------------- 读取对照表 ----------------
    def build_mapping_pairs(self, map_file):
        wb = load_workbook(map_file, data_only=True)
        ws = wb.active

        header = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col).value
            if val is not None:
                header[str(val).strip()] = col

        key_col = self.col_map_key.get().strip()
        val_col = self.col_map_val.get().strip()

        if key_col not in header or val_col not in header:
            raise ValueError("对照表缺少列！\n需要：%s、%s\n实际表头：%s"
                             % (key_col, val_col, list(header.keys())))

        key_idx = header[key_col]
        val_idx = header[val_col]

        pairs = []
        for row in range(2, ws.max_row + 1):
            k = ws.cell(row=row, column=key_idx).value
            v = ws.cell(row=row, column=val_idx).value
            if k is None and v is None:
                continue
            pairs.append((to_text(k), to_text(v)))
        wb.close()
        return pairs

    # ---------------- 主流程 ----------------
    def run_task(self):
        src_file = self.src_path.get().strip()
        map_file = self.map_path.get().strip()

        if not os.path.exists(src_file):
            messagebox.showerror("错误", "请先选择原数据文件！")
            return
        if not os.path.exists(map_file):
            messagebox.showerror("错误", "请先选择对照匹配表文件！")
            return

        self.clear_log()
        strip_bracket = self.opt_strip_bracket.get()
        self.log("=" * 62)
        self.log("原数据文件：%s" % src_file)
        self.log("对照匹配表：%s" % map_file)
        self.log("删除括号内容：%s" % ("是" if strip_bracket else "否"))
        self.log("-" * 62)

        # 1) 读取对照表并建索引
        self.log("正在读取对照表并构建多级匹配索引...")
        try:
            pairs = self.build_mapping_pairs(map_file)
        except Exception as e:
            self.log("[X] 对照表读取失败：%s" % str(e))
            messagebox.showerror("对照表错误", str(e))
            return
        idx = build_index(pairs)
        self.log("[OK] 对照表 %d 行，索引就绪" % len(pairs))

        # 2) 打开原数据
        self.log("正在读取原数据文件...")
        try:
            wb = load_workbook(src_file)
            ws = wb.active
        except Exception as e:
            self.log("[X] 原文件读取失败：%s" % str(e))
            messagebox.showerror("读取失败", str(e))
            return

        # 3) 定位列（已存在则复用，避免重复插入）
        target_col_name = self.col_target.get().strip()
        clean_col_name = self.col_clean.get().strip()
        new_col_name = self.col_new.get().strip()
        status_col_name = "匹配状态"

        headers = {}
        for col in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=col).value
            if v is not None:
                headers[str(v).strip()] = col

        if target_col_name not in headers:
            msg = "原表找不到列：%s\n实际表头：%s" % (target_col_name, list(headers.keys()))
            self.log("[X] " + msg)
            messagebox.showerror("列名错误", msg)
            wb.close()
            return
        target_col_idx = headers[target_col_name]
        self.log("[OK] 定位目标列：第 %d 列【%s】" % (target_col_idx, target_col_name))

        need_status = self.opt_status_col.get()
        out_names = [clean_col_name, new_col_name] + ([status_col_name] if need_status else [])
        missing = [n for n in out_names if n not in headers]
        if missing:
            insert_pos = target_col_idx + 1
            ws.insert_cols(insert_pos, len(out_names))
            for i, n in enumerate(out_names):
                ws.cell(row=1, column=insert_pos + i).value = n
            self.log("[OK] 插入 %d 列：%s" % (len(out_names), "、".join(out_names)))
            col_clean_i = insert_pos
            col_new_i = insert_pos + 1
            col_status_i = insert_pos + 2 if need_status else None
        else:
            col_clean_i = headers[clean_col_name]
            col_new_i = headers[new_col_name]
            col_status_i = headers[status_col_name] if need_status else None
            self.log("[OK] 已存在输出列，直接复用（不重复插入）")

        # 4) 逐行清洗 + 匹配
        match_count = 0
        total_row = 0
        rule_stat = {}
        ambiguous = []
        unmatched = []

        for row in range(2, ws.max_row + 1):
            raw_val = ws.cell(row=row, column=target_col_idx).value
            if raw_val is None or to_text(raw_val).strip() == "":
                continue
            total_row += 1

            clean_val = normalize(raw_val, strip_bracket)
            ws.cell(row=row, column=col_clean_i).value = clean_val
            if not clean_val:
                continue

            correct_val, rule, amb = lookup(raw_val, idx, strip_bracket)
            if correct_val:
                match_count += 1
                ws.cell(row=row, column=col_new_i).value = correct_val
                rule_stat[rule] = rule_stat.get(rule, 0) + 1
                if amb:
                    ambiguous.append((row, clean_val, correct_val))
                if col_status_i:
                    ws.cell(row=row, column=col_status_i).value = "匹配(%s)%s" % (rule, " 有歧义" if amb else "")
            else:
                unmatched.append((row, clean_val))
                if col_status_i:
                    ws.cell(row=row, column=col_status_i).value = "未匹配"

        # 5) 日志统计
        self.log("-" * 62)
        self.log("共处理数据行：%d 行" % total_row)
        self.log("[OK] 匹配成功：%d 行  (未匹配 %d 行)" % (match_count, len(unmatched)))
        if rule_stat:
            self.log("命中规则分布：")
            for k, v in sorted(rule_stat.items(), key=lambda x: -x[1]):
                self.log("    %-16s %d 行" % (k, v))

        if ambiguous:
            self.log("-" * 62)
            self.log("[!] 存在歧义（同一写法对应多个不同结果，已按规则择优）%d 行，前20条：" % len(ambiguous))
            for r, cv, val in ambiguous[:20]:
                self.log("    第%d行  %s -> %s" % (r, cv, val))

        if unmatched:
            self.log("-" * 62)
            self.log("[!] 未匹配 %d 行，前30条（可据此补充对照表）：" % len(unmatched))
            for r, cv in unmatched[:30]:
                self.log("    第%d行  %s" % (r, cv))

        # 6) 保存
        base, ext = os.path.splitext(src_file)
        output_file = "%s_编号匹配后%s" % (base, ext)
        try:
            wb.save(output_file)
            wb.close()
        except Exception as e:
            self.log("[X] 保存失败：%s" % str(e))
            messagebox.showerror("保存失败", str(e))
            return

        self.log("-" * 62)
        self.log("[OK] 处理完成！原格式已保留")
        self.log("输出文件：%s" % output_file)
        messagebox.showinfo(
            "执行完毕",
            "处理完成！\n处理行数：%d\n匹配成功：%d\n未匹配：%d\n输出：%s"
            % (total_row, match_count, len(unmatched), output_file))

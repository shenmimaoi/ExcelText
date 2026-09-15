# -*- coding: utf-8 -*-
"""按关键词匹配行，批量修改该行多个列的取值（含预置规则）

由原工具源码移植。逻辑未改动，仅把主窗口改成可嵌入的功能面板。
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from openpyxl import load_workbook

from .base import ToolFrame


class KeywordEditTool(ToolFrame):
    NAME = "关键词批量改列"
    DESC = "按关键词匹配行，批量修改该行多个列的取值（含预置规则）"
    def __init__(self, master):
        super().__init__(master)

        self.input_path = tk.StringVar()
        self.col_key = tk.StringVar(value="物料名称")  # 判断关键词的列

        # -------- 文件选择 --------
        frame_file = tk.LabelFrame(self, text="文件选择")
        frame_file.pack(padx=10, pady=5, fill=tk.X)
        tk.Entry(frame_file, textvariable=self.input_path, width=65).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(frame_file, text="选择Excel", command=self.select_file).pack(side=tk.LEFT, padx=5)

        # -------- 列配置 --------
        frame_col = tk.LabelFrame(self, text="匹配判断列：用来检索关键词的表头")
        frame_col.pack(padx=10, pady=5, fill=tk.X)
        tk.Label(frame_col, text="判断列(如物料名称)：").grid(row=0, column=0, padx=5, pady=3, sticky="e")
        tk.Entry(frame_col, textvariable=self.col_key, width=30).grid(row=0, column=1, padx=5, pady=3)

        # -------- 规则文本框 --------
        frame_rule = tk.LabelFrame(self, text="规则格式：关键词 列名1=值1,列名2=值2,列名3=值3")
        frame_rule.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        self.txt_rule = scrolledtext.ScrolledText(frame_rule, height=12)
        self.txt_rule.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        # 预置示例规则，和之前完全兼容
        demo_text = """纸箱 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
透明袋 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
不干胶 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
格挡 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
PVC盒 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
翻盖双钮簧SS304 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
印刷袋 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
自黏防滑硅胶胶粒 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
翻盖硅胶塞（气相胶40°） 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
地盖 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
淋膜纸 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
托盘 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
包装纸 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
拉链袋 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
包装膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
圆标 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
立柱护角 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
印刷膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
雪梨纸 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
外箱标贴-项目标 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
蜂窝板 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
天盖 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
按钮弹簧SS304 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
异形纸板 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
对折膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
皮筋 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
光膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
棉纸 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
PE塑片 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
梨纸 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
纸卡头 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
展示盒 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
纸扎丝 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
纸卡 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
防割衬板 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
隔板 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
热缩膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
附件 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
衬板 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
吊卡 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
OPP热封膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
透明插边袋 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
热缩带 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
外箱标贴-黑白标 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
RFID吊卡 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
围卡 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
纸盒 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
绑带 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
桶膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
翻盖五金销SS304 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
挂钩 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
挂条 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
大盖密封圈（气相胶50°） 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
金色扎丝 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
硅胶绑带 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
窗口盒 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
PET盒 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
瓦楞盒 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
片膜 包材-渠道=外购半成品,包材-单位=张,包材-采购单位=kg
胶印箱 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
密封圈 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
纸滑托 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
热缩袋 包材-渠道=外购成品,包材-单位=个,包材-采购单位=个
RFID不干胶 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
圆形不干胶 包材-渠道=外购成品,包材-单位=张,包材-采购单位=张
"""
        self.txt_rule.insert(tk.END, demo_text)

        # -------- 日志输出 --------
        frame_log = tk.LabelFrame(self, text="运行日志")
        frame_log.pack(padx=10, pady=5, fill=tk.X)
        self.log_text = scrolledtext.ScrolledText(frame_log, height=7)
        self.log_text.pack(padx=5, pady=5, fill=tk.X)

        # -------- 按钮 --------
        frame_btn = tk.Frame(self)
        frame_btn.pack(pady=8)
        tk.Button(frame_btn, text="开始处理", command=self.run_task, bg="#3080E0", fg="white", width=12).pack(side=tk.LEFT, padx=10)
        tk.Button(frame_btn, text="清空日志", command=self.clear_log, width=10).pack(side=tk.LEFT)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def select_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xlsm")]
        )
        if path:
            self.input_path.set(path)

    def parse_rules(self):
        """
        解析规则文本
        返回格式： [("纸箱", {"物料类型":"采购成品","单位":"个","采购单位":"个"}), ... ]
        """
        lines = self.txt_rule.get("1.0", tk.END).splitlines()
        rule_list = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            parts = line.split(maxsplit=1)
            if len(parts) != 2:
                continue
            keyword, assign_str = parts
            assign_dict = {}
            assign_items = assign_str.split(",")
            for item in assign_items:
                item = item.strip()
                if "=" in item:
                    col, val = item.split("=", maxsplit=1)
                    assign_dict[col.strip()] = val.strip()
            rule_list.append((keyword, assign_dict))
        return rule_list

    def run_task(self):
        input_file = self.input_path.get().strip()
        if not os.path.exists(input_file):
            messagebox.showerror("错误", "请先选择Excel文件！")
            return

        rule_list = self.parse_rules()
        if len(rule_list) == 0:
            messagebox.showwarning("提示", "没有配置任何匹配规则！")
            return

        judge_col_name = self.col_key.get().strip()

        self.clear_log()
        self.log(f"读取文件：{input_file}")
        self.log(f"解析完成 {len(rule_list)} 条匹配规则")
        self.log("模式：保留原文件所有格式，仅修改目标单元格的值")

        try:
            # 加载原工作簿，保留所有样式、公式、格式
            wb = load_workbook(input_file, data_only=False)
            ws = wb.active  # 默认第一个工作表
        except Exception as e:
            self.log(f"读取Excel失败:{str(e)}")
            messagebox.showerror("读取失败", str(e))
            return

        # 第一步：读取表头，建立 列名→列号 的映射
        header_row = 1  # 表头在第1行
        col_map = {}  # {"物料名称": 1, "物料类型": 2, ...}
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=header_row, column=col_idx).value
            if cell_val is not None:
                col_map[str(cell_val).strip()] = col_idx

        # 检查判断列是否存在
        if judge_col_name not in col_map:
            msg = f"判断列不存在！列名：{judge_col_name}\n实际表头：{list(col_map.keys())}"
            self.log(msg)
            messagebox.showerror("表头错误", msg)
            return

        # 检查所有要修改的列是否存在
        all_modify_cols = set()
        for _, assign_dict in rule_list:
            all_modify_cols.update(assign_dict.keys())
        miss_cols = [c for c in all_modify_cols if c not in col_map]
        if miss_cols:
            msg = f"发现表格缺失表头：{miss_cols}\n现有表头：{list(col_map.keys())}"
            self.log(msg)
            messagebox.showerror("表头缺失", msg)
            return

        judge_col_idx = col_map[judge_col_name]
        total_change = 0

        # 第二步：遍历所有数据行（从第2行开始）
        for row_idx in range(2, ws.max_row + 1):
            name_cell = ws.cell(row=row_idx, column=judge_col_idx)
            name_val = str(name_cell.value) if name_cell.value is not None else ""

            hit = False
            change_info = []
            for keyword, assign_dict in rule_list:
                if keyword in name_val:
                    hit = True
                    # 循环修改当前行的多个列
                    for col_name, new_val in assign_dict.items():
                        col_idx = col_map[col_name]
                        target_cell = ws.cell(row=row_idx, column=col_idx)
                        old_val = str(target_cell.value) if target_cell.value is not None else ""
                        if old_val != str(new_val):
                            target_cell.value = new_val
                            change_info.append(f"{col_name}:{old_val}→{new_val}")
                    break

            if hit and len(change_info) > 0:
                total_change += 1
                self.log(f"行{row_idx} | {name_val} | {' | '.join(change_info)}")

        # 第三步：另存为新文件，原文件不动
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_处理后{ext}"
        wb.save(output_file)
        wb.close()

        self.log("-" * 50)
        self.log(f"✅处理完成！一共改动 {total_change} 行数据")
        self.log(f"所有原格式、样式、公式已完整保留")
        self.log(f"输出文件路径：{output_file}")
        messagebox.showinfo("执行完毕", f"处理完成！\n修改行数：{total_change}\n输出：{output_file}\n原格式已完整保留")

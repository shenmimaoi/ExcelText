# -*- coding: utf-8 -*-
"""
工具基类 —— 所有功能模块都继承 ToolFrame。

【如何扩展新功能】
    1. 在 tools/ 下新建一个 .py，定义 class XxxTool(ToolFrame)
    2. 设置 NAME（左侧菜单显示名）和 DESC（功能说明）
    3. 在 _build() 里搭界面，需要日志就调 self._create_log(parent)
    4. 在 tools/__init__.py 的 TOOLS 列表里加一行
    —— 完成，重新打包即可
"""
import tkinter as tk
from tkinter import scrolledtext


class ToolFrame(tk.Frame):
    """功能面板基类。"""

    #: 左侧菜单里显示的名字
    NAME = "未命名工具"
    #: 顶部显示的说明
    DESC = ""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)

    # ------------------------------------------------------------------
    # 日志区（各工具共用）
    # ------------------------------------------------------------------
    def _create_log(self, parent, height=10, fill=tk.BOTH, expand=True, title="运行日志"):
        """创建标准「运行日志」面板，并绑定 self.log_text。返回该面板。"""
        box = tk.LabelFrame(parent, text=title)
        box.pack(padx=8, pady=6, fill=fill, expand=expand)
        self.log_text = scrolledtext.ScrolledText(box, height=height)
        self.log_text.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        return box

    def log(self, msg):
        """向日志区追加一行（自动滚动到底部）。"""
        if not hasattr(self, "log_text"):
            return
        self.log_text.insert(tk.END, str(msg) + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def clear_log(self):
        """清空日志区。"""
        if hasattr(self, "log_text"):
            self.log_text.delete(1.0, tk.END)

# -*- coding: utf-8 -*-
"""
Excel 工具箱 —— 统一入口

把多个 Excel 处理工具合并为单一程序：左侧菜单切换功能，右侧显示对应面板。
每个功能都是 tools/ 下的独立模块，新增功能无需改动本文件。

运行：python main.py
打包：python build.py   （或双击 build.cmd）
"""
import sys
import traceback
import tkinter as tk

from tools import TOOLS, tool_names

APP_NAME = "Excel 工具箱"
APP_VERSION = "v2.0.0"

# 配色
C_SIDE_BG = "#2B3A4A"
C_SIDE_FG = "#D6E0EA"
C_SIDE_FG_DIM = "#7C90A6"
C_SIDE_SEL = "#3C5A78"
C_MAIN_BG = "#FFFFFF"
C_TEXT = "#1F2D3D"
C_TEXT_DIM = "#7A8B9A"
C_LINE = "#E3E8EE"
C_STATUS_BG = "#F1F3F6"


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("%s %s" % (APP_NAME, APP_VERSION))
        self.root.geometry("1000x780")
        self.root.minsize(900, 620)
        self.root.configure(bg=C_MAIN_BG)

        self._instances = {}    # 工具索引 -> 面板实例（懒加载，切换时保留状态）
        self._current = None
        self._buttons = []

        self._build_statusbar()   # 先放，保证底部通栏
        self._build_sidebar()
        self._build_main()

        if TOOLS:
            self.select(0)

    # ------------------------------------------------------------------
    # 左侧：功能菜单
    # ------------------------------------------------------------------
    def _build_sidebar(self):
        bar = tk.Frame(self.root, width=178, bg=C_SIDE_BG)
        bar.pack(side=tk.LEFT, fill=tk.Y)
        bar.pack_propagate(False)

        tk.Label(bar, text="功能列表", bg=C_SIDE_BG, fg=C_SIDE_FG_DIM,
                 font=("微软雅黑", 9, "bold"), anchor="w").pack(fill=tk.X, padx=14, pady=(16, 6))

        for i, cls in enumerate(TOOLS):
            btn = tk.Button(
                bar, text="  " + cls.NAME, anchor="w", relief=tk.FLAT, bd=0,
                bg=C_SIDE_BG, fg=C_SIDE_FG,
                activebackground=C_SIDE_SEL, activeforeground="#FFFFFF",
                padx=10, pady=9, font=("微软雅黑", 10), cursor="hand2",
                command=lambda i=i: self.select(i),
            )
            btn.pack(fill=tk.X, padx=6, pady=1)
            self._buttons.append(btn)

        tk.Label(
            bar,
            text="新增功能：\n在 tools/ 里加模块，\n再到 __init__.py 注册",
            bg=C_SIDE_BG, fg=C_SIDE_FG_DIM, justify="left", font=("微软雅黑", 8),
        ).pack(side=tk.BOTTOM, fill=tk.X, padx=14, pady=14)

    # ------------------------------------------------------------------
    # 右侧：标题 + 说明 + 工具面板
    # ------------------------------------------------------------------
    def _build_main(self):
        self.main = tk.Frame(self.root, bg=C_MAIN_BG)
        self.main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        head = tk.Frame(self.main, bg=C_MAIN_BG)
        head.pack(fill=tk.X, padx=12, pady=(12, 0))
        self.lbl_name = tk.Label(head, text="", bg=C_MAIN_BG, fg=C_TEXT,
                                 font=("微软雅黑", 13, "bold"), anchor="w")
        self.lbl_name.pack(fill=tk.X)
        self.lbl_desc = tk.Label(head, text="", bg=C_MAIN_BG, fg=C_TEXT_DIM,
                                 font=("微软雅黑", 9), anchor="w")
        self.lbl_desc.pack(fill=tk.X, pady=(3, 0))

        tk.Frame(self.main, height=1, bg=C_LINE).pack(fill=tk.X, padx=12, pady=8)

        self.body = tk.Frame(self.main, bg=C_MAIN_BG)
        self.body.pack(fill=tk.BOTH, expand=True)

    # ------------------------------------------------------------------
    # 底部状态栏
    # ------------------------------------------------------------------
    def _build_statusbar(self):
        bar = tk.Frame(self.root, height=26, bg=C_STATUS_BG)
        bar.pack(side=tk.BOTTOM, fill=tk.X)
        bar.pack_propagate(False)
        tk.Label(bar, text="  %s %s" % (APP_NAME, APP_VERSION), bg=C_STATUS_BG,
                 fg=C_TEXT_DIM, font=("微软雅黑", 8)).pack(side=tk.LEFT, pady=4)
        tk.Label(bar, text="所有输出均另存为新文件，不覆盖原文件  ", bg=C_STATUS_BG,
                 fg=C_TEXT_DIM, font=("微软雅黑", 8)).pack(side=tk.RIGHT, pady=4)

    # ------------------------------------------------------------------
    # 切换工具
    # ------------------------------------------------------------------
    def select(self, idx):
        if self._current is not None:
            try:
                self._instances[self._current].pack_forget()
            except Exception:
                pass
            self._buttons[self._current].config(bg=C_SIDE_BG, fg=C_SIDE_FG)

        cls = TOOLS[idx]

        # 首次使用时才创建面板（懒加载），之后切回来会保留填写内容与日志
        if idx not in self._instances:
            try:
                self._instances[idx] = cls(self.body)
            except Exception:
                err = traceback.format_exc()
                # 输出到 stderr（无界面/自动化场景也能看到），并且不用模态弹窗以免卡死
                sys.stderr.write("[工具加载失败] %s\n%s\n" % (cls.NAME, err))
                box = tk.Frame(self.body, bg=C_MAIN_BG)
                box.pack(fill=tk.BOTH, expand=True)
                tk.Label(box, text="⚠ 工具加载失败：%s" % cls.NAME, bg=C_MAIN_BG,
                         fg="#C0392B", font=("微软雅黑", 11, "bold"), anchor="w").pack(
                    fill=tk.X, padx=12, pady=(14, 4))
                txt = tk.Text(box, wrap="none", height=20, bg="#FFF6F6", fg="#7B241C")
                txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
                txt.insert("1.0", err)
                self._instances[idx] = box
                self.lbl_name.config(text=cls.NAME)
                self.lbl_desc.config(text="加载失败，详见下方错误信息")
                self._buttons[idx].config(bg=C_SIDE_SEL, fg="#FFFFFF")
                self._current = idx
                return

        self._instances[idx].pack(fill=tk.BOTH, expand=True)
        self.lbl_name.config(text=cls.NAME)
        self.lbl_desc.config(text=cls.DESC)
        self._buttons[idx].config(bg=C_SIDE_SEL, fg="#FFFFFF")
        self._current = idx


def enable_hidpi():
    """Windows 高分屏适配：避免界面模糊。"""
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def selftest():
    """
    自检：验证打包后依赖与引擎是否正常。

    用法（在 exe 所在目录执行）：
        Excel工具箱.exe --selftest
    结果会写入当前目录的 selftest_report.txt（因为打包成无控制台程序，看不到标准输出）。
    """
    lines = ["Excel 工具箱 自检报告", "=" * 40, "版本: %s" % APP_VERSION, ""]
    ok = True

    def check(name, fn):
        nonlocal ok
        try:
            msg = fn()
            lines.append("[ OK ] %-16s %s" % (name, msg if msg else ""))
        except Exception as e:
            ok = False
            lines.append("[FAIL] %-16s %s" % (name, e))

    def _openpyxl():
        import openpyxl
        return "版本 " + openpyxl.__version__

    def _tk():
        import tkinter
        return "Tk " + str(tkinter.TkVersion)

    def _engine():
        from tools.matching import build_index, lookup
        idx = build_index([('FC2002(123)', 'FC2002(123)'), ('0650090', '0650090')])
        cases = [('123(FC2002)', 'FC2002(123)'), ('650090', '0650090'), ('7"Plate', '7inPlate')]
        idx2 = build_index([('7inPlate', '7inPlate')])
        got = lookup('123(FC2002)', idx)[0]
        got2 = lookup('650090', idx)[0]
        got3 = lookup('7"Plate', idx2)[0]
        if (got, got2, got3) != ('FC2002(123)', '0650090', '7inPlate'):
            raise RuntimeError('匹配结果异常: %r %r %r' % (got, got2, got3))
        return "10级匹配规则正常"

    def _tools():
        return "已注册：" + "、".join(tool_names())

    def _gui():
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        App(root)
        root.destroy()
        return "主窗口与全部面板可构建"

    check("openpyxl", _openpyxl)
    check("tkinter", _tk)
    check("匹配引擎", _engine)
    check("功能注册", _tools)
    check("界面构建", _gui)

    lines.append("")
    lines.append("结论: " + ("全部通过" if ok else "存在失败项"))
    report = "\n".join(lines)
    try:
        print(report)
    except Exception:
        pass
    try:
        import os
        with open(os.path.join(os.getcwd(), "selftest_report.txt"), "w", encoding="utf-8") as f:
            f.write(report)
    except Exception:
        pass
    return 0 if ok else 1


def main():
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    enable_hidpi()
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

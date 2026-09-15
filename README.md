# Excel 工具箱

把多个 Excel 处理工具合并为**单一程序**：左侧菜单切换功能，右侧显示对应面板。
新增功能只需加一个模块 + 注册一行，无需改动主程序。

---

## 一、运行

| 方式     | 操作                                                         |
| -------- | ------------------------------------------------------------ |
| 直接使用 | 双击 `dist\Excel工具箱.exe`                                  |
| 自检     | `Excel工具箱.exe --selftest`，结果写入当前目录 `selftest_report.txt` |
| 源码运行 | `python main.py`（需 Python 3.12 + openpyxl）                |

> 首次运行如被 Windows Defender / SmartScreen 拦截，属未签名程序的常见提示，选「仍要运行」。

---

## 二、当前功能

| 功能               | 说明                                                         |
| ------------------ | ------------------------------------------------------------ |
| **编号清洗匹配**   | 把写法不规范的编号自动匹配到规范写法。清洗：去空白/NBSP/零宽字符、全角转半角、`”″"` → `in`、各种横杠 `–—−` → `-`；匹配：10 级派生规则（精确 / 英寸→in / 去尾部序号 / 括号内外互换 / 去括号内容 / 去横杠 / 数字去前导零 等） |
| **关键词批量改列** | 按关键词匹配行，批量修改该行多个列的取值，内置包材类预置规则 |

两个工具都在目标列右侧插入结果列，**另存为新文件，不覆盖原文件**，原格式保留。

---

## 三、目录结构

```
ExcelToolbox/
├── main.py                 # 主入口：主窗口、左侧菜单、状态栏、自检
├── build.py                # 一键打包
├── README.md
├── tools/
│   ├── __init__.py         # ★ 功能注册表（新增功能在这里加一行）
│   ├── base.py             # 工具基类 ToolFrame（提供 log/clear_log/日志面板）
│   ├── matching.py         # 编号匹配引擎（纯逻辑，无界面依赖）
│   ├── clean_match.py      # 功能1：编号清洗匹配
│   └── keyword_edit.py     # 功能2：关键词批量改列
└── dist/
    └── Excel工具箱.exe      # 打包产物
```

---

## 四、如何新增一个功能（3 步）

### 第 1 步：新建 `tools/my_tool.py`

```python
# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog, messagebox

from .base import ToolFrame


class MyTool(ToolFrame):
    NAME = "我的新功能"        # 左侧菜单显示的名字
    DESC = "一句话说明这个功能做什么"   # 顶部显示的说明

    def __init__(self, master):
        super().__init__(master)

        # ---- 搭界面（父级一律写 self，不要写 root）----
        box = tk.LabelFrame(self, text="设置")
        box.pack(padx=10, pady=6, fill=tk.X)

        self.path = tk.StringVar()
        tk.Entry(box, textvariable=self.path, width=60).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(box, text="选择文件", command=self.pick).pack(side=tk.LEFT, padx=5)

        # ---- 日志区（用基类提供的，自带 log() / clear_log()）----
        self._create_log(self, height=10)

        btn = tk.Frame(self)
        btn.pack(pady=8)
        tk.Button(btn, text="开始处理", command=self.run, bg="#3080E0",
                  fg="white", width=12).pack(side=tk.LEFT, padx=10)
        tk.Button(btn, text="清空日志", command=self.clear_log, width=10).pack(side=tk.LEFT)

    def pick(self):
        p = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xlsm")])
        if p:
            self.path.set(p)

    def run(self):
        if not os.path.exists(self.path.get()):
            messagebox.showerror("错误", "请先选择文件")
            return
        self.clear_log()
        self.log("开始处理：" + self.path.get())
        # ... 你的业务逻辑 ...
        self.log("完成")


# 可选：本地调试用（打包时不影响）
if __name__ == "__main__":
    root = tk.Tk()
    MyTool(root).pack(fill=tk.BOTH, expand=True)
    root.mainloop()
```

### 第 2 步：在 `tools/__init__.py` 注册

```python
from .clean_match import CleanMatchTool
from .keyword_edit import KeywordEditTool
from .my_tool import MyTool          # ← 新增

TOOLS = [
    CleanMatchTool,
    KeywordEditTool,
    MyTool,                           # ← 新增，顺序即菜单顺序
]
```

### 第 3 步：重新打包

```powershell
python build.py
```

左侧菜单会自动多出「我的新功能」。

---

## 五、开发约定（重要）

| 约定                            | 说明                                                         |
| ------------------------------- | ------------------------------------------------------------ |
| **父级写 `self`**               | 工具是嵌进主窗口的 Frame，不是独立窗口。所以是 `tk.LabelFrame(self, ...)`，**不能写 `root`** |
| **不要设窗口标题/尺寸**         | `title()` / `geometry()` 由主程序统一管理                    |
| **用 `self.log()`**             | 基类已提供，自动滚动；配合 `self._create_log(...)` 创建面板  |
| **`messagebox` 可用于用户提示** | 但**不要在初始化/错误兜底里用模态弹窗**，否则自动化场景会卡住 |
| **业务逻辑与界面分离**          | 复杂算法抽到独立模块（参考 `matching.py`），便于复用与测试   |

---

## 六、打包说明

`build.py` 使用 PyInstaller 单文件模式，并排除了用不到的重量级库（numpy / pandas / scipy /
matplotlib / PIL / lxml / sqlite3 等），产物约 **12 MB**（原两个独立工具各 26.87 MB）。

如需调整：

```python
EXCLUDES = [...]      # build.py 里，确认你的新功能没用到某个库就保留排除
```

打包后**务必自检**：

```powershell
dist\Excel工具箱.exe --selftest
```

它会验证 openpyxl、tkinter、匹配引擎、功能注册、界面构建五项，并写入 `selftest_report.txt`。
（自检很重要——能发现"某个库被排除了但代码其实要用"这类问题。）

---

## 七、版本

- v2.0.0 —— 合并「编号清洗匹配」与「关键词批量改列」为单一程序，支持模块化扩展

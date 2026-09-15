# -*- coding: utf-8 -*-
"""
功能注册表 —— 左侧菜单就是按这个列表生成的。

【新增功能只需两步】
    1. 在 tools/ 下新建模块，定义 class XxxTool(ToolFrame)（照抄现有模块的写法）
    2. 在下面的 TOOLS 列表里加一行
    重新打包后，新功能就会自动出现在左侧菜单里。
"""
from .clean_match import CleanMatchTool
from .keyword_edit import KeywordEditTool
from .duo_excel_ui import DuoExcelUI

# 左侧菜单顺序 = 这个列表的顺序
TOOLS = [
    CleanMatchTool,      # 编号清洗匹配
    KeywordEditTool,     # 关键词批量改列
    DuoExcelUI,          # 多个表按照sheet名和选定字段进行合并
]


def tool_names():
    """返回所有已注册工具的名字（供界面与调试使用）。"""
    return [cls.NAME for cls in TOOLS]

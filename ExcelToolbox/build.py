# -*- coding: utf-8 -*-
"""
一键打包脚本 —— 生成单文件 exe

用法：
    python build.py

产物：dist\\Excel工具箱.exe
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "Excel工具箱"
ENTRY = "main.py"

# 明确用不到的重量级库，排除后体积更小（不影响本工具功能）
EXCLUDES = [
    "numpy", "pandas", "scipy", "matplotlib", "PIL", "IPython",
    "pytest", "setuptools", "pip", "wheel", "pydoc_data",
    "lxml", "sqlite3", "unittest", "pdb", "doctest", "test",
]


def run(cmd):
    print(">", " ".join(cmd))
    r = subprocess.run(cmd, cwd=HERE)
    return r.returncode


def main():
    # 清理旧产物
    for d in ("build", "dist"):
        p = os.path.join(HERE, d)
        if os.path.isdir(p):
            shutil.rmtree(p, ignore_errors=True)
    spec = os.path.join(HERE, APP_NAME + ".spec")
    if os.path.exists(spec):
        os.remove(spec)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onefile",          # 单文件
        "--windowed",         # 不显示控制台窗口
        "--name", APP_NAME,
    ]
    for m in EXCLUDES:
        cmd += ["--exclude-module", m]
    cmd.append(ENTRY)

    if run(cmd) != 0:
        print("\n打包失败")
        return 1

    exe = os.path.join(HERE, "dist", APP_NAME + ".exe")
    if os.path.exists(exe):
        size = os.path.getsize(exe) / 1048576
        print("\n打包完成：%s  (%.2f MB)" % (exe, size))
        print("\n下一步：")
        print("  1. 运行 %s 验证界面" % os.path.basename(exe))
        print("  2. 自检：%s --selftest  （结果写入当前目录 selftest_report.txt）" % os.path.basename(exe))
        return 0
    print("\n未找到产物 exe")
    return 1


if __name__ == "__main__":
    sys.exit(main())

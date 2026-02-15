#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本 - 用于生成 macOS .app 或 Windows .exe
"""

import os
import sys
import shutil
import subprocess
import platform

# 项目路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "达成度报告生成器")
MAIN_SCRIPT = os.path.join(BASE_DIR, "main.py")

# 应用名称
APP_NAME = "达成度报告生成器"


def build():
    """执行打包"""
    system = platform.system()
    print(f"当前系统: {system}")
    print(f"开始打包 {APP_NAME}...")

    # 确保输出目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # PyInstaller 命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--windowed",  # 无控制台窗口
        "--clean",     # 清理临时文件
        "--noconfirm", # 覆盖已有文件
        "--distpath", OUTPUT_DIR,
        "--workpath", os.path.join(BASE_DIR, "build"),
        "--specpath", BASE_DIR,
    ]

    # macOS 使用 onedir 模式生成 .app
    if system == "Darwin":
        cmd.append("--onedir")
    else:
        # Windows 使用 onefile 模式生成单个 .exe
        cmd.append("--onefile")

    # 排除不需要的包（Qt、matplotlib等大型库）
    excludes = [
        "PyQt5", "PyQt6", "PySide2", "PySide6",
        "matplotlib", "scipy", "IPython", "jupyter",
        "notebook", "nbformat", "nbconvert",
        "tornado", "zmq", "cryptography",
    ]
    for exc in excludes:
        cmd.extend(["--exclude-module", exc])

    # 添加隐藏导入（确保依赖被正确打包）
    hidden_imports = [
        "customtkinter",
        "pandas",
        "openpyxl",
        "openpyxl.chart",
        "openpyxl.chart.bar_chart",
        "openpyxl.chart.line_chart",
        "openpyxl.styles",
        "openpyxl.utils",
        "PIL",
        "PIL._tkinter_finder",
    ]

    for imp in hidden_imports:
        cmd.extend(["--hidden-import", imp])

    # 添加 customtkinter 数据文件
    try:
        import customtkinter
        ctk_path = os.path.dirname(customtkinter.__file__)
        if system == "Darwin":
            cmd.extend(["--add-data", f"{ctk_path}:customtkinter"])
        else:
            cmd.extend(["--add-data", f"{ctk_path};customtkinter"])
    except ImportError:
        print("警告: 未找到 customtkinter，请先安装")

    # 添加主脚本
    cmd.append(MAIN_SCRIPT)

    print("执行命令:")
    print(" ".join(cmd[:10]) + " ...")
    print("-" * 60)

    # 执行打包
    result = subprocess.run(cmd, cwd=BASE_DIR)

    if result.returncode == 0:
        print("-" * 60)
        print(f"打包成功！")
        print(f"输出位置: {OUTPUT_DIR}")

        if system == "Darwin":
            app_path = os.path.join(OUTPUT_DIR, f"{APP_NAME}.app")
            if os.path.exists(app_path):
                print(f"macOS 应用: {app_path}")
        elif system == "Windows":
            exe_path = os.path.join(OUTPUT_DIR, f"{APP_NAME}.exe")
            if os.path.exists(exe_path):
                print(f"Windows 程序: {exe_path}")
    else:
        print("打包失败！")
        sys.exit(1)

    # 清理 build 目录和 spec 文件
    build_dir = os.path.join(BASE_DIR, "build")
    spec_file = os.path.join(BASE_DIR, f"{APP_NAME}.spec")

    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
        print("已清理 build 目录")

    if os.path.exists(spec_file):
        os.remove(spec_file)
        print("已清理 spec 文件")


if __name__ == "__main__":
    build()

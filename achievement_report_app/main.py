#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
达成度报告生成器 - CustomTkinter桌面应用
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from core import Config, AchievementProcessor


class AchievementReportApp(ctk.CTk):
    """达成度报告生成器应用"""

    def __init__(self):
        super().__init__()

        self.config = Config()
        self.processor = AchievementProcessor(self.config)

        # 文件选择
        self.selected_files: list[str] = []
        self.output_dir: str = ""
        self.last_output_files: list[str] = []  # 最后生成的文件列表

        # 文件覆盖处理（用于线程同步）
        self._overwrite_event = threading.Event()
        self._overwrite_result: str = ""  # "overwrite", "rename", "skip"

        self._setup_window()
        self._build_ui()

    def _setup_window(self):
        """设置窗口属性"""
        self.title("达成度报告生成器")

        # 获取屏幕尺寸
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # 期望的窗口尺寸（窄而长）
        desired_width = 700
        desired_height = 850

        # 确保窗口不超过屏幕的90%
        max_width = int(screen_width * 0.9)
        max_height = int(screen_height * 0.9)

        window_width = min(desired_width, max_width)
        window_height = min(desired_height, max_height)

        # 窗口居中
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.minsize(550, 500)

        # 设置主题
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

    def _build_ui(self):
        """构建UI界面"""
        # 主容器（可滚动）
        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # === 标题区域 ===
        title_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        title_frame.pack(fill="x", pady=(0, 20))

        # 右上角说明书按钮（先放，浮动在右侧）
        ctk.CTkButton(
            title_frame,
            text="📖 说明书",
            command=self._open_manual,
            width=80,
            height=28,
            font=ctk.CTkFont(size=12),
            fg_color="#607D8B",
            hover_color="#455A64"
        ).place(relx=1.0, x=-5, y=5, anchor="ne")

        ctk.CTkLabel(
            title_frame,
            text="达成度报告生成器",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color="#1565C0"
        ).pack()

        ctk.CTkLabel(
            title_frame,
            text="从成绩单Excel生成达成度分析报告",
            font=ctk.CTkFont(size=14),
            text_color="#555555"
        ).pack()

        # === 文件选择区域 ===
        file_frame = ctk.CTkFrame(self.main_frame)
        file_frame.pack(fill="x", pady=10)

        file_header = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_header.pack(fill="x", padx=15, pady=(15, 10))

        ctk.CTkLabel(
            file_header,
            text="📄 选择成绩单文件",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(side="left")

        btn_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15)

        ctk.CTkButton(
            btn_frame,
            text="选择文件",
            command=self._pick_files,
            width=100
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame,
            text="清空",
            command=self._clear_files,
            width=80,
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90")
        ).pack(side="left", padx=(0, 10))

        self.file_count_label = ctk.CTkLabel(
            btn_frame,
            text="未选择文件",
            text_color="#555555"
        )
        self.file_count_label.pack(side="left")

        # 文件列表
        self.file_listbox = tk.Listbox(
            file_frame,
            height=5,
            font=("Microsoft YaHei", 14),  # Windows使用微软雅黑，字号加大
            selectmode=tk.SINGLE,
            bg="#F5F5F5",
            fg="#333333",  # 深色文字
            selectbackground="#2196F3",
            selectforeground="white",
            relief="flat",
            highlightthickness=1,
            highlightcolor="#2196F3"
        )
        self.file_listbox.pack(fill="x", padx=15, pady=(10, 15))

        # === 输出目录区域 ===
        output_frame = ctk.CTkFrame(self.main_frame)
        output_frame.pack(fill="x", pady=10)

        output_header = ctk.CTkFrame(output_frame, fg_color="transparent")
        output_header.pack(fill="x", padx=15, pady=15)

        ctk.CTkLabel(
            output_header,
            text="📁 输出目录",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(side="left")

        output_btn_frame = ctk.CTkFrame(output_frame, fg_color="transparent")
        output_btn_frame.pack(fill="x", padx=15, pady=(0, 15))

        ctk.CTkButton(
            output_btn_frame,
            text="选择目录",
            command=self._pick_output_dir,
            width=100
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            output_btn_frame,
            text="清空",
            command=self._clear_output_dir,
            width=60,
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90")
        ).pack(side="left", padx=(0, 10))

        self.output_dir_label = ctk.CTkLabel(
            output_btn_frame,
            text="未选择（将使用源文件所在目录）",
            text_color="#555555"
        )
        self.output_dir_label.pack(side="left")

        # === 配置参数区域 ===
        config_frame = ctk.CTkFrame(self.main_frame)
        config_frame.pack(fill="x", pady=10)

        config_header = ctk.CTkFrame(config_frame, fg_color="transparent")
        config_header.pack(fill="x", padx=15, pady=15)

        ctk.CTkLabel(
            config_header,
            text="⚙️ 配置参数",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(side="left")

        ctk.CTkButton(
            config_header,
            text="恢复默认",
            command=self._reset_config,
            width=80,
            height=26,
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90")
        ).pack(side="right")

        # 第一行参数
        row1 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row1.pack(fill="x", padx=15, pady=5)

        self.ratio1_entry = self._create_param_field(row1, "目标一占比(%)", "50")
        self.ratio2_entry = self._create_param_field(row1, "目标二占比(%)", "30")
        self.ratio3_entry = self._create_param_field(row1, "目标三占比(%)", "20")

        # 第二行参数
        row2 = ctk.CTkFrame(config_frame, fg_color="transparent")
        row2.pack(fill="x", padx=15, pady=5)

        self.regular_entry = self._create_param_field(row2, "平时成绩占比(%)", "30")
        self.final_entry = self._create_param_field(row2, "期末成绩占比(%)", "70")
        self.expectation_entry = self._create_param_field(row2, "达成度期望值", "0.6")

        self.config_status_label = ctk.CTkLabel(
            config_frame,
            text="",
            text_color="#555555"
        )
        self.config_status_label.pack(padx=15, pady=(5, 15))

        # === 按钮区域 ===
        btn_action_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btn_action_frame.pack(pady=20)

        self.generate_btn = ctk.CTkButton(
            btn_action_frame,
            text="▶ 生成报告",
            command=self._on_generate,
            height=45,
            width=150,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.generate_btn.pack(side="left", padx=(0, 15))

        self.open_dir_btn = ctk.CTkButton(
            btn_action_frame,
            text="📂 打开输出目录",
            command=self._open_output_dir,
            height=45,
            width=150,
            fg_color="#4CAF50",
            hover_color="#388E3C",
            font=ctk.CTkFont(size=14)
        )
        self.open_dir_btn.pack(side="left")

        # === 进度区域 ===
        progress_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        progress_frame.pack(fill="x", pady=10)

        self.progress_bar = ctk.CTkProgressBar(progress_frame, width=600)
        self.progress_bar.pack()
        self.progress_bar.set(0)
        self.progress_bar.pack_forget()  # 初始隐藏

        self.progress_label = ctk.CTkLabel(
            progress_frame,
            text="",
            text_color="#555555"
        )
        self.progress_label.pack(pady=5)

        self.result_label = ctk.CTkLabel(
            progress_frame,
            text="",
            font=ctk.CTkFont(size=14),
            wraplength=600,
            justify="left"
        )
        self.result_label.pack(pady=10)

        # === 版权信息 ===
        ctk.CTkLabel(
            self.main_frame,
            text="© 2025 集美大学 达成度报告生成器 v1.2 | 刘祉祁",
            font=ctk.CTkFont(size=11),
            text_color="#555555"
        ).pack(pady=(20, 5))

    def _create_param_field(self, parent, label: str, default: str) -> ctk.CTkEntry:
        """创建参数输入框"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", padx=(0, 20))

        ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(size=12)).pack(anchor="w")
        entry = ctk.CTkEntry(frame, width=100, justify="center")
        entry.insert(0, default)
        entry.pack()
        entry.bind("<KeyRelease>", lambda e: self._validate_config())

        return entry

    def _pick_files(self):
        """选择文件"""
        files = filedialog.askopenfilenames(
            title="选择成绩单Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if files:
            # 验证文件类型
            valid_files = []
            invalid_files = []
            for f in files:
                if f.lower().endswith(('.xlsx', '.xls')):
                    valid_files.append(f)
                else:
                    invalid_files.append(os.path.basename(f))

            if invalid_files:
                messagebox.showwarning(
                    "文件类型错误",
                    f"以下文件不是Excel格式，已被忽略：\n\n" + "\n".join(invalid_files)
                )

            if valid_files:
                self.selected_files = valid_files
                self._update_file_list()
            elif not invalid_files:
                # 没有选择任何文件
                pass
            else:
                # 全部都是无效文件
                self.file_count_label.configure(text="未选择有效文件")

    def _pick_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir = directory
            self.output_dir_label.configure(text=f"输出到: {directory}")

    def _clear_output_dir(self):
        """清空输出目录选择"""
        self.output_dir = ""
        self.output_dir_label.configure(text="未选择（将使用源文件所在目录）")

    def _reset_config(self):
        """恢复默认配置"""
        defaults = {
            'ratio1': '50',
            'ratio2': '30',
            'ratio3': '20',
            'regular': '30',
            'final': '70',
            'expectation': '0.6'
        }
        self.ratio1_entry.delete(0, tk.END)
        self.ratio1_entry.insert(0, defaults['ratio1'])
        self.ratio2_entry.delete(0, tk.END)
        self.ratio2_entry.insert(0, defaults['ratio2'])
        self.ratio3_entry.delete(0, tk.END)
        self.ratio3_entry.insert(0, defaults['ratio3'])
        self.regular_entry.delete(0, tk.END)
        self.regular_entry.insert(0, defaults['regular'])
        self.final_entry.delete(0, tk.END)
        self.final_entry.insert(0, defaults['final'])
        self.expectation_entry.delete(0, tk.END)
        self.expectation_entry.insert(0, defaults['expectation'])
        self._validate_config()

    def _update_file_list(self):
        """更新文件列表"""
        self.file_listbox.delete(0, tk.END)
        for idx, f in enumerate(self.selected_files, 1):
            self.file_listbox.insert(tk.END, f"  {idx}. 📄 {os.path.basename(f)}")
        self.file_count_label.configure(text=f"已选择 {len(self.selected_files)} 个文件")

    def _clear_files(self):
        """清空文件列表"""
        self.selected_files = []
        self.file_listbox.delete(0, tk.END)
        self.file_count_label.configure(text="未选择文件")

    def _validate_config(self) -> tuple[bool, str]:
        """验证配置参数

        Returns:
            (是否有效, 错误信息)
        """
        try:
            self.config.ratio_1 = int(self.ratio1_entry.get() or 0)
            self.config.ratio_2 = int(self.ratio2_entry.get() or 0)
            self.config.ratio_3 = int(self.ratio3_entry.get() or 0)
            self.config.regular_score_ratio = int(self.regular_entry.get() or 0)
            self.config.final_score_ratio = int(self.final_entry.get() or 0)
            self.config.achievement_expectation = float(self.expectation_entry.get() or 0)

            valid, error = self.config.validate()
            if not valid:
                self.config_status_label.configure(text=f"⚠️ {error}", text_color="red")
                return False, error
            else:
                self.config_status_label.configure(text="✓ 配置有效", text_color="green")
                return True, ""
        except ValueError:
            error = "请输入有效数字"
            self.config_status_label.configure(text=f"⚠️ {error}", text_color="red")
            return False, error

    def _on_generate(self):
        """点击生成按钮"""
        # 验证文件
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择成绩单文件")
            return

        # 验证配置
        valid, config_error = self._validate_config()
        if not valid:
            messagebox.showwarning("配置错误", f"请检查配置参数：\n\n{config_error}")
            return

        # 验证输出目录权限
        if self.output_dir:
            if not os.path.exists(self.output_dir):
                messagebox.showerror("错误", f"输出目录不存在：\n{self.output_dir}")
                return
            if not os.access(self.output_dir, os.W_OK):
                messagebox.showerror("错误", f"输出目录无写入权限：\n{self.output_dir}")
                return

        # 检查输入文件是否存在
        missing_files = [f for f in self.selected_files if not os.path.exists(f)]
        if missing_files:
            messagebox.showerror(
                "错误",
                f"以下文件不存在或已被移动：\n\n" + "\n".join([os.path.basename(f) for f in missing_files])
            )
            return

        # 禁用按钮，显示进度
        self.generate_btn.configure(state="disabled", text="处理中...")
        self.progress_bar.pack()
        self.progress_bar.set(0)
        self.result_label.configure(text="")

        # 在后台线程处理
        thread = threading.Thread(target=self._process_files, daemon=True)
        thread.start()

    def _get_unique_filename(self, filepath: str) -> str:
        """获取唯一的文件名，如果文件存在则添加后缀 _1, _2, ..."""
        if not os.path.exists(filepath):
            return filepath

        base, ext = os.path.splitext(filepath)
        counter = 1
        while True:
            new_path = f"{base}_{counter}{ext}"
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def _show_overwrite_dialog(self, filepath: str):
        """在主线程显示文件覆盖对话框"""
        filename = os.path.basename(filepath)
        result = messagebox.askyesnocancel(
            "文件已存在",
            f"文件 \"{filename}\" 已存在。\n\n"
            f"点击「是」覆盖原文件\n"
            f"点击「否」自动重命名（添加后缀）\n"
            f"点击「取消」跳过此文件"
        )

        if result is True:
            self._overwrite_result = "overwrite"
        elif result is False:
            self._overwrite_result = "rename"
        else:
            self._overwrite_result = "skip"

        self._overwrite_event.set()

    def _process_files(self):
        """后台处理文件"""
        success_count = 0
        fail_count = 0
        skip_count = 0
        results = []
        all_warnings = []
        output_files = []  # 记录成功生成的文件

        total_files = len(self.selected_files)

        for idx, input_file in enumerate(self.selected_files):
            filename = os.path.basename(input_file)
            try:
                name_without_ext = os.path.splitext(filename)[0]

                # 确定输出目录
                if self.output_dir:
                    output_dir = self.output_dir
                else:
                    output_dir = os.path.dirname(input_file)

                output_file = os.path.join(output_dir, f"{name_without_ext}_达成度报告.xlsx")

                # 检查文件是否存在
                if os.path.exists(output_file):
                    # 重置事件
                    self._overwrite_event.clear()
                    # 在主线程显示对话框
                    self.after(0, self._show_overwrite_dialog, output_file)
                    # 等待用户选择
                    self._overwrite_event.wait()

                    if self._overwrite_result == "skip":
                        skip_count += 1
                        results.append(f"⏭ {filename}: 已跳过（文件已存在）")
                        continue
                    elif self._overwrite_result == "rename":
                        output_file = self._get_unique_filename(output_file)

                # 设置进度回调
                def progress_callback(msg, percent, idx=idx, filename=filename):
                    file_progress = (idx + percent / 100) / total_files
                    self.progress_bar.set(file_progress)
                    self.progress_label.configure(
                        text=f"[{idx + 1}/{total_files}] {filename}: {msg}"
                    )

                self.processor.set_progress_callback(progress_callback)
                result = self.processor.process_file(input_file, output_file)

                success_count += 1
                output_files.append(output_file)  # 记录成功的输出文件
                # 显示学生数量和警告数量
                warn_count = len(result.get('warnings', []))
                output_basename = os.path.basename(output_file)
                if warn_count > 0:
                    results.append(f"✓ {output_basename}: {result['total_students']}名学生 (⚠️ {warn_count}个警告)")
                    all_warnings.extend([f"[{filename}] {w}" for w in result['warnings']])
                else:
                    results.append(f"✓ {output_basename}: {result['total_students']}名学生")

            except Exception as e:
                fail_count += 1
                error_msg = str(e).replace('\n', ' | ')  # 简化多行错误
                results.append(f"✗ {filename}: {error_msg}")

        # 保存输出文件列表
        self.last_output_files = output_files

        # 完成 - 在主线程更新UI
        self.after(0, self._on_process_complete, success_count, fail_count, results, all_warnings, skip_count)

    def _open_output_dir(self):
        """打开输出目录，如果有刚生成的文件则选中"""
        import subprocess
        import sys

        # 确定要打开的目录
        if self.last_output_files:
            # 有刚生成的文件，选中第一个
            file_to_select = self.last_output_files[0]
            target_dir = os.path.dirname(file_to_select)
        elif self.output_dir:
            # 使用选择的输出目录
            target_dir = self.output_dir
            file_to_select = None
        elif self.selected_files:
            # 使用第一个输入文件的目录
            target_dir = os.path.dirname(self.selected_files[0])
            file_to_select = None
        else:
            messagebox.showinfo("提示", "请先选择文件或输出目录")
            return

        # 检查目录是否存在
        if not os.path.exists(target_dir):
            messagebox.showerror("错误", f"目录不存在：\n{target_dir}")
            return

        try:
            if sys.platform == 'win32':
                # Windows
                if file_to_select and os.path.exists(file_to_select):
                    # 打开目录并选中文件
                    subprocess.run(['explorer', '/select,', file_to_select], check=False)
                else:
                    # 只打开目录
                    subprocess.run(['explorer', target_dir], check=False)
            elif sys.platform == 'darwin':
                # macOS
                if file_to_select and os.path.exists(file_to_select):
                    subprocess.run(['open', '-R', file_to_select], check=False)
                else:
                    subprocess.run(['open', target_dir], check=False)
            else:
                # Linux
                subprocess.run(['xdg-open', target_dir], check=False)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录：\n{str(e)}")

    def _open_manual(self):
        """打开说明书文件"""
        import subprocess
        import sys

        # 获取应用程序所在目录
        if getattr(sys, 'frozen', False):
            # 打包后的可执行文件
            app_dir = os.path.dirname(sys.executable)
        else:
            # 开发模式运行
            app_dir = os.path.dirname(os.path.abspath(__file__))

        manual_path = os.path.join(app_dir, '说明书.txt')

        if not os.path.exists(manual_path):
            messagebox.showwarning(
                "找不到说明书",
                "找不到 说明书.txt，请确认该文件位于与应用程序相同目录中。"
            )
            return

        try:
            if sys.platform == 'win32':
                os.startfile(manual_path)
            elif sys.platform == 'darwin':
                subprocess.run(['open', manual_path], check=False)
            else:
                subprocess.run(['xdg-open', manual_path], check=False)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开说明书：\n{str(e)}")

    def _on_process_complete(self, success_count: int, fail_count: int, results: list, warnings: list, skip_count: int = 0):
        """处理完成回调"""
        self.progress_bar.set(1)
        self.progress_label.configure(text="")
        self.generate_btn.configure(state="normal", text="▶ 生成报告")

        # 延迟隐藏进度条（让用户看到完成状态）
        self.after(1500, lambda: self.progress_bar.pack_forget())

        # 构建结果摘要
        summary_parts = [f"成功: {success_count}", f"失败: {fail_count}"]
        if skip_count > 0:
            summary_parts.append(f"跳过: {skip_count}")
        result_summary = f"处理完成！{', '.join(summary_parts)}\n\n"
        result_summary += "\n".join(results)

        # 如果有警告，显示警告信息
        if warnings:
            result_summary += f"\n\n⚠️ 警告信息 ({len(warnings)}条):\n"
            # 只显示前5条警告，避免信息过多
            for w in warnings[:5]:
                result_summary += f"  • {w}\n"
            if len(warnings) > 5:
                result_summary += f"  ... 还有 {len(warnings) - 5} 条警告"

        color = "green" if fail_count == 0 and len(warnings) == 0 else ("#CC7000" if fail_count == 0 else "red")
        self.result_label.configure(text=result_summary, text_color=color)

        # 弹出完成提示
        if fail_count == 0 and success_count > 0:
            if skip_count > 0:
                messagebox.showinfo("完成", f"成功处理 {success_count} 个文件，跳过 {skip_count} 个文件！")
            else:
                messagebox.showinfo("完成", f"成功处理 {success_count} 个文件！")
        elif fail_count == 0 and success_count == 0 and skip_count > 0:
            messagebox.showinfo("完成", f"所有 {skip_count} 个文件均已跳过")


def main():
    app = AchievementReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()

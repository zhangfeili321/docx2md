#!/usr/bin/env python3
"""
DOCX to Markdown Converter - Tkinter GUI Application
DOCX转Markdown转换器 - Tkinter图形界面版本（内置Python，无需安装额外依赖）
"""

import os
import sys
import threading
from pathlib import Path
from typing import Optional

# 尝试导入 tkinter（Python内置）
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    print("Error: tkinter not available in this Python environment")
    sys.exit(1)


class ConversionThread(threading.Thread):
    """后台转换线程"""
    def __init__(self, input_path: str, output_path: str, config: dict, callback):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.config = config
        self.callback = callback
        self.error = None
        self.content = None  # 存储转换后的内容
    
    def run(self):
        try:
            from enhanced_convert import DocxToMarkdownConverter, ProcessingConfig
            
            config = ProcessingConfig(
                enable_table=self.config.get('enable_table', True),
                fix_chinese_spacing=self.config.get('fix_chinese_spacing', True),
                remove_page_numbers=self.config.get('remove_page_numbers', True),
                merge_paragraphs=self.config.get('merge_paragraphs', False),
                image_enabled=self.config.get('image_enabled', True),
                image_to_base64=self.config.get('image_to_base64', False),
            )
            
            self.callback('progress', 20)
            
            converter = DocxToMarkdownConverter(config)
            self.callback('progress', 50)
            
            content, metadata = converter.convert(self.input_path, self.output_path)
            self.callback('progress', 80)
            
            if "error" in metadata:
                self.callback('error', metadata["error"])
                return
            
            self.content = content  # 保存内容
            
            self.callback('progress', 100)
            self.callback('finished', {'metadata': metadata, 'content': content})
            
        except Exception as e:
            self.callback('error', str(e))


class BatchConversionThread(threading.Thread):
    """批量转换线程"""
    def __init__(self, file_list: list, config: dict, callback):
        super().__init__()
        self.file_list = file_list
        self.config = config
        self.callback = callback
        self.results = []  # [(input, output, success, metadata), ...]
    
    def run(self):
        try:
            from enhanced_convert import DocxToMarkdownConverter, ProcessingConfig
            
            config = ProcessingConfig(
                enable_table=self.config.get('enable_table', True),
                fix_chinese_spacing=self.config.get('fix_chinese_spacing', True),
                remove_page_numbers=self.config.get('remove_page_numbers', True),
                merge_paragraphs=self.config.get('merge_paragraphs', False),
                image_enabled=self.config.get('image_enabled', True),
                image_to_base64=self.config.get('image_to_base64', False),
            )
            
            total = len(self.file_list)
            
            for i, input_path in enumerate(self.file_list):
                # 计算进度
                base_progress = (i * 100) // total
                self.callback('progress', base_progress)
                self.callback('status', f"正在转换: {os.path.basename(input_path)} ({i+1}/{total})")
                
                # 生成输出路径
                base_name = os.path.splitext(os.path.basename(input_path))[0]
                dir_name = os.path.dirname(input_path)
                output_path = os.path.join(dir_name, f"{base_name}.md")
                
                try:
                    converter = DocxToMarkdownConverter(config)
                    content, metadata = converter.convert(input_path, output_path)
                    
                    if "error" in metadata:
                        self.results.append((input_path, output_path, False, metadata))
                    else:
                        self.results.append((input_path, output_path, True, metadata))
                except Exception as e:
                    self.results.append((input_path, output_path, False, {"error": str(e)}))
            
            self.callback('progress', 100)
            self.callback('finished', {'results': self.results})
            
        except Exception as e:
            self.callback('error', str(e))


class DocxToMarkdownGUI:
    """主窗口类"""
    
    def __init__(self, root):
        self.root = root
        self.input_file: Optional[str] = None
        self.output_file: Optional[str] = None
        self.input_files: list = []  # 批量模式文件列表
        self.batch_mode: bool = False
        self.conversion_thread: Optional[threading.Thread] = None
        self.setup_ui()
    
    def setup_ui(self):
        self.root.title("DOCX → Markdown 转换器")
        self.root.geometry("600x650")
        self.root.resizable(True, True)
        self.root.minsize(550, 550)
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置进度条颜色为青色 80%透明度
        style.configure("Cyan.Horizontal.TProgressbar", 
                        background='#4a90d9',  # 青色
                        troughcolor='#25253a',  # 进度条背景
                        lightcolor='#6ab0e8',   # 浅青色
                        darkcolor='#3a80c9')     # 深青色
        
        # 配置颜色
        bg_color = '#1a1a2e'
        fg_color = '#ffffff'
        accent_color = '#4a90d9'
        success_color = '#52c41a'
        warning_color = '#faad14'
        error_color = '#ff4d4f'
        
        self.root.configure(bg=bg_color)
        
        # ===== 主容器 =====
        main_frame = tk.Frame(self.root, bg=bg_color)
        main_frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        # ===== 标题 =====
        title_label = tk.Label(
            main_frame,
            text="DOCX → Markdown 转换器",
            font=('Microsoft YaHei', 16, 'bold'),
            fg=fg_color,
            bg=bg_color
        )
        title_label.pack(pady=(0, 10))
        
        # ===== 文件拖放区域 =====
        drop_frame = tk.Frame(
            main_frame,
            bg='#25253a',
            highlightthickness=2,
            highlightbackground='#4a4a6a',
            highlightcolor=accent_color,
            cursor='hand2'
        )
        drop_frame.pack(fill='x', pady=(0, 5), ipady=15)
        
        self.drop_label = tk.Label(
            drop_frame,
            text="拖放 DOCX 文件到此处 或 点击选择",
            font=('Microsoft YaHei', 10),
            fg='#b4b4c7',
            bg='#25253a',
            justify='center'
        )
        self.drop_label.pack(expand=True, pady=10)
        
        drop_frame.bind('<Button-1>', lambda e: self.select_file())
        
        # ===== 文件路径 =====
        file_frame = tk.Frame(main_frame, bg=bg_color)
        file_frame.pack(fill='x', pady=3)
        
        tk.Label(file_frame, text="已选择:", font=('Microsoft YaHei', 9), fg='#b4b4c7', bg=bg_color).pack(side='left')
        
        self.file_path_label = tk.Label(file_frame, text="未选择文件", font=('Microsoft YaHei', 9), fg='#888888', bg=bg_color)
        self.file_path_label.pack(side='left', fill='x', expand=True, padx=(5, 0))
        
        self.clear_btn = tk.Button(file_frame, text="清除", font=('Microsoft YaHei', 8), fg=error_color, bg=bg_color, relief='flat', cursor='hand2', command=self.clear_file, state='disabled')
        self.clear_btn.pack(side='right')
        
        # ===== 输出路径 =====
        output_frame = tk.Frame(main_frame, bg=bg_color)
        output_frame.pack(fill='x', pady=3)
        
        tk.Label(output_frame, text="输出:", font=('Microsoft YaHei', 9), fg='#b4b4c7', bg=bg_color).pack(side='left')
        
        self.output_label = tk.Label(output_frame, text="(自动生成)", font=('Microsoft YaHei', 9), fg='#888888', bg=bg_color)
        self.output_label.pack(side='left', padx=(5, 0), fill='x', expand=True)
        
        # ===== 选项 =====
        options_frame = tk.Frame(main_frame, bg=bg_color)
        options_frame.pack(fill='x', pady=5)
        
        self.enable_table_var = tk.BooleanVar(value=True)
        self.enable_chinese_fix_var = tk.BooleanVar(value=True)
        self.remove_page_numbers_var = tk.BooleanVar(value=True)
        self.merge_paragraphs_var = tk.BooleanVar(value=False)
        self.enable_subheading_var = tk.BooleanVar(value=True)
        self.enable_image_var = tk.BooleanVar(value=True)
        self.image_base64_var = tk.BooleanVar(value=False)
        self.batch_mode_var = tk.BooleanVar(value=False)
        
        def small_checkbox(parent, text, var):
            cb = tk.Checkbutton(parent, text=text, variable=var, font=('Microsoft YaHei', 9), fg=fg_color, bg=bg_color, selectcolor='#25253a', activebackground=bg_color, activeforeground=fg_color, cursor='hand2', padx=8)
            cb.pack(side='left')
        
        small_checkbox(options_frame, "表格", self.enable_table_var)
        small_checkbox(options_frame, "中文间距", self.enable_chinese_fix_var)
        small_checkbox(options_frame, "移除页码", self.remove_page_numbers_var)
        small_checkbox(options_frame, '段落合并', self.merge_paragraphs_var)
        small_checkbox(options_frame, '小标题检测', self.enable_subheading_var)
        
        # 图片选项单独放一行
        img_frame = tk.Frame(main_frame, bg=bg_color)
        img_frame.pack(fill='x', pady=2)
        small_checkbox(img_frame, '图片提取', self.enable_image_var)
        small_checkbox(img_frame, '图片Base64', self.image_base64_var)
        
        # 批量模式单独放一行
        batch_frame = tk.Frame(main_frame, bg=bg_color)
        batch_frame.pack(fill='x', pady=2)
        
        batch_cb = tk.Checkbutton(batch_frame, text="□ 批量模式", variable=self.batch_mode_var, font=('Microsoft YaHei', 9, 'bold'), fg='#4a90d9', bg=bg_color, selectcolor='#25253a', activebackground=bg_color, activeforeground='#4a90d9', cursor='hand2', padx=8, command=self.on_batch_mode_changed)
        batch_cb.pack(side='left')
        
        # ===== 进度条 =====
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100, style="Cyan.Horizontal.TProgressbar")
        self.progress_bar.pack(fill='x', pady=5)
        
        # ===== 转换按钮 =====
        self.convert_btn = tk.Button(
            main_frame, text="🚀 开始转换", font=('Microsoft YaHei', 12, 'bold'),
            fg=fg_color, bg=accent_color, activebackground='#3a80c9', activeforeground=fg_color,
            relief='flat', cursor='hand2', command=self.start_conversion, state='disabled'
        )
        self.convert_btn.pack(fill='x', pady=5, ipady=6)
        
        # ===== 分隔线 =====
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=8)
        
        # ===== 结果区域 =====
        result_label = tk.Label(main_frame, text="转换结果", font=('Microsoft YaHei', 11, 'bold'), fg='#b4b4c7', bg=bg_color)
        result_label.pack(anchor='w')
        
        result_frame = tk.Frame(main_frame, bg='#25253a', relief='solid', bd=1)
        result_frame.pack(fill='both', expand=True, pady=5)
        
        # 状态和统计行
        top_row = tk.Frame(result_frame, bg='#25253a')
        top_row.pack(fill='x', pady=5)
        
        self.status_label = tk.Label(top_row, text="等待转换...", font=('Microsoft YaHei', 11, 'bold'), fg='#b4b4c7', bg='#25253a')
        self.status_label.pack(side='left', padx=10)
        
        # 统计
        stats_frame = tk.Frame(top_row, bg='#25253a')
        stats_frame.pack(side='right', padx=10)
        
        tk.Label(stats_frame, text="📑:", font=('Segoe UI Emoji', 10), fg='#4a90d9', bg='#25253a').pack(side='left')
        self.heading_value = tk.Label(stats_frame, text="-", font=('Microsoft YaHei', 10, 'bold'), fg=fg_color, bg='#25253a')
        self.heading_value.pack(side='left', padx=(0, 8))
        
        tk.Label(stats_frame, text="📄:", font=('Segoe UI Emoji', 10), fg='#4a90d9', bg='#25253a').pack(side='left')
        self.paragraph_value = tk.Label(stats_frame, text="-", font=('Microsoft YaHei', 10, 'bold'), fg=fg_color, bg='#25253a')
        self.paragraph_value.pack(side='left', padx=(0, 8))
        
        tk.Label(stats_frame, text="📊:", font=('Segoe UI Emoji', 10), fg='#4a90d9', bg='#25253a').pack(side='left')
        self.table_value = tk.Label(stats_frame, text="-", font=('Microsoft YaHei', 10, 'bold'), fg=fg_color, bg='#25253a')
        self.table_value.pack(side='left', padx=(0, 8))
        
        tk.Label(stats_frame, text="H:", font=('Microsoft YaHei', 10), fg='#4a90d9', bg='#25253a').pack(side='left')
        self.level_value = tk.Label(stats_frame, text="-", font=('Microsoft YaHei', 10, 'bold'), fg=fg_color, bg='#25253a')
        self.level_value.pack(side='left')
        
        # 验证结果
        self.validation_label = tk.Label(result_frame, text="", font=('Microsoft YaHei', 10), bg='#25253a')
        self.validation_label.pack(anchor='w', padx=10, pady=2)
        
        # 操作按钮
        btn_frame = tk.Frame(result_frame, bg='#25253a')
        btn_frame.pack(pady=3)
        
        self.open_file_btn = tk.Button(btn_frame, text="📂 打开文件", font=('Microsoft YaHei', 9), fg=accent_color, bg=bg_color, activebackground=accent_color, activeforeground=fg_color, relief='flat', cursor='hand2', command=self.open_output_file, state='disabled')
        self.open_file_btn.pack(side='left', padx=5)
        
        self.open_folder_btn = tk.Button(btn_frame, text="📁 打开文件夹", font=('Microsoft YaHei', 9), fg=accent_color, bg=bg_color, activebackground=accent_color, activeforeground=fg_color, relief='flat', cursor='hand2', command=self.open_output_folder, state='disabled')
        self.open_folder_btn.pack(side='left', padx=5)
        
        # ===== 预览区域（固定高度） =====
        preview_label = tk.Label(result_frame, text="内容预览:", font=('Microsoft YaHei', 9), fg='#888888', bg='#25253a')
        preview_label.pack(anchor='w', padx=10, pady=(5, 2))
        
        # 预览文本框（固定高度，有滚动条）
        preview_container = tk.Frame(result_frame, bg='#1a1a2e')
        preview_container.pack(fill='both', expand=True, padx=10, pady=(0, 8))
        
        scrollbar = tk.Scrollbar(preview_container)
        scrollbar.pack(side='right', fill='y')
        
        self.preview_text = tk.Text(
            preview_container,
            font=('Consolas', 9),
            bg='#1a1a2e',
            fg='#d4d4d4',
            insertbackground='#d4d4d4',
            relief='flat',
            wrap='word',
            state='disabled',
            yscrollcommand=scrollbar.set
        )
        self.preview_text.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.preview_text.yview)
    
    def select_file(self):
        if self.batch_mode_var.get():
            # 批量模式：选择多个文件
            file_paths = filedialog.askopenfilenames(
                title="选择DOCX文件（批量）",
                filetypes=[("DOCX文件", "*.docx"), ("所有文件", "*.*")]
            )
            if file_paths:
                self.set_input_files(list(file_paths))
        else:
            # 单文件模式
            file_path = filedialog.askopenfilename(
                title="选择DOCX文件",
                filetypes=[("DOCX文件", "*.docx"), ("所有文件", "*.*")]
            )
            if file_path:
                self.set_input_file(file_path)
    
    def set_input_file(self, file_path: str):
        self.input_file = file_path
        self.input_files = []
        self.batch_mode = False
        basename = os.path.basename(file_path)
        self.file_path_label.config(text=basename, fg='#ffffff')
        self.clear_btn.config(state='normal')
        self.convert_btn.config(state='normal')
        
        # 自动设置输出路径（同名.md）
        base_name = os.path.splitext(basename)[0]
        dir_name = os.path.dirname(file_path)
        default_output = os.path.join(dir_name, f"{base_name}.md")
        self.output_file = default_output
        self.output_label.config(text=default_output, fg='#b4b4c7')
    
    def set_input_files(self, file_paths: list):
        """设置批量文件列表"""
        self.input_files = file_paths
        self.input_file = None
        self.batch_mode = True
        
        count = len(file_paths)
        self.file_path_label.config(text=f"已选择 {count} 个文件", fg='#4a90d9')
        self.clear_btn.config(state='normal')
        self.convert_btn.config(state='normal')
        
        # 显示输出目录
        if file_paths:
            dir_name = os.path.dirname(file_paths[0])
            self.output_file = dir_name
            self.output_label.config(text=f"输出到: {dir_name}", fg='#4a90d9')
    
    def on_batch_mode_changed(self):
        """批量模式切换"""
        self.batch_mode = self.batch_mode_var.get()
        if self.batch_mode:
            self.drop_label.config(text="点击选择多个 DOCX 文件（批量模式）")
        else:
            self.drop_label.config(text="拖放 DOCX 文件到此处 或 点击选择")
            # 清除批量文件列表
            self.input_files = []
    
    def clear_file(self):
        self.input_file = None
        self.output_file = None
        self.input_files = []
        self.batch_mode = False
        self.file_path_label.config(text="未选择文件", fg='#888888')
        self.output_label.config(text="(自动生成)", fg='#888888')
        self.clear_btn.config(state='disabled')
        self.convert_btn.config(state='disabled')
        self.reset_result_display()
    
    def start_conversion(self):
        if self.batch_mode_var.get() and self.input_files:
            self._start_batch_conversion()
        elif self.input_file:
            self._start_single_conversion()
        else:
            messagebox.showwarning("警告", "请先选择DOCX文件")
            return
    
    def _start_single_conversion(self):
        """开始单个文件转换"""
        # 确保输出路径已设置
        if not self.output_file:
            base_name = os.path.splitext(os.path.basename(self.input_file))[0]
            dir_name = os.path.dirname(self.input_file)
            self.output_file = os.path.join(dir_name, f"{base_name}.md")
        
        output_path = self.output_file
        
        # 禁用按钮
        self.convert_btn.config(state='disabled', text='转换中...')
        
        # 收集配置
        config = {
            'enable_table': self.enable_table_var.get(),
            'fix_chinese_spacing': self.enable_chinese_fix_var.get(),
            'remove_page_numbers': self.remove_page_numbers_var.get(),
            'merge_paragraphs': self.merge_paragraphs_var.get(),
            'image_enabled': self.enable_image_var.get(),
            'image_to_base64': self.image_base64_var.get(),
        }
        
        # 启动转换线程
        self.conversion_thread = ConversionThread(
            self.input_file,
            output_path,
            config,
            self.on_conversion_callback
        )
        self.conversion_thread.start()
    
    def _start_batch_conversion(self):
        """开始批量转换"""
        # 禁用按钮
        self.convert_btn.config(state='disabled', text='批量转换中...')
        
        # 收集配置
        config = {
            'enable_table': self.enable_table_var.get(),
            'fix_chinese_spacing': self.enable_chinese_fix_var.get(),
            'remove_page_numbers': self.remove_page_numbers_var.get(),
            'merge_paragraphs': self.merge_paragraphs_var.get(),
            'image_enabled': self.enable_image_var.get(),
            'image_to_base64': self.image_base64_var.get(),
        }
        
        # 启动批量转换线程
        self.conversion_thread = BatchConversionThread(
            self.input_files,
            config,
            self.on_batch_conversion_callback
        )
        self.conversion_thread.start()
    
    def on_batch_conversion_callback(self, event_type: str, data):
        """批量转换回调"""
        if event_type == 'progress':
            self.progress_var.set(data)
            self.root.update_idletasks()
        
        elif event_type == 'status':
            self.status_label.config(text=data, fg='#b4b4c7')
            self.root.update_idletasks()
        
        elif event_type == 'finished':
            self.progress_var.set(100)
            self.convert_btn.config(state='normal', text='🚀 开始转换')
            
            results = data.get('results', [])
            total = len(results)
            success = sum(1 for r in results if r[2])
            
            # 更新状态
            self.status_label.config(text=f"✅ 批量转换完成！", fg='#52c41a')
            
            # 更新统计
            self.heading_value.config(text=f"{success}/{total}")
            self.paragraph_value.config(text="-")
            self.table_value.config(text="-")
            self.level_value.config(text="-")
            
            # 显示结果摘要
            summary = f"成功: {success}, 失败: {total - success}"
            self.validation_label.config(text=summary, fg='#52c41a' if success == total else '#faad14')
            
            # 预览区域显示文件列表
            if hasattr(self, 'preview_text'):
                self.preview_text.config(state='normal')
                self.preview_text.delete('1.0', 'end')
                preview_lines = [f"批量转换结果 ({success}/{total} 成功):\n"]
                for input_path, output_path, success, metadata in results:
                    status = "✅" if success else "❌"
                    preview_lines.append(f"{status} {os.path.basename(input_path)}")
                preview_text = '\n'.join(preview_lines[:50])  # 最多显示50个
                self.preview_text.insert('1.0', preview_text)
                self.preview_text.config(state='disabled')
            
            # 启用按钮
            self.open_file_btn.config(state='disabled')
            self.open_folder_btn.config(state='normal')
        
        elif event_type == 'error':
            self.progress_var.set(0)
            self.convert_btn.config(state='normal', text='🚀 开始转换')
            self.status_label.config(text="❌ 批量转换失败", fg='#ff4d4f')
            self.validation_label.config(text=data, fg='#ff4d4f')
            self.open_file_btn.config(state='disabled')
            self.open_folder_btn.config(state='disabled')
    
    def on_conversion_callback(self, event_type: str, data):
        if event_type == 'progress':
            self.progress_var.set(data)
            self.root.update_idletasks()
        
        elif event_type == 'finished':
            self.progress_var.set(100)
            self.convert_btn.config(state='normal', text='🚀 开始转换')
            
            # 解析返回的数据
            metadata = data.get('metadata', {})
            content = data.get('content', '')
            
            # 更新状态
            self.status_label.config(text="✅ 转换完成！", fg='#52c41a')
            
            # 更新统计
            self.heading_value.config(text=str(metadata.get('md_stats', {}).get('headings', 0)))
            self.paragraph_value.config(text=str(metadata.get('docx_stats', {}).get('paragraphs', 0)))
            self.table_value.config(text=str(metadata.get('docx_stats', {}).get('tables', 0)))
            self.level_value.config(text=f"H{metadata.get('max_heading_level', '?')}")
            
            # 验证结果
            validation = metadata.get('validation', {})
            if validation.get('is_valid', False):
                self.validation_label.config(text="✅ 验证通过", fg='#52c41a')
            else:
                issues = validation.get('issues', [])
                if issues:
                    self.validation_label.config(text=f"⚠️ {issues[0]}", fg='#faad14')
            
            # 显示转换结果预览
            if content and hasattr(self, 'preview_text'):
                self.preview_text.config(state='normal')
                self.preview_text.delete('1.0', 'end')
                # 显示前2000个字符
                preview_content = content[:2000] + ('...' if len(content) > 2000 else '')
                self.preview_text.insert('1.0', preview_content)
                self.preview_text.config(state='disabled')
            
            # 启用按钮
            self.open_file_btn.config(state='normal')
            self.open_folder_btn.config(state='normal')
        
        elif event_type == 'error':
            self.progress_var.set(0)
            self.convert_btn.config(state='normal', text='🚀 开始转换')
            self.status_label.config(text="❌ 转换失败", fg='#ff4d4f')
            self.validation_label.config(text=data, fg='#ff4d4f')
            self.open_file_btn.config(state='disabled')
            self.open_folder_btn.config(state='disabled')
    
    def reset_result_display(self):
        self.status_label.config(text="等待转换...", fg='#b4b4c7')
        self.heading_value.config(text="-")
        self.paragraph_value.config(text="-")
        self.table_value.config(text="-")
        self.level_value.config(text="-")
        self.validation_label.config(text="")
        self.progress_var.set(0)
        self.open_file_btn.config(state='disabled')
        self.open_folder_btn.config(state='disabled')
        
        # 清除预览内容
        if hasattr(self, 'preview_text'):
            self.preview_text.config(state='normal')
            self.preview_text.delete('1.0', 'end')
            self.preview_text.config(state='disabled')
    
    def open_output_file(self):
        if self.output_file and os.path.exists(self.output_file):
            os.startfile(self.output_file)
    
    def open_output_folder(self):
        if self.output_file and os.path.exists(self.output_file):
            folder = os.path.dirname(self.output_file)
            os.startfile(folder)


def main():
    root = tk.Tk()
    app = DocxToMarkdownGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

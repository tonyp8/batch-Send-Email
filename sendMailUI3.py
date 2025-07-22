# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, Frame, Label, Entry, Button, Checkbutton, StringVar, BooleanVar
import threading
import time
import sys
import os
import io
import json
import re
from datetime import datetime

from sendMailMain import *
import sendMailCore

class ColoredConsoleRedirector(io.StringIO):
    """重定向控制台输出到文本区域，支持颜色标记"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.color_map = {
            "INFO": "#4EC9B0",    # 青绿色
            "WARNING": "#d7ba7d", # 金色
            "ERROR": "#f48771",   # 红色
            "DEBUG": "#d4d4d4",   # 默认灰色
            "DEFAULT": "#9cdcfe"  # 浅蓝色
        }
    
    def write(self, message):
        # 解析日志等级
        color = self.color_map["DEFAULT"]
        if message.startswith('[INFO]'):
            color = self.color_map["INFO"]
        elif message.startswith('[WARNING]'):
            color = self.color_map["WARNING"]
        elif message.startswith('[ERROR]'):
            color = self.color_map["ERROR"]
        elif message.startswith('[DEBUG]'):
            color = self.color_map["DEBUG"]
        
        # 添加带颜色的文本
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, message, color)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')
        self.text_widget.update_idletasks()
    
    def flush(self):
        pass

class LogFileWriter:
    def __init__(self):
        # 创建log目录
        self.log_dir = "./log"
        os.makedirs(self.log_dir, exist_ok=True)
        
        # 生成日志文件名
        self.log_file = os.path.join(self.log_dir, f"mail_sender_{datetime.now().strftime('%Y%m%d')}.log")
        
        # 打开日志文件
        self.file = open(self.log_file, 'a', encoding='utf-8')
        self.write("[INFO] 日志文件已创建: {}\n".format(self.log_file))
    
    def write(self, message):
        # 只在消息开头添加时间戳，避免重复添加
        if not hasattr(self, '_last_char') or self._last_char == '\n':
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.file.write(f"[{timestamp}] ")
        
        # 记录消息的最后一个字符
        self._last_char = message[-1] if message else ''
        
        self.file.write(message)
        self.file.flush()
    
    def flush(self):
        self.file.flush()
    
    def close(self):
        self.file.close()

class DoubleWriter:
    def __init__(self, console_writer, log_writer):
        self.console_writer = console_writer
        self.log_writer = log_writer
    
    def write(self, message):
        self.console_writer.write(message)  # UI显示原始消息
        self.log_writer.write(message)      # 日志记录带时间戳的消息
    
    def flush(self):
        self.console_writer.flush()
        self.log_writer.flush()
        
class EmailSenderUI:
    def __init__(self, root):
        self.root = root
        self.root.title("邮件批量发送系统")
        self.root.geometry("1100x750")
        self.root.configure(bg="#f0f0f0")
        
        # 全局停止事件
        self.stop_event = threading.Event()
        self.sending_thread = None
        self.config = None
        
        # 创建主框架
        self.create_widgets()

        # 初始化日志系统
        self.setup_logging()
        
        # 加载配置
        self.load_config()
        
        # 重定向控制台输出
        #self.redirect_console_output()
        
        # 绑定退出事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

        # 添加表格修改锁
        self.sheet_lock = threading.Lock()

    def setup_logging(self):
        """初始化日志记录系统"""
        # 创建日志写入器
        self.log_writer = LogFileWriter()
        
        # 创建控制台输出重定向器
        self.console_redirector = ColoredConsoleRedirector(self.console_output)
        
        # 创建双重写入器
        self.double_writer = DoubleWriter(self.console_redirector, self.log_writer)
        
        # 重定向标准输出和错误输出
        sys.stdout = self.double_writer
        sys.stderr = self.double_writer
        print('[INFO] 欢迎使用达人邮件发送系统!')
    
    def redirect_console_output(self):#已弃用
        """重定向控制台输出到文本区域和日志文件"""
        
        print('[INFO] 欢迎使用达人邮件发送系统!')
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架布局
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧控制面板
        left_frame = tk.LabelFrame(main_frame, text="设置面板", bg="#f0f0f0", padx=10, pady=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 右侧控制台区域
        right_frame = tk.LabelFrame(main_frame, text="控制台输出", bg="#f0f0f0", padx=10, pady=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 创建控制台输出区域
        self.console_output = scrolledtext.ScrolledText(
            right_frame, wrap=tk.WORD, width=80, height=35, bg="#1e1e1e"
        )
        self.console_output.pack(fill=tk.BOTH, expand=True)

        
        # 配置文本颜色标签
        self.console_output.tag_configure("#4EC9B0", foreground="#4EC9B0")  # INFO
        self.console_output.tag_configure("#d7ba7d", foreground="#d7ba7d")  # WARNING
        self.console_output.tag_configure("#9cdcfe", foreground="#9cdcfe")  # ERROR
        self.console_output.tag_configure("#f48771", foreground="#f48771")  # DEBUG
        self.console_output.tag_configure("#d4d4d4", foreground="#d4d4d4")  # DEFAULT
        
        # 创建控制按钮
        self.create_control_buttons(left_frame)
        
        # 创建配置区域
        self.create_config_sections(left_frame)
    
    def create_control_buttons(self, parent):
        """创建控制按钮"""
        btn_frame = tk.Frame(parent, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.start_btn = tk.Button(
            btn_frame, text="开始发送", width=12, height=2, 
            bg="#4CAF50", fg="white", font=("Arial", 10, "bold"),
            command=self.start_sending
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = tk.Button(
            btn_frame, text="终止发送", width=12, height=2, 
            bg="#F44336", fg="white", font=("Arial", 10, "bold"),
            command=self.stop_sending, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.preview_btn = tk.Button(
            btn_frame, text="预览模板", width=12, height=2,
            bg="#2196F3", fg="white", font=("Arial", 10, "bold"),
            command=self.preview_template
        )
        self.preview_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = tk.Button(
            btn_frame, text="保存配置", width=12, height=2,
            bg="#FF9800", fg="white", font=("Arial", 10, "bold"),
            command=self.save_config
        )
        self.save_btn.pack(side=tk.LEFT, padx=5)
    
    def create_config_sections(self, parent):
        """创建整合的设置区域"""
        # 创建滚动区域
        canvas = tk.Canvas(parent, bg="#f0f0f0", highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f0f0f0")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 基本设置
        basic_frame = tk.LabelFrame(scrollable_frame, text="基本设置", bg="#f0f0f0", padx=10, pady=10)
        basic_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 发送间隔设置
        tk.Label(basic_frame, text="发送间隔(秒):", bg="#f0f0f0", anchor="w").grid(row=0, column=0, sticky="w", pady=5)
        self.interval_var = tk.StringVar()
        interval_entry = tk.Entry(basic_frame, textvariable=self.interval_var, width=10)
        interval_entry.grid(row=0, column=1, sticky="w", pady=5)
        
        # 交错发送设置
        self.staggered_var = tk.BooleanVar()
        self.staggered_cb = tk.Checkbutton(
            basic_frame, text="关闭交错发送(多邮箱时有效)", 
            variable=self.staggered_var, bg="#f0f0f0", command=self.toggle_staggered
        )
        self.staggered_cb.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        
        # 单邮箱发送设置
        self.single_sender_enabled = tk.BooleanVar()
        single_sender_cb = tk.Checkbutton(
            basic_frame, text="启用单邮箱发送", 
            variable=self.single_sender_enabled, bg="#f0f0f0",
            command=self.toggle_single_sender
        )
        single_sender_cb.grid(row=2, column=0, sticky="w", pady=5)
        
        tk.Label(basic_frame, text="默认发送邮箱:", bg="#f0f0f0").grid(row=3, column=0, sticky="w", pady=5)
        self.single_sender_var = tk.StringVar()
        self.sender_combobox = ttk.Combobox(basic_frame, textvariable=self.single_sender_var, width=25)
        self.sender_combobox.grid(row=3, column=1, sticky="w", pady=5)
        
        # 筛选设置
        filter_frame = tk.LabelFrame(scrollable_frame, text="筛选设置", bg="#f0f0f0", padx=10, pady=10)
        filter_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 筛选模式
        tk.Label(filter_frame, text="筛选模式:", bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)
        self.filter_mode_var = tk.StringVar()
        filter_mode_menu = ttk.Combobox(filter_frame, textvariable=self.filter_mode_var, width=15, state="readonly")
        filter_mode_menu['values'] = ('匹配负责人', '排除负责人')
        filter_mode_menu.grid(row=0, column=1, sticky="w", pady=5)
        
        # 负责人名称
        tk.Label(filter_frame, text="负责人名称:", bg="#f0f0f0").grid(row=1, column=0, sticky="w", pady=5)
        self.charger_name_var = tk.StringVar()
        charger_name_entry = tk.Entry(filter_frame, textvariable=self.charger_name_var, width=20)
        charger_name_entry.grid(row=1, column=1, sticky="w", pady=5)
        
        # 署名设置
        sign_frame = tk.LabelFrame(scrollable_frame, text="署名设置", bg="#f0f0f0", padx=10, pady=10)
        sign_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 署名模式
        tk.Label(sign_frame, text="署名模式:", bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)
        self.sign_mode_var = tk.StringVar()
        sign_mode_menu = ttk.Combobox(sign_frame, textvariable=self.sign_mode_var, width=15, state="readonly")
        sign_mode_menu['values'] = ('固定署名', '使用账户配置')
        sign_mode_menu.grid(row=0, column=1, sticky="w", pady=5)
        sign_mode_menu.bind("<<ComboboxSelected>>", self.toggle_sign_mode)
        
        # 固定署名
        tk.Label(sign_frame, text="固定署名:", bg="#f0f0f0").grid(row=1, column=0, sticky="w", pady=5)
        self.fixed_sign_var = tk.StringVar()
        self.fixed_sign_entry = tk.Entry(sign_frame, textvariable=self.fixed_sign_var, width=20)
        self.fixed_sign_entry.grid(row=1, column=1, sticky="w", pady=5)
        
        # 模板设置
        template_frame = tk.LabelFrame(scrollable_frame, text="邮件模板设置", bg="#f0f0f0", padx=10, pady=10)
        template_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 语言模板表格
        columns = ("language", "path")
        self.template_tree = ttk.Treeview(template_frame, columns=columns, show="headings", height=5)
        
        # 设置列标题
        self.template_tree.heading("language", text="语言")
        self.template_tree.heading("path", text="模板路径")
        
        # 设置列宽
        self.template_tree.column("language", width=80, anchor="center")
        self.template_tree.column("path", width=300)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(template_frame, orient="vertical", command=self.template_tree.yview)
        self.template_tree.configure(yscrollcommand=scrollbar.set)
        
        self.template_tree.grid(row=0, column=0, columnspan=3, sticky="nsew", pady=5)
        scrollbar.grid(row=0, column=3, sticky="ns")
        
        # 添加操作按钮
        btn_frame = tk.Frame(template_frame, bg="#f0f0f0")
        btn_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=5)
        
        # 添加语言输入框
        self.new_lang_var = tk.StringVar()
        lang_entry = tk.Entry(btn_frame, textvariable=self.new_lang_var, width=10)
        lang_entry.pack(side=tk.LEFT, padx=5)
        
        add_btn = tk.Button(btn_frame, text="添加语言", command=self.add_template_language)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        remove_btn = tk.Button(btn_frame, text="删除语言", command=self.remove_template_language)
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        browse_btn = tk.Button(btn_frame, text="浏览路径", command=self.browse_template_path)
        browse_btn.pack(side=tk.LEFT, padx=5)
    
    def toggle_staggered(self):
        """切换交错发送状态"""
        if self.staggered_var.get():
            print("[INFO] 已关闭交错发送模式")
        else:
            print("[INFO] 已启用交错发送模式")
    
    def toggle_single_sender(self):
        """切换单邮箱发送状态"""
        if self.single_sender_enabled.get():
            self.sender_combobox.config(state='normal')
            print("[INFO] 已启用单邮箱发送模式")
        else:
            self.sender_combobox.config(state='disabled')
            print("[INFO] 已关闭单邮箱发送模式")
    
    def toggle_sign_mode(self, event=None):
        """切换署名模式"""
        if self.sign_mode_var.get() == '固定署名':
            self.fixed_sign_entry.config(state='normal')
            print("[INFO] 已选择固定署名模式")
        else:
            self.fixed_sign_entry.config(state='disabled')
            print("[INFO] 已选择使用账户配置署名模式")
    
    def load_config(self):
        """加载配置文件到UI"""
        try:
            # 确保配置文件存在
            findConfig()
            self.config = readConfig()
            
            # 基本设置
            settings = self.config['settings']
            self.interval_var.set(settings.get('intervalSendingTime', 180))
            
            # 加载账户列表
            accounts = list(self.config['accounts'].keys())
            accounts = [acc for acc in accounts if not acc.startswith('//') and acc != 'exampleAccount']
            self.sender_combobox['values'] = accounts
            if accounts:
                self.sender_combobox.set(accounts[0])
                
            # 修复交错发送设置
            staggered_value = settings.get('staggeredSending', 0)
            self.staggered_var.set(bool(staggered_value))
            self.staggered_cb.config(text=f"关闭交错发送(多邮箱时有效)")
            
            # 单邮箱设置
            single_sender = settings.get('singleSender', {})
            self.single_sender_enabled.set(bool(single_sender.get('enabled', 0)))
            #print(single_sender.get('defaultSenderName', ''))
            self.single_sender_var.set(single_sender.get('defaultSenderName', ''))
            
            # 筛选设置
            charger_name = settings.get('chargerName', {})
            self.filter_mode_var.set('匹配负责人' if charger_name.get('statue', 0) == 0 else '排除负责人')
            self.charger_name_var.set(charger_name.get('chargerName', 'bobo'))
            
            # 署名设置 - 移除不支持的选项
            sender_name = settings.get('senderName', {})
            sign_mode_index = sender_name.get('statue', 1)
            if sign_mode_index == 2:  # 如果是不支持的选项，降级为使用账户配置
                sign_mode_index = 1
                print("[WARNING] '按照表格署名'模式暂不支持，已自动切换为'使用账户配置'")
            self.sign_mode_var.set(['固定署名', '使用账户配置'][sign_mode_index])
            self.fixed_sign_var.set(sender_name.get('senderName', 'abc'))
            
            # 加载模板设置 - 忽略注释
            template_settings = settings['mailModelContent']
            for key, path in template_settings.items():
                if not key.startswith('//'):  # 忽略注释
                    self.template_tree.insert("", "end", values=(key, path))
            
            # 更新UI状态
            self.toggle_single_sender()
            self.toggle_sign_mode()
            
            print("[INFO] 配置加载成功")
        except Exception as e:
            print(f"[ERROR] 加载配置失败: {str(e)}")
    
    def save_config(self):
        """保存UI设置到配置文件"""
        try:
            # 确保配置文件存在
            findConfig()
            
            with open('config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 更新基本设置
            settings = config['settings']
            settings['intervalSendingTime'] = int(self.interval_var.get())
            
            # 修复交错发送设置
            settings['staggeredSending'] = 1 if self.staggered_var.get() else 0
            
            # 更新单邮箱设置
            settings['singleSender']['enabled'] = 1 if self.single_sender_enabled.get() else 0
            settings['singleSender']['defaultSenderName'] = self.single_sender_var.get()
            
            # 更新筛选设置
            charger_name = settings['chargerName']
            charger_name['statue'] = 0 if self.filter_mode_var.get() == '匹配负责人' else 1
            charger_name['chargerName'] = self.charger_name_var.get()
            
            # 更新署名设置
            sender_name = settings['senderName']
            sender_name['statue'] = 0 if self.sign_mode_var.get() == '固定署名' else 1
            sender_name['senderName'] = self.fixed_sign_var.get()
            
            # 更新模板设置
            template_settings = {}
            for child in self.template_tree.get_children():
                lang, path = self.template_tree.item(child)['values']
                template_settings[lang] = path
            settings['mailModelContent'] = template_settings
            
            # 保存配置文件
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            
            print("[INFO] 配置保存成功")
            return True
        except Exception as e:
            print(f"[ERROR] 保存配置失败: {str(e)}")
            return False
    
    def add_template_language(self):
        """添加新的语言模板"""
        lang = self.new_lang_var.get().strip().upper()
        if not lang:
            messagebox.showwarning("警告", "请输入语言代码")
            return
            
        # 检查是否已存在
        for child in self.template_tree.get_children():
            existing_lang = self.template_tree.item(child)['values'][0]
            if existing_lang == lang:
                messagebox.showwarning("警告", f"语言 {lang} 已存在")
                return
        
        path = filedialog.askopenfilename(title=f"选择 {lang} 语言模板文件", 
                                        filetypes=[("文本文件", "*.txt")])
        if path:
            self.template_tree.insert("", "end", values=(lang, path))
            self.new_lang_var.set("")  # 清空输入框
            print(f"[INFO] 已添加 {lang} 语言模板")
    
    def remove_template_language(self):
        """删除选中的语言模板"""
        selected = self.template_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的语言")
            return
            
        lang = self.template_tree.item(selected)['values'][0]
        if messagebox.askyesno("确认删除", f"确定要删除 {lang} 语言模板吗?"):
            self.template_tree.delete(selected)
            print(f"[INFO] 已删除 {lang} 语言模板")
    
    def browse_template_path(self):
        """浏览并更新选中的模板路径"""
        selected = self.template_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要修改的语言")
            return
            
        lang, old_path = self.template_tree.item(selected)['values']
        path = filedialog.askopenfilename(title=f"选择 {lang} 语言模板文件", 
                                        filetypes=[("文本文件", "*.txt")],
                                        initialfile=os.path.basename(old_path),
                                        initialdir=os.path.dirname(old_path))
        if path:
            self.template_tree.item(selected, values=(lang, path))
            print(f"[INFO] 已更新 {lang} 语言模板路径")
    
    def preview_template(self):
        """预览邮件模板"""
        try:
            # 创建模板选择窗口
            preview_win = tk.Toplevel(self.root)
            preview_win.title("邮件模板预览")
            preview_win.geometry("900x700")
            
            # 语言选择
            lang_frame = tk.Frame(preview_win)
            lang_frame.pack(fill=tk.X, padx=10, pady=10)
            
            tk.Label(lang_frame, text="选择语言模板:").pack(side=tk.LEFT)
            
            lang_var = tk.StringVar()
            lang_combobox = ttk.Combobox(lang_frame, textvariable=lang_var, width=10)
            
            # 获取可用语言
            languages = []
            for child in self.template_tree.get_children():
                lang = self.template_tree.item(child)['values'][0]
                languages.append(lang)
            
            if not languages:
                messagebox.showwarning("警告", "没有配置任何语言模板")
                return
                
            lang_combobox['values'] = languages
            lang_combobox.set(languages[0])
            lang_combobox.pack(side=tk.LEFT, padx=10)
            
            # 预览数据输入
            data_frame = tk.Frame(preview_win)
            data_frame.pack(fill=tk.X, padx=10, pady=5)
            
            tk.Label(data_frame, text="发送人姓名:").grid(row=0, column=0, sticky="w", padx=5)
            self.sender_name_var = tk.StringVar(value="Coco")
            tk.Entry(data_frame, textvariable=self.sender_name_var, width=15).grid(row=0, column=1, sticky="w", padx=5)
            
            tk.Label(data_frame, text="接收人姓名:").grid(row=0, column=2, sticky="w", padx=5)
            self.receiver_name_var = tk.StringVar(value="达人姓名")
            tk.Entry(data_frame, textvariable=self.receiver_name_var, width=15).grid(row=0, column=3, sticky="w", padx=5)
            
            # 模板内容显示
            content_frame = tk.Frame(preview_win)
            content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            scrollbar = tk.Scrollbar(content_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            self.template_text = tk.Text(
                content_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set,
                bg="white", font=("Consolas", 10)
            )
            self.template_text.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.template_text.yview)
            
            # 操作按钮
            btn_frame = tk.Frame(preview_win)
            btn_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # 替换按钮
            replace_btn = tk.Button(
                btn_frame, text="替换标记", 
                command=lambda: self.replace_template_tags(
                    self.sender_name_var.get(),
                    self.receiver_name_var.get()
                )
            )
            replace_btn.pack(side=tk.LEFT, padx=5)
            
            # 保存按钮
            save_btn = tk.Button(
                btn_frame, text="保存修改", 
                command=lambda: self.save_template_changes(lang_var.get())
            )
            save_btn.pack(side=tk.LEFT, padx=5)
            
            # 标记说明
            note_frame = tk.Frame(preview_win)
            note_frame.pack(fill=tk.X, padx=10, pady=5)
            
            note = "模板标记说明:\n"
            note += "  [senderName] - 将被替换为发送人姓名\n"
            note += "  [recName] - 将被替换为接收人姓名\n"
            note += "  <替换标记> 在模板中替换上面两者\n"
            note += "  <保存修改> 保存上面显示的模板"
            tk.Label(note_frame, text=note, justify=tk.LEFT, anchor="w").pack(fill=tk.X)
            
            # 加载模板函数
            def load_template(lang):
                # 查找模板路径
                template_path = None
                for child in self.template_tree.get_children():
                    l, path = self.template_tree.item(child)['values']
                    if l == lang:
                        template_path = path
                        break
                
                if not template_path:
                    self.template_text.delete(1.0, tk.END)
                    self.template_text.insert(tk.END, f"未找到{lang}语言的模板路径")
                    return
                
                try:
                    with open(template_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    self.template_text.delete(1.0, tk.END)
                    self.template_text.insert(tk.END, content)
                    
                    # 高亮标记位置
                    self.highlight_tags()
                    
                except Exception as e:
                    self.template_text.delete(1.0, tk.END)
                    self.template_text.insert(tk.END, f"加载模板失败: {str(e)}")
            
            def update_template_preview(*args):
                load_template(lang_var.get())
            
            # 绑定数据变化
            self.sender_name_var.trace_add('write', lambda *args: load_template(lang_var.get()))
            self.receiver_name_var.trace_add('write', lambda *args: load_template(lang_var.get()))
            
            # 初始加载
            load_template(languages[0])
            
            # 绑定语言选择事件
            lang_var.trace_add('write', lambda *args: load_template(lang_var.get()))
            
        except Exception as e:
            messagebox.showerror("错误", f"预览模板失败: {str(e)}")
    
    def replace_template_tags(self, sender, receiver):
        """替换模板中的标记 - 需要您实现具体替换逻辑"""
        # 这里只是示例，您需要根据实际需求实现替换逻辑
        content = self.template_text.get("1.0", tk.END)
        
        # 示例替换逻辑 - 请根据您的需求修改
        content = content.replace("[senderName]", sender)
        content = content.replace("[recName]", receiver)
        #content = content.replace("[日期]", time.strftime("%Y-%m-%d"))
        
        self.template_text.delete("1.0", tk.END)
        self.template_text.insert("1.0", content)
        
        # 重新高亮标记
        self.highlight_tags()
        
        print("[INFO] 已执行标记替换")
    
    def save_template_changes(self, lang):
        """保存模板修改 - 需要您实现具体保存逻辑"""
        # 查找模板路径
        template_path = None
        for child in self.template_tree.get_children():
            l, path = self.template_tree.item(child)['values']
            if l == lang:
                template_path = path
                break
        
        if not template_path:
            messagebox.showerror("错误", f"找不到{lang}语言的模板路径")
            return
        
        try:
            content = self.template_text.get("1.0", tk.END)
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"[INFO] {lang} 语言模板已保存")
            messagebox.showinfo("成功", "模板修改已保存")
        except Exception as e:
            print(f"[ERROR] 保存模板失败: {str(e)}")
            messagebox.showerror("错误", f"保存模板失败: {str(e)}")
    
    def highlight_tags(self):
        """高亮模板中的标记"""
        # 清除之前的高亮
        self.template_text.tag_remove("tag_highlight", "1.0", tk.END)
        
        # 定义要匹配的标记模式
        patterns = [
            r"\[发送人\]", r"\[接收人\]", r"\[日期\]", 
            r"\[.*?\]"  # 匹配所有其他标记,通过这里高亮[recName],[senderName].模板替换函数replace_template_tags
        ]
        
        # 为每个模式添加高亮
        for pattern in patterns:
            start = "1.0"
            while True:
                start = self.template_text.search(pattern, start, stopindex=tk.END, regexp=True)
                if not start:
                    break
                end = f"{start}+{len(self.template_text.get(start, f'{start} lineend'))}c"
                self.template_text.tag_add("tag_highlight", start, end)
                start = end
        
        # 配置高亮样式
        self.template_text.tag_config("tag_highlight", background="#fffacd", foreground="#d2691e")
    
    def start_sending(self):
        """开始发送邮件"""
        if not self.save_config():
            messagebox.showerror("错误", "无法保存配置，请检查配置是否正确")
            return
        
        # 禁用开始按钮，启用停止按钮
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 创建停止事件
        self.stop_event.clear()
        
        # 在新线程中运行发送过程
        self.sending_thread = threading.Thread(target=self.run_sending_process, daemon=True)
        self.sending_thread.start()
    
    def run_sending_process(self):
        """运行发送过程"""
        try:
            print("[INFO] 开始发送邮件...")
            
            # 加载配置
            findConfig()
            config = readConfig()
            
            # 删除不需要的配置项
            #del config['accounts']['//']
            #del config['accounts']['exampleAccount']
            #del config['settings']['mailModelContent']['//']
            
            # 读取表格
            sheet,workbook = readSheet(config['settings']['content'])
            
            # 获取收件人
            reciverdict = getReciver(sheet, config['settings']["xlsxKeys"], config['settings']["chargerName"])
            
            if not reciverdict:
                print("[WARNING] 没有找到符合条件的收件人, 请更改筛选条件再试")
                self.reset_buttons()
                return
            
            reciver = [reciverdict[name]['mail'] for name in reciverdict]
            print(f"邮件将发送到{len(reciver)}个收件人")
            print(f'收件人名单: \n{reciver}')
            
            # 准备账户配置
            accounts_config = dict(config['accounts'])

            if 'exampleAccount' in accounts_config:
                del accounts_config['exampleAccount']
            if '//' in accounts_config:
                del accounts_config['//']

            if not accounts_config:
                print("[ERROR] 没有配置任何邮箱账号!")
                self.reset_buttons()
                return
            
            # 设置发送参数
            settings = config['settings']
            
            # 根据配置选择发送方式
            if len(accounts_config) == 1 or settings['singleSender']['enabled'] or len(reciverdict) == 1:
                # 单邮箱发送
                if settings['singleSender']['defaultSenderName'] in accounts_config:
                    account_name = settings['singleSender']['defaultSenderName']
                else:
                    print(f'[WARNING] 默认发送邮箱配置有误!')
                    print(f'[WARNING] 可供选择的账户: {list(accounts_config.keys())}')
                    account_name = list(accounts_config.keys())[0]
                    print(f'[INFO] 已自动选择{account_name}')
                account_config = accounts_config[account_name]
                print(f"\n使用单邮箱发送: {account_name}")
                print(f'''预计用时{round(((len(reciverdict)-1)*settings["intervalSendingTime"])/60,1)}分钟''')
                sendMailCore.send_emails(
                    account_name, 
                    account_config, 
                    reciverdict, 
                    settings,
                    stop_event=self.stop_event
                )
            else:
                # 多邮箱发送
                print(f"\n使用多邮箱发送 ({len(accounts_config)}个邮箱)")
                recDictList = split_dict_avg(reciverdict, len(accounts_config))
                #print(f'{settings["staggeredSending"]} statue')
                if settings["staggeredSending"]: #关闭交错发送
                    if len(reciverdict)<=2:
                        if len(reciverdict) == 1:
                            print(f'''预计用时0分钟''')
                        else:
                            print(f'''预计用时round(settings["intervalSendingTime"]/60,1)分钟''')
                    else:
                        print(f'''预计用时{round((len(reciverdict)-2)*settings["intervalSendingTime"]/60,1)}分钟''')
                else:  #开启交错发送
                    print(f'''预计用时{round(((len(reciverdict)-1)/2)*settings["intervalSendingTime"]/60,1)}分钟''')
                
                threads = []
                for i, account_name in enumerate(accounts_config.keys()):
                    if self.stop_event.is_set():
                        break
                    
                    account_config = accounts_config[account_name]
                    print(f"[INFO] 启动发送线程: {account_name}")
                    
                    thread = threading.Thread(
                        target=sendMailCore.send_emails,
                        args=(account_name, account_config, recDictList[i]),
                        kwargs={
                            'settings': settings,
                            'threadID': account_name,
                            'stop_event': self.stop_event,
                            'workbook': workbook,  # 传入工作表对象
                            'key_column': config['settings']["xlsxKeys"]["sendingStatus"],
                            'sheet_lock': self.sheet_lock  # 传入锁对象
                        }
                    )
                    thread.start()
                    threads.append(thread)
                    time.sleep(0.5)

                    if i+2 <= len(recDictList) and (not recDictList[i+1]):
                        continue
                    
                    # 交错发送间隔
                    #print(settings["intervalSendingTime"] // len(accounts_config))
                    ttime = time.time()
                    halfWaitingTime = settings["intervalSendingTime"] // len(accounts_config)
                    while time.time() < ttime + halfWaitingTime :
                        if self.stop_event.is_set():
                            #print(f'[WARNING] 用户终止了线程:{threadID}')
                            break  
                        time.sleep(1)
                    
                    if settings["staggeredSending"]:
                        thread.join()
                
                # 等待所有线程完成
                for thread in threads:
                    thread.join()
            
            if self.stop_event.is_set():
                print("[INFO] 邮件发送失败: 用户终止了发送")
            else:
                print("[INFO] 邮件发送完成!")
        
        except Exception as e:
            print(f"[ERROR] 发送过程中出错: {str(e)}")
        finally:
            self.reset_buttons()
    
    def stop_sending(self):
        """终止发送过程"""
        self.stop_event.set()
        print("[DEBUG] 用户终止了线程")
        print("[INFO] 正在终止发送过程, 请等待线程响应...")
        self.stop_btn.config(state=tk.DISABLED)
    
    def reset_buttons(self):
        """重置按钮状态"""
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
    
    def on_exit(self):
        """退出程序时自动保存配置并关闭日志文件"""
        if messagebox.askyesno("退出", "确定要退出吗? 配置将自动保存"):
            self.save_config()
            if hasattr(self, 'log_writer'):
                self.log_writer.close()  # 安全关闭日志文件
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderUI(root)
    root.mainloop()

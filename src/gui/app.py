# -*- coding: utf-8 -*-
"""
主界面程序
==========
使用 tkinter 构建的交互式对话框，用于：
1. 输入数据库连接信息（自动保留上次输入）
2. 连接后展示账套列表（多选勾选框）
3. 执行生成合并汇总报表

设计原则：
- 数据库相关模块使用延迟导入，确保GUI界面可独立启动
- 连接、查询、生成均在后台线程执行，不阻塞UI
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import sys

# 确保项目根目录在 sys.path 中
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from src.gui.config_manager import ConnectionConfigManager


class AccountCheckboxFrame(ttk.LabelFrame):
    """账套勾选列表组件"""

    def __init__(self, parent, accounts: list, **kwargs):
        """
        :param parent: 父容器
        :param accounts: AccountInfo 列表
        """
        super().__init__(parent, text="请选择需生成报表的账套（默认全选）", padding=5, **kwargs)
        self.accounts = accounts
        self._checkboxes: list[ttk.Checkbutton] = []
        self._vars: list[tk.BooleanVar] = []
        self._build_ui()

    def _build_ui(self):
        """构建勾选列表界面"""
        # 使用 Canvas + Scrollbar 实现滚动
        canvas = tk.Canvas(self, highlightthickness=0, height=200)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # 全选/取消全选按钮
        btn_frame = ttk.Frame(scroll_frame)
        btn_frame.pack(fill="x", pady=(0, 5))

        def select_all():
            for v in self._vars:
                v.set(True)

        def deselect_all():
            for v in self._vars:
                v.set(False)

        ttk.Button(btn_frame, text="全选", width=6, command=select_all).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="取消全选", width=8, command=deselect_all).pack(side="left", padx=2)
        ttk.Label(btn_frame, text=f"共 {len(self.accounts)} 个账套").pack(side="right", padx=5)

        # 表头
        header = ttk.Frame(scroll_frame)
        header.pack(fill="x", pady=2)
        ttk.Label(header, text="  选择", width=6, anchor="w", font=("微软雅黑", 9, "bold")).pack(side="left")
        ttk.Label(header, text="账套号", width=10, anchor="w", font=("微软雅黑", 9, "bold")).pack(side="left")
        ttk.Label(header, text="账套名称", width=30, anchor="w", font=("微软雅黑", 9, "bold")).pack(side="left", fill="x")

        ttk.Separator(scroll_frame, orient="horizontal").pack(fill="x", pady=2)

        # 逐行添加账套勾选框
        for acc in self.accounts:
            var = tk.BooleanVar(value=True)  # 默认全选
            row_frame = ttk.Frame(scroll_frame)
            row_frame.pack(fill="x", pady=1)

            cb = ttk.Checkbutton(row_frame, variable=var, text="")
            cb.pack(side="left", padx=(5, 2))

            ttk.Label(row_frame, text=acc.cAcc_Id, width=10, anchor="w").pack(side="left")
            ttk.Label(row_frame, text=acc.cAcc_Name, anchor="w").pack(side="left", fill="x", padx=2)

            self._checkboxes.append(cb)
            self._vars.append(var)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def get_selected_accounts(self) -> list:
        """
        获取用户勾选的所有账套
        :return: 选中的 AccountInfo 列表
        """
        return [
            acc for acc, var in zip(self.accounts, self._vars)
            if var.get()
        ]


class MainWindow:
    """主界面窗口"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("分店财务报表汇总工具 v2.0")
        self.root.geometry("600x520")
        self.root.minsize(560, 460)

        # 尝试设置图标（可选）
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

        # 居中显示
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"+{x}+{y}")

        # 全局状态（DatabaseConnector 在后台线程延迟导入，此处用 object）
        self._db_connector = None
        self._all_accounts: list = []
        self._is_connected: bool = False

        # 构建UI
        self._build_connection_panel()
        self._build_account_panel()
        self._build_action_panel()
        self._build_status_bar()

        # 加载上次保存的配置
        self._load_saved_config()

        # 协议处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ==================== UI 构建 ====================

    def _build_connection_panel(self):
        """构建数据库连接信息输入面板"""
        conn_frame = ttk.LabelFrame(self.root, text="数据库连接设置", padding=8)
        conn_frame.pack(fill="x", padx=8, pady=(8, 3))

        # 网格布局
        conn_frame.columnconfigure(1, weight=1)

        # 第0行：服务器
        ttk.Label(conn_frame, text="服务器地址：", width=14, anchor="e").grid(
            row=0, column=0, sticky="e", padx=(0, 3), pady=3
        )
        self.entry_server = ttk.Entry(conn_frame)
        self.entry_server.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=3)
        self.entry_server.bind("<Return>", lambda e: self.entry_username.focus_set())
        self.entry_server.bind("<KP_Enter>", lambda e: self.entry_username.focus_set())

        # 第1行：用户名
        ttk.Label(conn_frame, text="登录用户名：", width=14, anchor="e").grid(
            row=1, column=0, sticky="e", padx=(0, 3), pady=3
        )
        self.entry_username = ttk.Entry(conn_frame)
        self.entry_username.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=3)
        self.entry_username.bind("<Return>", lambda e: self.entry_password.focus_set())
        self.entry_username.bind("<KP_Enter>", lambda e: self.entry_password.focus_set())

        # 第2行：密码
        ttk.Label(conn_frame, text="登录密码：", width=14, anchor="e").grid(
            row=2, column=0, sticky="e", padx=(0, 3), pady=3
        )
        self.entry_password = ttk.Entry(conn_frame, show="*")
        self.entry_password.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=3)
        self.entry_password.bind("<Return>", lambda e: self._connect_database())
        self.entry_password.bind("<KP_Enter>", lambda e: self._connect_database())

        # 第3行：连接按钮
        btn_frame = ttk.Frame(conn_frame)
        btn_frame.grid(row=3, column=1, sticky="w", pady=(4, 0))

        self.btn_connect = ttk.Button(
            btn_frame, text="连接数据库", width=14,
            command=self._connect_database
        )
        self.btn_connect.pack(side="left", padx=(0, 8))

        self.label_conn_status = ttk.Label(btn_frame, text="", foreground="gray")
        self.label_conn_status.pack(side="left")

    def _build_account_panel(self):
        """构建账套选择面板"""
        self.account_frame = ttk.LabelFrame(self.root, text="账套列表（等待连接...）", padding=3)
        self.account_frame.pack(fill="both", expand=True, padx=8, pady=3)

        # 占位提示
        self.placeholder_label = ttk.Label(
            self.account_frame,
            text="请先连接数据库后，自动加载账套列表",
            foreground="gray",
            anchor="center"
        )
        self.placeholder_label.pack(expand=True, fill="both")

        # 账套勾选框容器（初始隐藏）
        self.account_inner_frame = ttk.Frame(self.account_frame)
        self.account_checkbox_frame: AccountCheckboxFrame | None = None

    def _build_action_panel(self):
        """构建操作按钮面板"""
        action_frame = ttk.Frame(self.root, padding=6)
        action_frame.pack(fill="x", padx=8, pady=(0, 6))

        self.btn_generate = ttk.Button(
            action_frame, text="生成合并汇总报表", width=22,
            command=self._generate_report,
            state="disabled"
        )
        self.btn_generate.pack(side="left", padx=(0, 10))

        self.btn_cancel = ttk.Button(
            action_frame, text="取消", width=8,
            command=self._on_close
        )
        self.btn_cancel.pack(side="left")

        # 进度条（初始隐藏）
        self.progress = ttk.Progressbar(
            action_frame, mode="indeterminate", length=150
        )

        # 生成按钮的快捷键：先绑定到窗口
        self.root.bind("<Return>", self._on_enter_global)
        self.root.bind("<KP_Enter>", self._on_enter_global)

    def _build_status_bar(self):
        """构建底部状态栏"""
        self.status_bar = ttk.Label(
            self.root, text="就绪", relief="sunken",
            anchor="w", padding=(3, 1)
        )
        self.status_bar.pack(side="bottom", fill="x")

    # ==================== 逻辑方法 ====================

    def _load_saved_config(self):
        """加载上次保存的连接配置"""
        config = ConnectionConfigManager.load()
        self.entry_server.insert(0, config.get("server", ""))
        self.entry_username.delete(0, tk.END)
        self.entry_username.insert(0, config.get("username", "sa"))
        self.entry_password.delete(0, tk.END)
        self.entry_password.insert(0, config.get("password", ""))

    def _connect_database(self):
        """连接数据库（在后台线程执行）"""
        server = self.entry_server.get().strip()
        username = self.entry_username.get().strip()
        password = self.entry_password.get().strip()

        if not server:
            messagebox.showwarning("提示", "请输入服务器地址或IP")
            self.entry_server.focus_set()
            return
        if not username:
            messagebox.showwarning("提示", "请输入数据库登录用户名")
            self.entry_username.focus_set()
            return

        # 禁用界面
        self._set_connection_ui_enabled(False)
        self._set_status("正在连接数据库...")

        # 保存配置
        ConnectionConfigManager.save(server, username, password)

        # 启动线程
        thread = threading.Thread(
            target=self._connect_worker,
            args=(server, username, password),
            daemon=True
        )
        thread.start()

    def _connect_worker(self, server: str, username: str, password: str):
        """
        后台线程：连接数据库并查询账套
        """
        try:
            from src.database.connector import DatabaseConnector
            from src.database.ufsystem import UFSystemQuerier

            config = {
                "server": server,
                "port": 1433,
                "username": username,
                "password": password,
                "timeout": 30,
                "charset": "GBK",
            }

            connector = DatabaseConnector(config)
            # 尝试连接 UFSystem 公共数据库
            connector.connect("UFSystem")
            self._db_connector = connector

            # 查询所有账套列表
            querier = UFSystemQuerier(connector)
            all_accounts = querier.get_all_accounts()

            # 过滤 998, 999 演示账套
            self._all_accounts = [
                acc for acc in all_accounts
                if acc.cAcc_Id not in ("998", "999")
            ]

            # 主线程更新UI
            self.root.after(0, self._on_connect_success)

        except Exception as e:
            self.root.after(0, self._on_connect_failed, str(e))

    def _on_connect_success(self):
        """连接成功，更新UI"""
        self._is_connected = True
        self._set_connection_ui_enabled(True)
        self._set_status(f"已连接 | 共发现 {len(self._all_accounts)} 个有效账套")

        # 更新账套列表区域
        self.placeholder_label.pack_forget()
        for w in self.account_inner_frame.winfo_children():
            w.destroy()
        self.account_inner_frame.pack(fill="both", expand=True)

        self.account_checkbox_frame = AccountCheckboxFrame(
            self.account_inner_frame, self._all_accounts
        )
        self.account_checkbox_frame.pack(fill="both", expand=True)

        # 更新标签文字
        self.account_frame.configure(text="账套列表（请勾选需生成报表的账套）")

        # 启用生成按钮
        self.btn_generate.config(state="normal")
        self.btn_generate.focus_set()

        # 更新状态栏
        self.label_conn_status.config(text="✅ 已连接", foreground="green")
        self.btn_connect.config(text="重新连接")

    def _on_connect_failed(self, error_msg: str):
        """连接失败"""
        self._is_connected = False
        self._set_connection_ui_enabled(True)
        self._set_status("连接失败")

        self.label_conn_status.config(text="❌ 连接失败", foreground="red")
        self.btn_connect.config(text="🔗  连接数据库")

        messagebox.showerror("数据库连接失败", f"无法连接到数据库：\n{error_msg}")

    def _on_enter_global(self, event):
        """
        全局回车键处理：
        - 如果账套列表已加载且有焦点在列表上，选中确定后跳到生成按钮
        - 如果是生成按钮获得焦点，执行生成
        """
        focused = self.root.focus_get()

        # 如果焦点在账套列表区域且已连接，转到生成按钮
        if self._is_connected and self.btn_generate.cget("state") == "normal":
            # 检查焦点是否在勾选框列表内
            if isinstance(focused, ttk.Checkbutton) or isinstance(focused, ttk.Button) or isinstance(focused, ttk.Entry):
                pass  # 让默认行为处理
            self.btn_generate.focus_set()
            return "break"

    def _generate_report(self):
        """生成合并汇总报表"""
        if not self._is_connected or not self.account_checkbox_frame:
            messagebox.showwarning("提示", "请先连接数据库")
            return

        selected = self.account_checkbox_frame.get_selected_accounts()
        if not selected:
            messagebox.showwarning("提示", "请至少勾选一个账套")
            return

        # 禁用操作按钮，显示进度
        self.btn_generate.config(state="disabled")
        self.btn_cancel.config(state="disabled")
        self.progress.pack(side="left", padx=(20, 0))
        self.progress.start()
        self._set_status(f"正在生成报表，共 {len(selected)} 个账套...")

        # 后台执行
        thread = threading.Thread(
            target=self._generate_worker,
            args=(selected,),
            daemon=True
        )
        thread.start()

    def _generate_worker(self, selected_accounts: list):
        """
        后台线程：生成报表
        """
        try:
            # 动态构建配置
            db_config = {
                "server": self.entry_server.get().strip(),
                "port": 1433,
                "username": self.entry_username.get().strip(),
                "password": self.entry_password.get().strip(),
                "timeout": 30,
                "charset": "GBK",
            }

            config = {
                "database": db_config,
                "account_filter": {
                    "include_ids": [acc.cAcc_Id for acc in selected_accounts],
                    "exclude_ids": ["998", "999"],
                },
                "template": {
                    "filepath": os.path.join(_PROJECT_ROOT, "分店财务报表模板.xlsx"),
                    "source_sheet": "三江店报表",
                    "year_row": 1,
                    "header_row": 2,
                    "data_start_row": 3,
                },
                "output": {
                    "dir": os.path.join(_PROJECT_ROOT, "output"),
                    "filename_prefix": "分店财务报表_",
                    "file_extension": ".xlsx",
                    "open_when_done": True,
                },
                "report_year": 2024,
                "report_months": [],
                "logging": {
                    "level": "INFO",
                    "file": os.path.join(_PROJECT_ROOT, "logs", "app.log"),
                    "console": False,
                },
            }

            from src.report.builder import ReportBuilder
            builder = ReportBuilder(config)
            output_path = builder.build()

            self.root.after(0, self._on_generate_success, output_path)

        except Exception as e:
            self.root.after(0, self._on_generate_failed, str(e))

    def _on_generate_success(self, output_path: str):
        """生成成功"""
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_generate.config(state="normal")
        self.btn_cancel.config(state="normal")
        self._set_status(f"报表生成完成: {output_path}")

        result = messagebox.askyesno(
            "生成成功",
            f"✅ 合并汇总报表生成成功！\n\n文件路径：{output_path}\n\n是否立即打开文件？"
        )
        if result:
            try:
                os.startfile(output_path)
            except Exception:
                pass

    def _on_generate_failed(self, error_msg: str):
        """生成失败"""
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_generate.config(state="normal")
        self.btn_cancel.config(state="normal")
        self._set_status("报表生成失败")
        messagebox.showerror("生成失败", f"报表生成过程中发生错误：\n{error_msg}")

    def _set_connection_ui_enabled(self, enabled: bool):
        """启用/禁用连接面板控件"""
        state = "normal" if enabled else "disabled"
        self.entry_server.config(state=state)
        self.entry_username.config(state=state)
        self.entry_password.config(state=state)
        self.btn_connect.config(state=state)

    def _set_status(self, text: str):
        """设置状态栏文字"""
        self.status_bar.config(text=text)
        self.root.update_idletasks()

    def _on_close(self):
        """关闭窗口前保存配置"""
        server = self.entry_server.get().strip()
        username = self.entry_username.get().strip()
        password = self.entry_password.get().strip()
        if server:
            ConnectionConfigManager.save(server, username, password)
        self.root.destroy()

    def run(self):
        """启动主循环"""
        self.root.mainloop()
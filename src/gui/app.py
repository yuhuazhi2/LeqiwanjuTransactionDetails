# -*- coding: utf-8 -*-
"""
主界面程序
==========
使用 tkinter 构建的交互式对话框，用于：
1. 输入数据库连接信息（自动保留上次输入）
2. 连接后展示账套列表（每行带年份下拉选择框）
3. 执行生成报表框架雏形

设计原则：
- 数据库相关模块使用延迟导入，确保GUI界面可独立启动
- 连接、查询、生成均在后台线程执行，不阻塞UI
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import sys
from datetime import datetime

# 确保项目根目录在 sys.path 中
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from src.gui.config_manager import ConnectionConfigManager


class AccountYearFrame(ttk.LabelFrame):
    """
    账套列表组件（每行用年份下拉框替代勾选框）
    =========================================
    每个账套一行，包含：
      - 年份下拉框（从 UA_Period 读取）
      - 账套号
      - 账套名称
    """

    def __init__(self, parent, accounts: list, years_map: dict[str, list[int]],
                 default_year: int, **kwargs):
        """
        :param parent: 父容器
        :param accounts: AccountInfo 列表
        :param years_map: {账套号: [年份列表(降序)]}
        :param default_year: 默认选中年份
        """
        super().__init__(parent, text="请选择每个账套的报表年份", padding=5, **kwargs)
        self.accounts = accounts
        self.years_map = years_map
        self.default_year = default_year
        self._combos: list[ttk.Combobox] = []
        self._year_vars: list[tk.StringVar] = []
        self._build_ui()

    def _build_ui(self):
        """构建账套列表界面"""
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

        # 表头
        header = ttk.Frame(scroll_frame)
        header.pack(fill="x", pady=2)
        ttk.Label(header, text="  年份", width=10, anchor="w",
                  font=("微软雅黑", 9, "bold")).pack(side="left")
        ttk.Label(header, text="账套号", width=10, anchor="w",
                  font=("微软雅黑", 9, "bold")).pack(side="left")
        ttk.Label(header, text="账套名称", width=30, anchor="w",
                  font=("微软雅黑", 9, "bold")).pack(side="left", fill="x")

        ttk.Separator(scroll_frame, orient="horizontal").pack(fill="x", pady=2)

        # 逐行添加账套（年份下拉框 + 账套号 + 账套名称）
        for acc in self.accounts:
            row_frame = ttk.Frame(scroll_frame)
            row_frame.pack(fill="x", pady=1)

            # 年份下拉框
            year_var = tk.StringVar()
            # 获取该账套对应的年份列表
            years = self.years_map.get(acc.cAcc_Id, [])
            year_strings = [str(y) for y in years] if years else [str(self.default_year)]

            combo = ttk.Combobox(
                row_frame, textvariable=year_var,
                values=year_strings, width=8, state="readonly"
            )
            # 确定默认值
            default_str = str(self.default_year)
            if default_str in year_strings:
                combo.set(default_str)
            elif year_strings:
                combo.set(year_strings[0])
            combo.pack(side="left", padx=(5, 2))

            # 账套号
            ttk.Label(row_frame, text=acc.cAcc_Id, width=10,
                      anchor="w").pack(side="left")
            # 账套名称
            ttk.Label(row_frame, text=acc.cAcc_Name,
                      anchor="w").pack(side="left", fill="x", padx=2)

            self._combos.append(combo)
            self._year_vars.append(year_var)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def get_selected_items(self) -> list[tuple]:
        """
        获取每个账套对应的选中年份
        :return: [(AccountInfo, 年份int), ...]
        """
        result = []
        for acc, var in zip(self.accounts, self._year_vars):
            year_str = var.get()
            try:
                year = int(year_str)
            except (ValueError, TypeError):
                continue
            result.append((acc, year))
        return result


class MainWindow:
    """主界面窗口"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("分店财务报表汇总工具 v2.0")
        self.root.geometry("640x520")
        self.root.minsize(600, 460)

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

        # 全局状态
        self._db_connector = None
        self._all_accounts: list = []
        self._years_map: dict[str, list[int]] = {}
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

        # 账套年份选择框容器（初始隐藏）
        self.account_inner_frame = ttk.Frame(self.account_frame)
        self.account_year_frame: AccountYearFrame | None = None

    def _build_action_panel(self):
        """构建操作按钮面板"""
        action_frame = ttk.Frame(self.root, padding=6)
        action_frame.pack(fill="x", padx=8, pady=(0, 6))

        # ---- 输出格式选择（单选按钮组） ----
        # 注意：必须先生成 self.format_frame 再创建 Radiobutton 并"先打包后放置"，
        #       pack(side="left") 保证按钮组水平排列在操作按钮左侧
        format_frame = ttk.LabelFrame(action_frame, text="输出格式", padding=2)
        format_frame.pack(side="left", padx=(0, 10))

        self.format_var = tk.StringVar(value="xlsx")  # 默认 Excel

        rb_excel = ttk.Radiobutton(
            format_frame, text="Excel (.xlsx)", variable=self.format_var,
            value="xlsx"
        )
        rb_excel.pack(side="left", padx=2)

        rb_html = ttk.Radiobutton(
            format_frame, text="HTML (.html)", variable=self.format_var,
            value="html"
        )
        rb_html.pack(side="left", padx=2)

        rb_both = ttk.Radiobutton(
            format_frame, text="同时生成", variable=self.format_var,
            value="both"
        )
        rb_both.pack(side="left", padx=2)

        # ---- 主按钮：完整报表生成（绿色） ----
        self.btn_generate = tk.Button(
            action_frame, text="生成合并汇总报表", width=22,
            command=self._generate_report,
            state="disabled",
            bg="#4CAF50", fg="white", font=("微软雅黑", 9, "bold"),
            relief="raised", bd=2,
            activebackground="#45a049", activeforeground="white",
            disabledforeground="#cccccc"
        )
        self.btn_generate.pack(side="left", padx=(0, 6))

        self.btn_cancel = tk.Button(
            action_frame, text="取消", width=8,
            command=self._on_close,
            bg="#E53935", fg="white", font=("微软雅黑", 9, "bold"),
            relief="raised", bd=2,
            activebackground="#c62828", activeforeground="white"
        )
        self.btn_cancel.pack(side="left")

        # 进度条（初始隐藏）
        self.progress = ttk.Progressbar(
            action_frame, mode="indeterminate", length=150
        )

        # 生成按钮的快捷键：使用 bind_all 确保焦点在 Canvas/Frame 内时也能触发
        self.root.bind_all("<Return>", self._on_enter_global)
        self.root.bind_all("<KP_Enter>", self._on_enter_global)

    def _build_status_bar(self):
        """构建底部状态栏"""
        self.status_bar = ttk.Label(
            self.root, text="就绪", relief="sunken",
            anchor="w", padding=(3, 1)
        )
        self.status_bar.pack(side="bottom", fill="x")

    # ==================== 逻辑方法 ====================

    @staticmethod
    def _get_default_year() -> int:
        """
        根据当前日期确定默认年份：
        - 1月或2月 → 上一年
        - 3月~12月 → 当年
        """
        today = datetime.now()
        month = today.month
        if month <= 2:
            return today.year - 1
        else:
            return today.year

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
        后台线程：连接数据库并查询账套和年份映射
        """
        try:
            from src.database.connector import DatabaseConnector
            from src.database.ufsystem import UFSystemQuerier

            db_config = {
                "server": server,
                "port": 1433,
                "username": username,
                "password": password,
                "timeout": 30,
                "charset": "GBK",
            }

            connector = DatabaseConnector(db_config)
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

            # 批量查询所有账套的年份映射（UA_Period 表）
            # 返回 {账套号: [年份列表(降序)]}
            self._years_map = querier.get_all_account_years_map()

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

        # 计算默认年份
        default_year = self._get_default_year()

        # 创建带年份下拉框的账套列表
        self.account_year_frame = AccountYearFrame(
            self.account_inner_frame,
            self._all_accounts,
            self._years_map,
            default_year
        )
        self.account_year_frame.pack(fill="both", expand=True)

        # 更新标签文字
        self.account_frame.configure(text="账套列表（请为每个账套选择报表年份）")

        # ---- 启用生成按钮，并将焦点移到"生成合并汇总报表"按钮 ----
        self.btn_generate.config(state="normal")
        # 焦点移到"生成合并汇总报表"按钮，这样用户可以立即按回车键执行生成
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
        - 已连接 + 生成按钮可用时：直接触发生成（除非焦点在年份下拉框中）
        - 未连接时：让各输入框自己的 <Return> 处理
        - 生成过程中：不处理
        """
        if not self._is_connected:
            return
        if self.btn_generate.cget("state") != "normal":
            return

        # 如果焦点在 Combobox（年份下拉框）上，不干预——让下拉框正常选择
        focused = self.root.focus_get()
        if isinstance(focused, ttk.Combobox):
            return

        # 直接触发生成，不需要判断焦点位置
        self._generate_report()
        return "break"

    def _generate_report(self):
        """生成合并汇总报表"""
        if not self._is_connected or not self.account_year_frame:
            messagebox.showwarning("提示", "请先连接数据库")
            return

        items = self.account_year_frame.get_selected_items()
        if not items:
            messagebox.showwarning("提示", "账套列表为空，请检查数据库")
            return

        # 禁用操作按钮，显示进度
        self.btn_generate.config(state="disabled")
        self.btn_cancel.config(state="disabled")
        self.progress.pack(side="left", padx=(20, 0))
        self.progress.start()
        self._set_status(f"正在生成报表框架，共 {len(items)} 个账套...")

        # 后台执行
        thread = threading.Thread(
            target=self._generate_worker,
            args=(items,),
            daemon=True
        )
        thread.start()

    def _generate_worker(self, items: list[tuple]):
        """
        后台线程：生成报表框架雏形
        每个账套使用各自选择的年份
        """
        try:
            from src.report.builder import ReportBuilder

            # 按年份分组构建配置，每个年份单独生成一个文件
            # 先按年份分组
            year_groups: dict[int, list] = {}
            for acc, year in items:
                if year not in year_groups:
                    year_groups[year] = []
                year_groups[year].append(acc)

            output_paths = []

            # 为每个年份生成一个文件
            for report_year, accounts in year_groups.items():
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
                        "include_ids": [acc.cAcc_Id for acc in accounts],
                        "exclude_ids": ["998", "999"],
                    },
                    "template": {
                        "filepath": os.path.join(_PROJECT_ROOT, "分店财务报表模板.xlsx"),
                        "source_sheet": "sheet",
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
                    "report_year": report_year,
                    "report_months": [],
                    "logging": {
                        "level": "INFO",
                        "file": os.path.join(_PROJECT_ROOT, "logs", "app.log"),
                        "console": False,
                    },
                }

                # 构建账套->年份映射 {账套号: 年度}
                account_years = {acc.cAcc_Id: year for acc, year in items if year == report_year}

                # 获取用户选择的输出格式（"xlsx" / "html" / "both"）
                output_format = self.format_var.get()

                builder = ReportBuilder(config)
                generated_paths = builder.build_framework(
                    accounts=accounts,
                    account_years=account_years,
                    output_format=output_format  # 传递输出格式参数
                )
                output_paths.extend(generated_paths)

            self.root.after(0, self._on_generate_success, output_paths)

        except Exception as e:
            self.root.after(0, self._on_generate_failed, str(e))

    def _on_generate_success(self, output_paths: list[str]):
        """框架生成成功"""
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_generate.config(state="normal")
        self.btn_cancel.config(state="normal")

        count = len(output_paths)
        path_info = "\n".join(output_paths)
        self._set_status(f"报表框架生成完成，共 {count} 个文件")

        result = messagebox.askyesno(
            "生成成功",
            f"✅ 报表框架生成成功！\n\n"
            f"共生成 {count} 个文件：\n{path_info}\n\n"
            f"是否立即打开最后一个文件预览？"
        )
        if result and output_paths:
            try:
                os.startfile(output_paths[-1])
            except Exception:
                pass

    def _on_generate_failed(self, error_msg: str):
        """框架生成失败"""
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_generate.config(state="normal")
        self.btn_cancel.config(state="normal")
        self._set_status("报表框架生成失败")
        messagebox.showerror("生成失败", f"报表框架生成过程中发生错误：\n{error_msg}")

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
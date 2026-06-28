"""
报表生成器
========
核心类：ReportBuilder
负责协调模板解析、数据提取和Excel生成全流程。
"""

import os
import logging
from datetime import datetime
from typing import Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from src.database.connector import DatabaseConnector
from src.database.ufsystem import UFSystemQuerier, AccountInfo
from src.database.t3data import T3DataExtractor
from src.template.parser import TemplateParser, TemplateLayout

logger = logging.getLogger(__name__)


class ReportBuilder:
    """
    报表生成器
    流程: 加载模板 → 查询账套列表 → 遍历账套取数 → 填充到各页签
    """

    # 默认样式
    HEADER_FONT = Font(name="微软雅黑", bold=True, size=11)
    HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2",
                              fill_type="solid")
    DATA_FONT = Font(name="微软雅黑", size=10)
    TITLE_FONT = Font(name="微软雅黑", bold=True, size=14)
    THIN_BORDER = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    CURRENCY_FORMAT = '#,##0.00'

    # ================================================================
    # 模板行标签 → T3科目映射表
    # key: 模板中的行标签（标准化后），value: 数据提取方法
    # 这是核心映射配置，后续可扩展为外部配置文件
    # ================================================================
    LABEL_TO_DATA_KEY = {
        # 销售业绩（各渠道）
        "油菜花":     "油菜花收入",
        "现金":       "现金收入",
        "美团":       "美团收入",
        "抖音":       "抖音收入",
        "其他业务收入": "其他业务收入",
        # 成本
        "主营业务成本": "__cost__",
        # 营业费用
        "广告费":     "广告费",
        "物料费":     "物料费",
        "设备":       "设备",
        "折旧费":     "折旧费",
        "房租":       "房租",
        "物业费":     "物业费",
        "电费":       "电费",
        "修配费":     "修配费",
        "运杂费":     "运杂费",
        "其他":       "其他",
        # 管理费用
        "工资":       "工资",
        "办公费":     "办公费",
        "差旅费":     "差旅费",
        "业务招待费": "业务招待费",
        "员工福利":   "员工福利",
        "装修费":     "装修费",
        "开办费":     "开办费",
        "服务咨询费": "服务咨询费",
        "社保":       "社保",
        "管理公司费用分摊": "管理公司费用分摊",
        "奖金":       "奖金",
        "税费":       "税费",
        # 财务费用
        "手续费":     "手续费",
    }

    def __init__(self, config: dict):
        """
        :param config: 全局配置字典（从settings.yaml加载）
        """
        self.config = config
        self.db_config = config["database"]
        self.template_config = config["template"]
        self.output_config = config["output"]
        self.account_filter = config["account_filter"]
        self.report_year = config["report_year"]
        self.report_months = config.get("report_months", [])

        # 核心组件
        self.connector = DatabaseConnector(self.db_config)
        self.ufsystem = UFSystemQuerier(self.connector)
        self.extractor = T3DataExtractor(self.connector)
        self._template_parser: Optional[TemplateParser] = None
        self.template_path = self._resolve_template_path()

        # 输出
        self._wb: Optional[Workbook] = None

    @property
    def template_parser(self) -> TemplateParser:
        if self._template_parser is None:
            # 优先取配置中指定的 source_sheet，未指定则传空字符串，
            # TemplateParser 内部会自动取第一个sheet
            sheet_name = self.template_config.get("source_sheet", "")
            self._template_parser = TemplateParser(
                self.template_path,
                sheet_name
            )
        return self._template_parser

    def _resolve_template_path(self) -> str:
        """解析模板文件路径（支持相对/绝对路径）"""
        path = self.template_config["filepath"]
        if not os.path.isabs(path):
            # 相对于项目根目录
            root = os.path.dirname(os.path.dirname(
                os.path.dirname(__file__)))
            path = os.path.join(root, path)
        if not os.path.exists(path):
            raise FileNotFoundError(f"模板文件不存在: {path}")
        return path

    # ================================================================
    # 方法 build_framework — 生成报表初步框架雏形（V2.0 新增）
    # ================================================================
    # 功能说明：
    #   以"分店财务报表模板.xlsx"为雏形，从 UA_Account 表查询所有
    #   合规账套（cacc_name字段），每个账套复制一个表页，表页名称
    #   即为账套名称，在每个表页的 A2 位置用红色字体填入账套名称。
    #
    # 与现有 build() 方法的区别：
    #   - build()：完整流程（查询模板 → 取数 → 填充数据 → 出报表）
    #   - build_framework()：仅做框架（复制模板页签 + 写入账套名），
    #     不填充任何财务数据，供客户预览效果
    # ================================================================

    def build_framework(self, accounts: list = None) -> str:
        """
        生成报表框架雏形（V2.0新增）
        ========================
        功能：以模板为基础，为每个合规账套复制一个表页，
             在A2位置用红色字体填入账套名称，形成报表初步框架。
             生成的工作簿不保留模板原sheet，直接按账套名称命名各表页。

        流程（基于 openpyxl.copy_worksheet 在同一工作簿内复制）：
          1. 加载模板文件，获取模板sheet（取第一个sheet，或配置中指定的sheet）
          2. 如果未传入 accounts 参数，则从 UA_Account 表自动查询
          3. 对每个账套，用 copy_worksheet() 复制模板sheet，
             再以账套名重命名新sheet
          4. 所有账套复制完毕后，删除原始的模板sheet
          5. 在每个新sheet的 A2 单元格写入账套名称并设为红色字体
          6. 保存到 output/ 目录

        参数:
            accounts: 指定的账套列表（可选），
                      不传则自动查询全部合规账套（过滤 998/999）

        返回:
            str: 生成的Excel文件绝对路径
        """
        logger.info("=" * 60)
        logger.info("开始生成报表框架雏形")
        logger.info(f"模板: {self.template_path}")

        # ---- 步骤1：如果未传入账套列表，则自动从 UA_Account 查询 ----
        if accounts is None:
            accounts = self._get_accounts()
        logger.info(f"共查询到 {len(accounts)} 个合规账套")

        if not accounts:
            logger.warning("未查询到任何账套，请检查数据库配置！")
            raise ValueError("未查询到账套信息")

        # ---- 步骤2：加载模板工作簿 ----
        # 直接使用模板工作簿，在其内部用 copy_worksheet 复制sheet
        # 这是 openpyxl 原生支持的同一工作簿内复制方式，不会出现
        # 跨工作簿的 StyleProxy 兼容性问题
        self._wb = load_workbook(self.template_path)
        # 获取源sheet名称：优先取配置中指定的 sheet 名，
        # 如果配置未指定或找不到，则取模板文件中的第一个 sheet
        template_sheet_name = self.template_config.get("source_sheet", "")
        if not template_sheet_name or template_sheet_name not in self._wb.sheetnames:
            template_sheet_name = self._wb.sheetnames[0]
            logger.info(f"自动获取模板sheet名称: {template_sheet_name}")
        template_ws = self._wb[template_sheet_name]

        # ---- 步骤3：遍历每个账套，用 copy_worksheet 复制sheet ----
        for acc in accounts:
            sheet_name = acc.sheet_name[:31]  # Excel sheet名最多31字符
            logger.info(f"  处理账套: {acc.cAcc_Name} (ID={acc.cAcc_Id})")
            # copy_worksheet 在同一工作簿内复制全部内容和样式（原生支持）
            new_ws = self._wb.copy_worksheet(template_ws)
            new_ws.title = sheet_name

        # ---- 步骤4：删除原始的模板sheet，只保留账套sheet ----
        if template_sheet_name in self._wb.sheetnames:
            del self._wb[template_sheet_name]

        # ---- 步骤5：遍历每个账套sheet，在A2填入账套名称并设红色字体 ----
        for acc in accounts:
            sheet_name = acc.sheet_name[:31]
            if sheet_name in self._wb.sheetnames:
                ws = self._wb[sheet_name]
                a2_cell = ws.cell(row=2, column=1)
                a2_cell.value = acc.cAcc_Name
                a2_cell.font = Font(
                    name="微软雅黑", size=11, bold=True,
                    color="FF0000"  # 红色
                )

        # ---- 步骤6：保存工作簿到输出目录 ----
        output_path = self._save_framework_workbook()

        logger.info(f"框架雏形生成完成，共 {len(accounts)} 个表页")
        logger.info(f"文件路径: {output_path}")
        return output_path

    def build(self) -> str:
        """
        执行完整的报表生成流程
        :return: 生成的Excel文件路径
        """
        logger.info("=" * 60)
        logger.info("开始生成财务报表")
        logger.info(f"年份: {self.report_year}")
        logger.info(f"模板: {self.template_path}")

        # 1. 解析模板布局
        template = self._parse_template()
        logger.info(f"模板解析完成: {len(template.row_labels)}行, "
                    f"{len(template.columns)}列")

        # 2. 查询所有账套
        accounts = self._get_accounts()
        logger.info(f"共查询到 {len(accounts)} 个账套")

        if not accounts:
            logger.warning("未查询到任何账套，请检查数据库配置！")
            raise ValueError("未查询到账套信息")

        # 3. 创建工作簿并填充
        self._create_workbook(template, accounts)
        output_path = self._save_workbook()

        # 4. 可选：自动打开文件
        if self.output_config.get("open_when_done", True):
            self._open_file(output_path)

        logger.info(f"报表生成完成: {output_path}")
        return output_path

    def _parse_template(self) -> TemplateLayout:
        """解析模板文件"""
        return self.template_parser.parse()

    def _get_accounts(self) -> list[AccountInfo]:
        """获取需要生成报表的账套列表"""
        include = self.account_filter.get("include_ids", [])
        exclude = self.account_filter.get("exclude_ids", [])

        if include:
            accounts = []
            for aid in include:
                acc = self.ufsystem.get_account_by_id(aid)
                if acc:
                    accounts.append(acc)
                else:
                    logger.warning(f"账套 {aid} 未找到，已跳过")
            return accounts
        else:
            return self.ufsystem.get_filtered_accounts(exclude_ids=exclude)

    def _create_workbook(self, template: TemplateLayout,
                         accounts: list[AccountInfo]):
        """
        创建工作簿，为每个账套创建一个页签
        """
        # 加载模板作为基础样式参考
        template_wb = load_workbook(self.template_path)
        template_ws = template_wb[template.sheet_name]

        self._wb = Workbook()
        # 删除默认页
        self._wb.remove(self._wb.active)

        for acc in accounts:
            sheet_name = acc.sheet_name
            logger.info(f"  处理账套: {acc.cAcc_Name} ({acc.cAcc_Id})")

            # 构建数据库名（含年度）
            db_name = f"{acc.db_name}_{self.report_year}"

            # 创建新页签
            ws = self._wb.create_sheet(title=sheet_name)

            # 从模板复制结构和样式
            self._copy_sheet_structure(template_ws, ws, template)

            # 填充数据
            self._fill_sheet_data(ws, template, db_name)

        template_wb.close()
        logger.info(f"所有页签创建完成，共 {len(accounts)} 个")

    def _copy_sheet_structure(self, src_ws, dst_ws, template: TemplateLayout):
        """
        从模板页签复制结构和样式到目标页签
        （保留格式、行标签、表头，清空数据区域）
        """
        # 复制A列（行标签）和格式
        for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row,
                                    min_col=1, max_col=1):
            for cell in row:
                new_cell = dst_ws.cell(row=cell.row, column=1)
                new_cell.value = cell.value
                if cell.has_style:
                    # 按原始属性值重新构造 StyleProxy 对象，避免跨工作簿直接赋值
                    try:
                        new_cell.font = Font(
                            name=cell.font.name,
                            size=cell.font.size,
                            bold=cell.font.bold,
                            italic=cell.font.italic,
                            color=cell.font.color,
                        )
                    except Exception:
                        pass
                    try:
                        new_cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=cell.alignment.vertical,
                            wrap_text=cell.alignment.wrap_text,
                        )
                    except Exception:
                        pass
                    try:
                        new_cell.fill = PatternFill(
                            fill_type=cell.fill.fill_type,
                            start_color=cell.fill.start_color,
                            end_color=cell.fill.end_color,
                        )
                    except Exception:
                        pass

        # 复制表头行（第2行）
        for col_idx, col_header in enumerate(template.columns, 1):
            new_cell = dst_ws.cell(row=1, column=col_idx)
            new_cell.value = col_header.label
            new_cell.font = self.HEADER_FONT
            new_cell.fill = self.HEADER_FILL
            new_cell.alignment = Alignment(horizontal="center")
            new_cell.border = self.THIN_BORDER

        # 设置店名（第一行第一列）
        dst_ws.cell(1, 1).value = template.store_name
        dst_ws.cell(1, 1).font = self.TITLE_FONT

    # ================================================================
    # 数据填充核心逻辑
    # ================================================================
    def _fill_sheet_data(self, ws, template: TemplateLayout, db_name: str):
        """
        向指定页签填充数据
        :param ws: 目标工作表
        :param template: 模板布局
        :param db_name: 账套数据库名，如 UFDATA_001_2024
        """
        months = self.report_months or list(range(1, 13))

        try:
            # 批量提取各科目数据
            revenue = self.extractor.get_monthly_revenue(
                db_name, self.report_year, months
            )
            cost = self.extractor.get_monthly_cost(
                db_name, self.report_year, months
            )
            expenses = self.extractor.get_monthly_expenses(
                db_name, self.report_year, months
            )

            # 合并所有数据源，方便查找
            all_data = {}
            all_data.update(revenue)        # 各渠道收入
            all_data.update(cost)           # 成本
            all_data.update(expenses)       # 各项费用

            # 遍历模板的每一行，匹配并填充数据
            for row_label in template.row_labels:
                row = row_label.row_index
                label_text = self._normalize_label(row_label.label)

                # ---- 查找数据源 ----
                data_key = self._find_matching_key(label_text)
                data_source = all_data.get(data_key)

                if data_source is None and data_key == "__cost__":
                    data_source = cost.get("主营业务成本", {})

                if data_source is None:
                    continue  # 无对应数据源的跳过（如合计行、空白行）

                # ---- 填充各月数据 ----
                for col_header in template.columns:
                    if col_header.month > 0:
                        val = data_source.get(col_header.month, 0)
                        self._set_cell_value(ws, row, col_header.col_index, val)
                    elif col_header.is_total:
                        total = sum(
                            data_source.get(m, 0) for m in months
                        )
                        self._set_cell_value(ws, row, col_header.col_index, total)

            # ---- 补充计算行（如毛利率、费用率等） ----
            self._fill_calculated_rows(ws, template, db_name, months)

            logger.debug(f"  {db_name} 数据填充完成")

        except Exception as e:
            logger.error(f"  {db_name} 数据提取失败: {e}", exc_info=True)
            ws.cell(2, 1).value = f"[数据提取错误] {e}"

    def _normalize_label(self, label: str) -> str:
        """标准化标签文字，去除空格和括号变体"""
        return (label.replace(" ", "")
                     .replace("（", "(")
                     .replace("）", ")")
                     .replace("　", "")
                     .strip())

    def _find_matching_key(self, label_text: str) -> Optional[str]:
        """
        根据标准化后的标签文字，在映射表中查找匹配的数据键
        使用包含匹配（子串匹配），以适应 "（1）油菜花" 这种带前缀的写法
        """
        # 精确匹配
        if label_text in self.LABEL_TO_DATA_KEY:
            return self.LABEL_TO_DATA_KEY[label_text]

        # 模糊匹配：查找映射表的key是否在标签文字中
        for map_key, data_key in self.LABEL_TO_DATA_KEY.items():
            if map_key in label_text or label_text in map_key:
                return data_key

        return None

    def _fill_calculated_rows(self, ws, template: TemplateLayout,
                               db_name: str, months: list[int]):
        """
        填充需要计算的行，如毛利、毛利率、费用率、利润等
        """
        # 获取各月总收入、总成本
        monthly_total_revenue = {}
        monthly_total_cost = {}
        monthly_total_expense = {}

        revenue = self.extractor.get_monthly_revenue(db_name, self.report_year, months)
        cost_data = self.extractor.get_monthly_cost(db_name, self.report_year, months)
        expense_data = self.extractor.get_monthly_expenses(db_name, self.report_year, months)

        # 汇总收入
        for _, mdata in revenue.items():
            for m, v in mdata.items():
                monthly_total_revenue[m] = monthly_total_revenue.get(m, 0) + v

        # 汇总成本
        cost = cost_data.get("主营业务成本", {})
        for m, v in cost.items():
            monthly_total_cost[m] = monthly_total_cost.get(m, 0) + v

        # 汇总费用
        for _, mdata in expense_data.items():
            for m, v in mdata.items():
                monthly_total_expense[m] = monthly_total_expense.get(m, 0) + v

        # 遍历模板行，找到需要计算的行
        for row_label in template.row_labels:
            label_text = row_label.label.strip()
            row = row_label.row_index

            if "毛利" == label_text and "毛利率" not in label_text:
                # 毛利 = 总收入 - 总成本
                for col in template.columns:
                    if col.month > 0:
                        val = (monthly_total_revenue.get(col.month, 0)
                               - monthly_total_cost.get(col.month, 0))
                        self._set_cell_value(ws, row, col.col_index, val)
                    elif col.is_total:
                        total_rev = sum(monthly_total_revenue.values())
                        total_cost = sum(monthly_total_cost.values())
                        self._set_cell_value(ws, row, col.col_index,
                                             total_rev - total_cost)

            elif "毛利率" == label_text:
                for col in template.columns:
                    if col.month > 0:
                        rev = monthly_total_revenue.get(col.month, 0)
                        cst = monthly_total_cost.get(col.month, 0)
                        val = (rev - cst) / rev * 100 if rev != 0 else 0
                        self._set_cell_value(ws, row, col.col_index, val)
                        # 毛利率显示为百分比
                        ws.cell(row, col.col_index).number_format = '0.00"%"'
                    elif col.is_total:
                        total_rev = sum(monthly_total_revenue.values())
                        total_cost = sum(monthly_total_cost.values())
                        val = (total_rev - total_cost) / total_rev * 100 if total_rev != 0 else 0
                        self._set_cell_value(ws, row, col.col_index, val)
                        ws.cell(row, col.col_index).number_format = '0.00"%"'

            elif "利润" == label_text and "利润率" not in label_text and "总分红" not in label_text:
                # 利润 = 毛利 - 费用
                for col in template.columns:
                    if col.month > 0:
                        gross = (monthly_total_revenue.get(col.month, 0)
                                 - monthly_total_cost.get(col.month, 0))
                        val = gross - monthly_total_expense.get(col.month, 0)
                        self._set_cell_value(ws, row, col.col_index, val)
                    elif col.is_total:
                        total_gross = (sum(monthly_total_revenue.values())
                                       - sum(monthly_total_cost.values()))
                        total_exp = sum(monthly_total_expense.values())
                        self._set_cell_value(ws, row, col.col_index,
                                             total_gross - total_exp)

            elif "利润率" == label_text:
                for col in template.columns:
                    if col.month > 0:
                        rev = monthly_total_revenue.get(col.month, 0)
                        cst = monthly_total_cost.get(col.month, 0)
                        exp = monthly_total_expense.get(col.month, 0)
                        profit = (rev - cst - exp)
                        val = profit / rev * 100 if rev != 0 else 0
                        self._set_cell_value(ws, row, col.col_index, val)
                        ws.cell(row, col.col_index).number_format = '0.00"%"'

    def _set_cell_value(self, ws, row: int, col: int, value: float):
        """设置单元格数值并应用格式"""
        cell = ws.cell(row=row, column=col)
        cell.value = value if value != 0 else 0
        cell.font = self.DATA_FONT
        cell.number_format = self.CURRENCY_FORMAT
        cell.alignment = Alignment(horizontal="right")
        cell.border = self.THIN_BORDER

    def _save_workbook(self) -> str:
        """保存（完整报表）工作簿到输出文件"""
        output_dir = self.output_config.get("dir", "output")
        os.makedirs(output_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = self.output_config.get("filename_prefix", "分店财务报表_")
        ext = self.output_config.get("file_extension", ".xlsx")
        filename = f"{prefix}{self.report_year}_{timestamp}{ext}"
        output_path = os.path.join(output_dir, filename)

        self._wb.save(output_path)
        return os.path.abspath(output_path)

    # ================================================================
    # _save_framework_workbook — 保存框架雏形工作簿（V2.0 新增）
    # ================================================================
    # 专用于 build_framework() 的保存方法，与 _save_workbook() 的区别：
    #   - 文件名加 "_框架雏形" 后缀，与完整报表区分
    #   - 其余逻辑相同（输出目录、时间戳、扩展名）
    # ================================================================

    def _save_framework_workbook(self) -> str:
        """
        保存框架雏形工作簿到输出文件
        ==============================
        生成的文件名格式如：分店财务报表_框架雏形_2024_20260626_112233.xlsx
        保存后自动关闭工作簿释放资源。
        """
        output_dir = self.output_config.get("dir", "output")
        os.makedirs(output_dir, exist_ok=True)

        # 生成带"框架雏形"后缀的文件名，便于与完整报表区分
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = self.output_config.get("filename_prefix", "分店财务报表_")
        ext = self.output_config.get("file_extension", ".xlsx")
        filename = f"{prefix}框架雏形_{self.report_year}_{timestamp}{ext}"
        output_path = os.path.join(output_dir, filename)

        self._wb.save(output_path)
        logger.info(f"框架雏形文件已保存: {output_path}")
        return os.path.abspath(output_path)

    @staticmethod
    def _open_file(filepath: str):
        """尝试自动打开文件（Windows系统）"""
        try:
            os.startfile(filepath)
        except Exception as e:
            logger.warning(f"无法自动打开文件: {e}")
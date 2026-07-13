# -*- coding: utf-8 -*-
"""
HTML报表渲染器
==============
将财务数据渲染为自包含的HTML文件，严格复刻Excel报表的表格结构和样式。
适用于在浏览器中预览、打印或做表内运算分析。

设计原则：
- 与Excel流程共用同一套数据提取器（T3DataExtractor），不重复查询数据库
- 严格复刻Excel的行顺序、行标签、背景色方案
- 输出单一HTML文件，内嵌CSS+JavaScript，不依赖外部资源
- 多账套通过JavaScript Tab切换展示
- 合计值由JavaScript自动计算（为未来扩展运算分析功能预留接口）

使用方式（由 ReportBuilder 在 output_format='html' 或 'both' 时调用）：
    renderer = HtmlRenderer(config)
    html_path = renderer.render(accounts, account_years, connector, extractor)
"""

import os
import logging
from datetime import datetime
from typing import Optional

from src.database.connector import DatabaseConnector
from src.database.t3data import T3DataExtractor

logger = logging.getLogger(__name__)


class HtmlRenderer:
    """
    HTML报表渲染器
    ==============
    职责：将财务数据渲染为自包含的HTML文件。
    """

    # ================================================================
    # 颜色常量（与 ReportBuilder 中的 FILL_* 严格对应）
    # 这些颜色直接映射到HTML表格行背景，确保HTML与Excel视觉效果一致
    # ================================================================
    COLOR_PERFORMANCE = "#E8D5F5"   # 业绩完成率 - 浅紫色
    COLOR_TARGET      = "#D5F5E3"   # 目标预算 - 浅绿色
    COLOR_SALES       = "#D6EAF8"   # 销售业绩 - 浅蓝色
    COLOR_COST        = "#FDEBD0"   # 主营业务成本 - 浅橙色
    COLOR_EXP_HEADER  = "#F9E79F"   # 营业/管理/财务费用（标题行）- 深黄色
    COLOR_EXPENSE     = "#FCF3CF"   # 费用明细行 - 浅黄色
    COLOR_RATE        = "#F2F3F4"   # 费用率 - 浅灰色
    COLOR_PROFIT      = "#AED6F1"   # 利润 - 中浅蓝色
    COLOR_DIVIDEND    = "#F5B7B1"   # 分红 - 浅粉色
    COLOR_BANK        = "#D2B4DE"   # 银行余额 - 淡紫色

    # ================================================================
    # 行标签背景色映射（与 ReportBuilder._apply_row_colors 完全一致）
    # ================================================================
    LABEL_COLOR_MAP = {
        "业绩完成率":           COLOR_PERFORMANCE,
        "目标预算（一档）":     COLOR_TARGET,
        "销售业绩":             COLOR_SALES,
        "主营业务成本":         COLOR_COST,
        "费用支出":             COLOR_RATE,      # 费用支出 - 浅灰色
        "费用率":               COLOR_RATE,
        "利润":                 COLOR_PROFIT,
        "利润率":               COLOR_PROFIT,
        "分红":                 COLOR_DIVIDEND,
        "总分红（往年分红万）": COLOR_DIVIDEND,
        "投资":                 COLOR_DIVIDEND,
        "银行余额":             COLOR_BANK,
    }

    # ================================================================
    # 标准科目编码映射（与 T3DataExtractor 一致）
    # 用于从 GL_AccSum 表查询各科目的月度发生额
    # ================================================================
    # ================================================================
    # 注意：不再使用 6001/6401/6601 硬编码科目体系
    # 所有收入/费用科目均通过 Code 表动态查询 5101/5501/5502 开头的科目
    # 参见 T3DataExtractor.get_revenue_subjects_from_code() 等方法
    # ================================================================

    def __init__(self, config: dict):
        """
        :param config: 全局配置字典（从 settings.yaml 加载）
        """
        self.config = config
        self.output_config = config["output"]
        self.template_config = config["template"]
        self.report_year = config["report_year"]
        self.report_months = config.get("report_months", [])

        # 月份标题映射
        self.MONTH_TITLES = {
            1: "1月", 2: "2月", 3: "3月", 4: "4月",
            5: "5月", 6: "6月", 7: "7月", 8: "8月",
            9: "9月", 10: "10月", 11: "11月", 12: "12月",
        }

    # ================================================================
    # 主入口方法
    # ================================================================

    def render(self, accounts, account_years: dict[str, int],
               db_connector: DatabaseConnector,
               extractor: T3DataExtractor) -> str:
        """
        主入口：渲染完整 HTML 报表文件。

        流程：
          1. 遍历每个账套
          2. 查询数据库获取该账套的科目数据、已结账月份、期间损益结转等
          3. 构建行标签列表和列结构（与Excel完全一致）
          4. 填充数据到表格单元格
          5. 生成包含所有账套 Tab 切换的完整 HTML

        :param accounts: AccountInfo 列表（与 build_framework 传入的相同）
        :param account_years: {账套号: 年度}
        :param db_connector: DatabaseConnector 实例（用于查询已结账月份）
        :param extractor: T3DataExtractor 实例（用于提取财务数据）
        :return: 生成的 HTML 文件绝对路径
        """
        months = self.report_months or list(range(1, 13))

        # 构建所有账套的表格数据
        sheets_data = []  # 列表，每个元素是一个账套的完整表格数据

        for acc in accounts:
            sheet_name = acc.sheet_name
            year_str = str(account_years.get(acc.cAcc_Id, "")) if account_years else ""

            # 构建数据库名：UFDATA_账套号_年度
            db_name = f"UFDATA_{acc.cAcc_Id.zfill(3)}_{year_str}" if year_str else ""
            if not db_name:
                logger.warning(f"HtmlRenderer: 账套 {acc.cAcc_Name} 无年度信息，跳过")
                continue

            logger.info(f"HtmlRenderer: 处理账套 {sheet_name} ({db_name})")

            # ---- 1. 查询已结账月份（确定列结构） ----
            closed_months = self._get_closed_months(db_connector, db_name, months)

            # ---- 2. 查询 5101/5501/5502 科目列表（确定行结构） ----
            revenue_subjects = extractor.get_revenue_subjects_from_code(db_name)
            expense_subjects = extractor.get_expense_subjects_from_code(db_name)
            manage_subjects = extractor.get_manage_subjects_from_code(db_name)

            # ---- 3. 构建行标签列表 ----
            rows = self._build_row_labels(revenue_subjects, expense_subjects,
                                          manage_subjects)

            # ---- 4. 构建列标题 ----
            columns = self._build_columns(closed_months)

            # ---- 5. 填充单元格数据（标准科目余额 + 期间损益结转 + 银行余额） ----
            cell_data = self._fill_cell_data(
                extractor, db_connector, db_name, int(year_str),
                rows, columns, closed_months
            )

            # ---- 6. 计算汇总行合计 ----
            self._calculate_summary_sums(cell_data, rows, columns)

            sheets_data.append({
                "sheet_name": sheet_name,
                "acc_name": acc.cAcc_Name,
                "year_str": year_str,
                "rows": rows,
                "columns": columns,
                "cell_data": cell_data,
            })

        if not sheets_data:
            raise ValueError("HtmlRenderer: 无有效账套数据，无法生成HTML报表")

        # ---- 7. 渲染完整 HTML ----
        html_content = self._render_html(sheets_data)

        # ---- 8. 保存文件 ----
        output_path = self._save_html(html_content)

        return output_path

    # ================================================================
    # 列结构构建
    # ================================================================

    def _get_closed_months(self, db_connector: DatabaseConnector,
                           db_name: str, default_months: list[int]) -> list[int]:
        """
        查询已结账月份，查询失败则使用默认月份列表。
        :return: 升序排列的月份列表
        """
        try:
            closed = db_connector.get_closed_periods(db_name)
            if closed:
                return closed
        except Exception as e:
            logger.warning(f"  {db_name} 查询已结账月份失败: {e}")
        return sorted(default_months)

    def _build_columns(self, closed_months: list[int]) -> list[dict]:
        """
        构建列标题列表，根据已结账月份确定显示哪些月份列。

        列结构（与Excel模板完全一致）：
          - 每月一列，列标题为 "X月"
          - 最后一列为"年度合计"

        :param closed_months: 已结账月份列表（升序）
        :return: [{"key": "month_1", "label": "1月", "month": 1, "is_total": False}, ...]
        """
        columns = []
        for m in sorted(closed_months):
            columns.append({
                "key": f"month_{m}",
                "label": self.MONTH_TITLES[m],
                "month": m,
                "is_total": False,
            })
        columns.append({
            "key": "total",
            "label": "年度合计",
            "month": 0,
            "is_total": True,
        })
        return columns

    # ================================================================
    # 行结构构建（严格复刻Excel模板的行顺序和分组）
    # ================================================================

    def _build_row_labels(self, revenue_subjects: list[dict],
                          expense_subjects: list[dict],
                          manage_subjects: list[dict]) -> list[dict]:
        """
        构建完整的行标签列表（与 Excel 模板的结构完全一致）。

        行顺序：
          1.  业绩完成率
          2.  目标预算（一档）
          3.  销售业绩
          4.  revenue_subjects (5101 各渠道收入明细，动态从数据库读取)
          5.  主营业务成本
          6.  毛利
          7.  毛利率
          8.  营业费用
          9.  expense_subjects (5501 营业费用明细，动态，各间隔一空白行)
         10.  管理费用
         11.  manage_subjects (5502 管理费用明细，动态，各间隔一空白行)
         12.  财务费用
         13.  财务费用明细（5503）
         14.  费用率
         15.  利润
         16.  利润率
         17.  分红
         18.  总分红（往年分红万）
         19.  投资
         20.  银行余额

        :return: [{"label": "...", "row_type": "...",
                    "is_calculated": bool}, ...]
        """
        rows = []

        # ---- 顶部固定行 ----
        rows.append({"label": "业绩完成率", "row_type": "performance",
                     "is_calculated": True})
        rows.append({"label": "目标预算（一档）", "row_type": "target",
                     "is_calculated": False})
        rows.append({"label": "销售业绩", "row_type": "sales",
                     "is_calculated": True})

        # ---- 各渠道收入明细（5101 科目） ----
        # 仅使用从数据库动态查询的 5101 科目列表
        # 如果某账套无 5101 科目，则不插入任何收入明细行（与 Excel builder 行为一致）
        if revenue_subjects:
            for subj in revenue_subjects:
                rows.append({"label": subj['ccode_name'],
                             "row_type": "revenue_detail",
                             "is_calculated": False})

        rows.append({"label": "主营业务成本", "row_type": "cost",
                     "is_calculated": False})
        rows.append({"label": "毛利", "row_type": "gross_profit",
                     "is_calculated": True})
        rows.append({"label": "毛利率", "row_type": "gross_rate",
                     "is_calculated": True})

        # ---- 费用支出/费用率（V2.0 新增：复刻模板第13/14行） ----
        rows.append({"label": "费用支出", "row_type": "expense_total",
                     "is_calculated": True})
        rows.append({"label": "费用率", "row_type": "expense_rate",
                     "is_calculated": True})

        # ---- 营业费用标题行 ----
        rows.append({"label": "营业费用", "row_type": "expense_header",
                     "is_calculated": True})

        # ---- 营业费用明细（5501 科目 + 间隔空白行） ----
        if expense_subjects:
            for subj in expense_subjects:
                rows.append({"label": subj['ccode_name'],
                             "row_type": "expense_detail",
                             "is_calculated": False})
                # 每个科目行之后留一个空白行（与Excel行为一致）
                rows.append({"label": "", "row_type": "blank",
                             "is_calculated": False})
        else:
            # 默认营业费用科目列表
            for exp_name in ["广告费", "物料费", "设备", "折旧费", "房租",
                             "物业费", "电费", "修配费", "运杂费", "其他"]:
                rows.append({"label": exp_name,
                             "row_type": "expense_detail",
                             "is_calculated": False})
                rows.append({"label": "", "row_type": "blank",
                             "is_calculated": False})

        # ---- 管理费用标题行 ----
        rows.append({"label": "管理费用", "row_type": "manage_header",
                     "is_calculated": True})

        # ---- 管理费用明细（5502 科目 + 间隔空白行） ----
        if manage_subjects:
            for subj in manage_subjects:
                rows.append({"label": subj['ccode_name'],
                             "row_type": "manage_detail",
                             "is_calculated": False})
                rows.append({"label": "", "row_type": "blank",
                             "is_calculated": False})
        else:
            # 默认管理费用科目列表
            for mgr_name in ["工资", "办公费", "差旅费", "业务招待费",
                             "员工福利", "装修费", "开办费", "服务咨询费",
                             "社保", "管理公司费用分摊", "奖金", "税费"]:
                rows.append({"label": mgr_name,
                             "row_type": "manage_detail",
                             "is_calculated": False})
                rows.append({"label": "", "row_type": "blank",
                             "is_calculated": False})

        # ---- 财务费用标题行 ----
        rows.append({"label": "财务费用", "row_type": "finance_header",
                     "is_calculated": True})

        # ---- 财务费用明细（5503 科目，动态从期间损益结转数据填充） ----
        # 注：财务费用明细行不在此处硬编码添加，而是由 _query_period_transfer_data
        # 方法从 GL_AccVouch 的期间损益结转凭证中自动匹配 5503 开头的 ccode。
        # 若需要显式添加 5503 明细行，可将来通过 Code 表动态查询。

        # ---- 利润/利润率 ----
        rows.append({"label": "利润", "row_type": "profit",
                     "is_calculated": True})
        rows.append({"label": "利润率", "row_type": "profit_rate",
                     "is_calculated": True})

        # ---- 分红/投资/银行余额 ----
        rows.append({"label": "分红", "row_type": "dividend",
                     "is_calculated": False})
        rows.append({"label": "总分红（往年分红万）", "row_type": "total_dividend",
                     "is_calculated": False})
        rows.append({"label": "投资", "row_type": "invest",
                     "is_calculated": False})
        rows.append({"label": "银行余额", "row_type": "bank_balance",
                     "is_calculated": False})

        logger.debug(f"  _build_row_labels: 共 {len(rows)} 行")
        return rows

    # ================================================================
    # 数据填充（三源合并：标准科目余额 + 期间损益结转 + 银行余额）
    # ================================================================

    def _fill_cell_data(self, extractor: T3DataExtractor,
                        db_connector: DatabaseConnector,
                        db_name: str, year: int,
                        rows: list[dict], columns: list[dict],
                        closed_months: list[int]) -> dict:
        """
        从各数据源填充单元格数据，构建完整的 (row_idx, col_key) 到 value 映射。

        数据优先级（后面的覆盖前面的）：
          1. 标准科目余额（从 GL_AccSum 按月查询 md/mc 发生额）
          2. 期间损益结转数据（从 GL_AccVouch 查询，覆盖标准数据）
          3. 银行余额（从 GL_AccSum 查询 me 期末余额）
          4. 汇总行合计（由 _calculate_summary_sums 计算）

        :return: {(row_idx, col_key): float_value, ...}
        """
        cell_data = {}

        # ---- 1. 填充标准科目余额数据 ----
        standard_data = self._query_standard_subject_data(
            extractor, db_name, year, rows, columns
        )
        for key, val in standard_data.items():
            cell_data[key] = cell_data.get(key, 0) + val

        # ---- 2. 填充期间损益结转数据（覆盖标准数据） ----
        period_data = self._query_period_transfer_data(
            extractor, db_name, rows, columns
        )
        for key, val in period_data.items():
            cell_data[key] = val  # 期间损益数据优先级更高，覆盖而非累加

        # ---- 3. 填充银行余额 ----
        bank_balances = self._query_bank_balance(
            extractor, db_name, year, columns
        )
        bank_row_idx = None
        for i, row in enumerate(rows):
            if row["row_type"] == "bank_balance":
                bank_row_idx = i
                break
        if bank_row_idx is not None:
            for col in columns:
                if col["month"] > 0 and col["month"] in bank_balances:
                    cell_data[(bank_row_idx, col["key"])] = \
                        bank_balances[col["month"]]

        logger.info(f"  {db_name}: 共填充 {len(cell_data)} 个数据单元格")
        return cell_data

    def _query_standard_subject_data(self, extractor: T3DataExtractor,
                                      db_name: str, year: int,
                                      rows: list[dict],
                                      columns: list[dict]) -> dict:
        """
        从 GL_AccSum 表动态查询 5101/5401/5501/5502/5503 科目的月度发生额，
        按科目名称匹配行标签填充。

        注意：不再使用 6001/6401/6601 硬编码体系，也不调用已删除的
        get_monthly_revenue/cost/expenses 方法。所有科目均通过 Code 表
        动态查询。此方法作为期间损益结转数据的兜底补充。

        收入类（5101）-> 取 md（借方发生额）
        成本类（5401）-> 取 mc（贷方发生额）
        费用类（5501/5502/5503）-> 取 mc（贷方发生额）

        :return: {(row_idx, col_key): value, ...}
        """
        result = {}
        months = [col["month"] for col in columns if col["month"] > 0]
        if not months:
            return result

        # 构建行标签到 row_idx 映射（标准化后，去除前导空格）
        label_row_map = {}
        for i, row in enumerate(rows):
            label = row["label"].strip()
            if label:
                label_row_map[label] = i

        # 定义要查询的科目前缀及其取值方向
        QUERY_PREFIXES = {
            "5101": "md",    # 主营业务收入 - 取借方
            "5102": "md",    # 其他业务收入 - 取借方
            "5401": "mc",    # 成本 - 取贷方
            "5501": "mc",    # 营业费用 - 取贷方
            "5502": "mc",    # 管理费用 - 取贷方
            "5503": "mc",    # 财务费用 - 取贷方
        }

        for prefix, direction in QUERY_PREFIXES.items():
            try:
                # 查询该科目前缀下所有子科目的全年总账数据
                rows_data = extractor.get_subject_balances(db_name, f"{prefix}%", year)
                if not rows_data:
                    continue

                # 按 ccode 汇总各月数据：{ccode: {月份: 金额}}
                subject_monthly = {}
                for row in rows_data:
                    ccode = str(row.get("cCode", "")).strip()
                    iperiod = int(row.get("iPeriod", 0))
                    val = float(row.get(direction, 0) or 0)
                    if not ccode or iperiod not in months:
                        continue
                    if ccode not in subject_monthly:
                        subject_monthly[ccode] = {}
                    subject_monthly[ccode][iperiod] = \
                        subject_monthly[ccode].get(iperiod, 0) + val

                if not subject_monthly:
                    continue

                # 批量查询科目名称
                all_ccodes = list(subject_monthly.keys())
                code_name_map = extractor.get_code_subject_name_batch(
                    db_name, all_ccodes
                )

                # 按科目名称匹配行标签并填充
                for ccode, monthly_vals in subject_monthly.items():
                    ccode_name = code_name_map.get(ccode, "")
                    if not ccode_name:
                        continue
                    row_idx = label_row_map.get(ccode_name)
                    if row_idx is None:
                        continue
                    for month_num, val in monthly_vals.items():
                        if val == 0:
                            continue
                        col_key = f"month_{month_num}"
                        key = (row_idx, col_key)
                        result[key] = result.get(key, 0) + val

            except Exception as e:
                logger.warning(f"  {db_name} 查询 {prefix} 科目数据失败: {e}")

        # ---- 特殊处理：如果"主营业务成本"行在 label_row_map 中但没有数据
        #      且期间损益结转也没有成本数据，尝试从 GL_AccSum 查 5401 单一级科目
        cost_label = "主营业务成本"
        if cost_label in label_row_map and not any(
            (label_row_map[cost_label], k) in result or
            (label_row_map[cost_label], k) in ((r, c) for r, c in result)
            for k in [f"month_{m}" for m in months]
        ):
            try:
                rows_data = extractor.get_subject_balances(db_name, "5401", year)
                cost_monthly = {}
                for row in rows_data:
                    ccode = str(row.get("cCode", "")).strip()
                    iperiod = int(row.get("iPeriod", 0))
                    val = float(row.get("mc", 0) or 0)  # 成本取贷方
                    if ccode == "5401" and iperiod in months:
                        cost_monthly[iperiod] = cost_monthly.get(iperiod, 0) + val
                row_idx = label_row_map[cost_label]
                for month_num, val in cost_monthly.items():
                    if val:
                        col_key = f"month_{month_num}"
                        result[(row_idx, col_key)] = \
                            result.get((row_idx, col_key), 0) + val
            except Exception as e:
                logger.debug(f"  {db_name} 查询主营业务成本（5401）失败: {e}")

        return result

    def _query_period_transfer_data(self, extractor: T3DataExtractor,
                                     db_name: str, rows: list[dict],
                                     columns: list[dict]) -> dict:
        """
        查询期间损益结转凭证数据（从 GL_AccVouch 表）。
        逻辑与 ReportBuilder._fill_period_transfer_data 完全一致。

        收入类科目（5101xx / 5102xx）--> 取 md（借方发生额）
        成本类科目（5401）--> 取 mc（贷方发生额）
        费用类科目（5501/5502/5503）--> 取 mc（贷方发生额）

        :return: {(row_idx, col_key): value, ...}
        """
        result = {}
        try:
            vouchers = extractor.get_period_transfer_vouchers(db_name)
            if not vouchers:
                return result

            # 批量查询科目名称
            all_ccodes = list(set(v["ccode"] for v in vouchers))
            code_name_map = extractor.get_code_subject_name_batch(
                db_name, all_ccodes
            )

            # 构建行标签到 row_idx 映射
            label_row_map = {}
            for i, row in enumerate(rows):
                label = row["label"].strip()
                if label:
                    label_row_map[label] = i

            # 构建月份到 col_key 映射
            month_col_map = {}
            for col in columns:
                if col["month"] > 0:
                    month_col_map[col["month"]] = col["key"]

            for vouch in vouchers:
                ccode = str(vouch.get("ccode", "")).strip()
                iperiod = int(vouch.get("iperiod", 0))
                md_val = float(vouch.get("md", 0) or 0)
                mc_val = float(vouch.get("mc", 0) or 0)

                # 查找科目名称
                ccode_name = code_name_map.get(ccode, "")
                if not ccode_name:
                    continue

                # 查找对应行号
                row_idx = label_row_map.get(ccode_name)
                if row_idx is None:
                    continue

                # 查找对应月份列
                col_key = month_col_map.get(iperiod)
                if col_key is None:
                    continue

                # 确定取值方向
                is_revenue = ccode.startswith(("5101", "5102"))
                is_cost = ccode.startswith("5401")
                is_expense = (ccode.startswith("5501") or
                              ccode.startswith("5502") or
                              ccode.startswith("5503"))

                if is_revenue:
                    val = md_val
                elif is_cost or is_expense:
                    val = mc_val
                else:
                    continue

                if val == 0:
                    continue

                # 累加（同一科目同一月份可能有多个凭证分录）
                key = (row_idx, col_key)
                result[key] = result.get(key, 0) + val

            logger.info(f"  {db_name}: 期间损益结转数据 {len(result)} 个单元格")
        except Exception as e:
            logger.warning(f"  {db_name} 期间损益结转查询失败: {e}")

        return result

    def _query_bank_balance(self, extractor: T3DataExtractor,
                            db_name: str, year: int,
                            columns: list[dict]) -> dict[int, float]:
        """
        查询银行余额数据（从 GL_AccSum 表取 1002 科目的 me 字段）。

        :param columns: 列结构（用于确定需要查询的月份列表）
        :return: {月份: 期末余额, ...}
        """
        months = [col["month"] for col in columns if col["month"] > 0]
        if not months:
            return {}

        try:
            balances = extractor.get_bank_balance_from_gl_accsum(
                db_name, year, months
            )
            logger.info(f"  {db_name}: 银行余额查询结果 {len(balances)} 个月")
            return balances
        except Exception as e:
            logger.warning(f"  {db_name} 银行余额查询失败: {e}")
            return {}

    # ================================================================
    # 汇总行合计计算（与 ReportBuilder._calculate_summary_sums 逻辑一致）
    # ================================================================

    def _calculate_summary_sums(self, cell_data: dict, rows: list[dict],
                                 columns: list[dict]):
        """
        计算各汇总行的合计值，包括：
          - 销售业绩行：取"销售业绩"与"主营业务成本"之间所有行的合计
          - 营业费用行：取"营业费用"与"管理费用"之间有标签行的合计
          - 管理费用行：取"管理费用"与"财务费用"之间有标签行的合计
          - 费用支出行：营业费用 + 管理费用 + 财务费用（各列的合计）
          - 费用率行：费用支出 / 销售业绩 * 100
          - 合计列：每行各月份之和
          - 利润行：销售业绩 - 主营业务成本 - 费用支出
          - 利润率行：利润 / 销售业绩 * 100

        :param cell_data: {(row_idx, col_key): value}，将被就地更新
        """
        # ---- 1. 定位关键行索引 ----
        key_indices = {}
        for i, row in enumerate(rows):
            rt = row["row_type"]
            key_indices[rt] = i

        sales_idx = key_indices.get("sales")
        cost_idx = key_indices.get("cost")
        exp_header_idx = key_indices.get("expense_header")
        mgr_header_idx = key_indices.get("manage_header")
        fin_header_idx = key_indices.get("finance_header")
        exp_total_idx = key_indices.get("expense_total")
        exp_rate_idx = key_indices.get("expense_rate")
        profit_idx = key_indices.get("profit")
        profit_rate_idx = key_indices.get("profit_rate")

        month_cols = [c for c in columns if c["month"] > 0]
        total_col = next((c for c in columns if c["is_total"]), None)

        def _get_val(row_i, col_key):
            return float(cell_data.get((row_i, col_key), 0))

        def _set_val(row_i, col_key, val):
            if val != 0:
                cell_data[(row_i, col_key)] = val

        # ---- 2. 销售业绩合计：区间内所有行的合计 ----
        if sales_idx is not None and cost_idx is not None and sales_idx < cost_idx:
            for col in month_cols:
                total = sum(
                    _get_val(r, col["key"])
                    for r in range(sales_idx + 1, cost_idx)
                )
                _set_val(sales_idx, col["key"], total)

        # ---- 3. 营业费用合计：区间内有标签行的合计 ----
        if exp_header_idx is not None and mgr_header_idx is not None \
                and exp_header_idx < mgr_header_idx:
            for col in month_cols:
                total = sum(
                    _get_val(r, col["key"])
                    for r in range(exp_header_idx + 1, mgr_header_idx)
                    if rows[r]["label"].strip()
                )
                _set_val(exp_header_idx, col["key"], total)

        # ---- 4. 管理费用合计：区间内有标签行的合计 ----
        if mgr_header_idx is not None and fin_header_idx is not None \
                and mgr_header_idx < fin_header_idx:
            for col in month_cols:
                total = sum(
                    _get_val(r, col["key"])
                    for r in range(mgr_header_idx + 1, fin_header_idx)
                    if rows[r]["label"].strip()
                )
                _set_val(mgr_header_idx, col["key"], total)

        # ---- 5. 费用支出合计 = 营业费用 + 管理费用 + 财务费用 ----
        if exp_total_idx is not None:
            all_cols = list(month_cols)
            if total_col:
                all_cols.append(total_col)
            for col in all_cols:
                exp_v = _get_val(exp_header_idx, col["key"]) if exp_header_idx else 0
                mgr_v = _get_val(mgr_header_idx, col["key"]) if mgr_header_idx else 0
                fin_v = _get_val(fin_header_idx, col["key"]) if fin_header_idx else 0
                total_expense = exp_v + mgr_v + fin_v
                _set_val(exp_total_idx, col["key"], total_expense)

        # ---- 6. 费用率 = 费用支出 / 销售业绩 * 100 ----
        if exp_rate_idx is not None:
            all_cols = list(month_cols)
            if total_col:
                all_cols.append(total_col)
            for col in all_cols:
                te = _get_val(exp_total_idx, col["key"]) if exp_total_idx else 0
                sv = _get_val(sales_idx, col["key"]) if sales_idx else 0
                rate_val = (te / sv * 100) if sv != 0 else 0
                _set_val(exp_rate_idx, col["key"], rate_val)

        # ---- 7. 所有行的合计列（费用支出/费用率/利润/利润率/银行余额行暂不处理，由后面覆盖） ----
        bank_idx = key_indices.get("bank_balance")
        if total_col is not None:
            skip_rows = {exp_total_idx, exp_rate_idx,
                         profit_idx, profit_rate_idx, bank_idx}
            for row_i in range(len(rows)):
                if row_i in skip_rows:
                    continue
                total = sum(_get_val(row_i, col["key"]) for col in month_cols)
                _set_val(row_i, total_col["key"], total)

        # ---- 8. 利润 = 销售业绩 - 成本 - 费用支出 ----
        if profit_idx is not None:
            all_cols = list(month_cols)
            if total_col:
                all_cols.append(total_col)
            for col in all_cols:
                sales_v = _get_val(sales_idx, col["key"]) if sales_idx else 0
                cost_v = _get_val(cost_idx, col["key"]) if cost_idx else 0
                te_v = _get_val(exp_total_idx, col["key"]) if exp_total_idx else 0
                profit_val = sales_v - cost_v - te_v
                _set_val(profit_idx, col["key"], profit_val)

        # ---- 9. 利润率 = 利润 / 销售业绩 * 100（百分比） ----
        if profit_rate_idx is not None:
            all_cols = list(month_cols)
            if total_col:
                all_cols.append(total_col)
            for col in all_cols:
                pv = _get_val(profit_idx, col["key"]) if profit_idx else 0
                sv = _get_val(sales_idx, col["key"]) if sales_idx else 0
                rate_val = (pv / sv * 100) if sv != 0 else 0
                _set_val(profit_rate_idx, col["key"], rate_val)

        logger.debug("  汇总行合计计算完成（含费用支出/费用率）")

    # ================================================================
    # HTML 渲染
    # ================================================================

    def _render_html(self, sheets_data: list[dict]) -> str:
        """
        渲染完整 HTML 字符串。
        包含：
          - <style> 内嵌 CSS（严格复刻 Excel 样式和行背景色）
          - <script> 内嵌 JavaScript（Tab切换 + 数字格式化）
          - 每个账套一个 Tab 面板的 <table>
        """
        css = self._build_css()
        js = self._build_js()

        # 构建 Tab 按钮和面板
        tabs_html = ""
        panels_html = ""

        for idx, sheet in enumerate(sheets_data):
            sheet_name = sheet["sheet_name"]
            active_class = " class='active'" if idx == 0 else ""
            tabs_html += (
                f'<button{active_class} onclick="switchTab({idx})">'
                f'{self._escape_html(sheet_name)}</button>\n'
            )

            panel_html = self._build_table(sheet, idx)
            show_class = " show" if idx == 0 else ""
            panels_html += (
                f'<div class="tab-panel{show_class}" id="panel_{idx}">\n'
                f'{panel_html}\n</div>\n'
            )

        timestamp = datetime.now().strftime("%Y年%m月%d日 %H:%M")

        html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>分店财务报表汇总 - {self.report_year}年</title>
<style>
{css}
</style>
</head>
<body>

<div class="container">
    <div class="report-header">
        <h1>分店财务报表汇总</h1>
        <div class="report-meta">
            <span>年度：{self.report_year}年</span>
            <span>生成时间：{timestamp}</span>
            <span>账套数：{len(sheets_data)}</span>
        </div>
    </div>

    <div class="tab-bar">
{tabs_html}    </div>

    <div class="tab-content">
{panels_html}    </div>

    <div class="report-footer">
        <span>编制：______</span>
        <span>审核：______</span>
        <span>签批：______</span>
    </div>
</div>

<script>
{js}
</script>

</body>
</html>"""
        return html

    def _build_css(self) -> str:
        """构建完整的 CSS 样式表（严格复刻 Excel 风格）"""
        return """/* ========== 全局样式 ========== */
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: "微软雅黑", "Microsoft YaHei", "宋体", sans-serif;
    background: #f5f5f5;
    color: #333;
    padding: 20px;
}
.container {
    max-width: 1400px;
    margin: 0 auto;
    background: #fff;
    border-radius: 8px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.1);
    overflow: hidden;
}

/* ========== 报表头部 ========== */
.report-header {
    padding: 20px 24px 12px;
    border-bottom: 2px solid #4472C4;
}
.report-header h1 {
    font-size: 22px;
    color: #2c3e50;
    font-weight: bold;
}
.report-meta {
    margin-top: 6px;
    font-size: 13px;
    color: #888;
}
.report-meta span { margin-right: 20px; }

/* ========== Tab 选项卡 ========== */
.tab-bar {
    display: flex;
    flex-wrap: wrap;
    background: #f0f4f8;
    border-bottom: 2px solid #4472C4;
    padding: 4px 8px 0;
}
.tab-bar button {
    padding: 8px 16px;
    margin: 0 2px;
    border: 1px solid transparent;
    border-bottom: none;
    background: transparent;
    cursor: pointer;
    font-size: 13px;
    font-family: inherit;
    color: #555;
    border-radius: 4px 4px 0 0;
    transition: all 0.2s;
    white-space: nowrap;
}
.tab-bar button:hover { background: #e0e7ef; color: #2c3e50; }
.tab-bar button.active {
    background: #fff;
    border-color: #4472C4;
    color: #4472C4;
    font-weight: bold;
}
.tab-panel { display: none; padding: 16px; overflow-x: auto; }
.tab-panel.show { display: block; }

/* ========== 表格通用样式 ========== */
table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    min-width: 800px;
}
thead th {
    background: #D9E1F2;
    font-weight: bold;
    font-size: 13px;
    border: 1px solid #999;
    padding: 6px 8px;
    text-align: center;
    white-space: nowrap;
}
thead th:first-child { text-align: left; min-width: 120px; }
tbody td {
    border: 1px solid #999;
    padding: 4px 6px;
    text-align: right;
    white-space: nowrap;
    font-variant-numeric: tabular-nums;
}
tbody td:first-child {
    text-align: left;
    font-weight: normal;
    min-width: 120px;
    padding-left: 8px;
}

/* ========== 行标签缩进层级 ========== */
td.indent-1 { padding-left: 24px !important; }
td.indent-2 { padding-left: 36px !important; }

/* ========== 行背景色（严格复刻 Excel 颜色方案） ========== */
.row-performance td { background: #E8D5F5; }
.row-target td { background: #D5F5E3; }
.row-sales td { background: #D6EAF8; }
.row-revenue-detail td { background: #EBF5FB; }     /* 收入明细 - 浅蓝（比销售业绩略浅，体现从属关系） */
.row-cost td { background: #FDEBD0; }
.row-gross-profit td { background: #D5F5E3; }       /* 毛利 - 浅绿 */
.row-gross-rate td { background: #F2F3F4; }         /* 毛利率 - 浅灰 */
.row-expense-header td { background: #F9E79F; }
.row-manage-header td { background: #F9E79F; }
.row-finance-header td { background: #F9E79F; }
.row-expense-detail td { background: #FCF3CF; }
.row-manage-detail td { background: #FCF3CF; }
.row-finance-detail td { background: #FCF3CF; }
.row-expense-total td { background: #F2F3F4; }   /* 费用支出 - 浅灰 */
.row-expense-rate td { background: #F2F3F4; }
.row-profit td { background: #AED6F1; }
.row-profit-rate td { background: #AED6F1; }
.row-dividend td { background: #F5B7B1; }
.row-total-dividend td { background: #F5B7B1; }
.row-invest td { background: #F5B7B1; }
.row-bank-balance td { background: #D2B4DE; }
.row-blank td { background: #ffffff; }

/* ========== 数字格式化 ========== */
td.num { text-align: right; }
td.num.zero { color: #bbb; }
td.num.negative { color: #c0392b; }

/* ========== 报表底部 ========== */
.report-footer {
    padding: 12px 24px 16px;
    border-top: 1px solid #ddd;
    display: flex;
    justify-content: space-around;
    font-size: 13px;
    color: #666;
    margin-top: 8px;
}

/* ========== 响应式适配 ========== */
@media (max-width: 768px) {
    body { padding: 8px; }
    .tab-bar button { padding: 6px 10px; font-size: 12px; }
    table { font-size: 12px; }
    thead th, tbody td { padding: 3px 4px; }
}"""

    def _build_table(self, sheet: dict, idx: int) -> str:
        """
        构建单个账套的 HTML 表格。
        
        :param sheet: 包含 rows, columns, cell_data, acc_name, year_str 的字典
        :param idx: 面板索引（用于生成唯一 ID）
        :return: 完整的 HTML 表格字符串
        """
        columns = sheet["columns"]
        rows = sheet["rows"]
        cell_data = sheet["cell_data"]
        acc_name = sheet["acc_name"]
        year_str = sheet["year_str"]

        # ---- 标题 ----
        title = f"{acc_name} {year_str}" if year_str else acc_name
        title_html = (
            f'<h2 style="font-size:16px;color:#4472C4;'
            f'margin-bottom:10px;">{self._escape_html(title)}</h2>\n'
        )

        # ---- 表头 ----
        header_cells = "<th>项目</th>\n"
        for col in columns:
            header_cells += f"<th>{self._escape_html(col['label'])}</th>\n"

        # ---- 数据行 ----
        # 定义需要缩进的 row_type
        INDENT_TYPES = {"revenue_detail", "expense_detail", "manage_detail", "finance_detail"}
        
        body_rows = ""
        for row_i, row in enumerate(rows):
            row_type = row["row_type"]
            label = row["label"]
            if not label:
                # 空白行显示为 &nbsp; 以保证行高
                label_display = "&nbsp;"
            else:
                label_display = self._escape_html(label)

            # 判断第一列是否需要缩进
            label_class = ""
            if row_type in INDENT_TYPES:
                label_class = ' class="indent-1"'

            cells = ""
            for col in columns:
                val = cell_data.get((row_i, col["key"]), None)
                if val is not None:
                    formatted = f"{val:,.2f}"
                    css_class = "num"
                    if val == 0:
                        css_class += " zero"
                    elif val < 0:
                        css_class += " negative"
                    cells += f'<td class="{css_class}">{formatted}</td>\n'
                else:
                    cells += "<td></td>\n"

            body_rows += (
                f'<tr class="row-{row_type.replace("_", "-")}">\n'
                f'<td{label_class}>{label_display}</td>\n{cells}'
                f'</tr>\n'
            )

        return (
            f'{title_html}'
            f'<table>\n'
            f'<thead>\n<tr>\n{header_cells}</tr>\n</thead>\n'
            f'<tbody>\n{body_rows}</tbody>\n'
            f'</table>\n'
        )

    def _build_js(self) -> str:
        """构建 JavaScript 代码（Tab切换 + 页面初始化）。"""
        return """// ============================================================
// Tab选项卡切换
// ============================================================
function switchTab(index) {
    // 更新按钮高亮状态
    var buttons = document.querySelectorAll('.tab-bar button');
    for (var i = 0; i < buttons.length; i++) {
        buttons[i].classList.toggle('active', i === index);
    }
    // 切换面板显示
    var panels = document.querySelectorAll('.tab-panel');
    for (var i = 0; i < panels.length; i++) {
        panels[i].classList.toggle('show', i === index);
    }
}

// ============================================================
// 页面加载后初始化
// ============================================================
document.addEventListener('DOMContentLoaded', function() {
    console.log('分店财务报表HTML已加载完成');
    var tabCount = document.querySelectorAll('.tab-bar button').length;
    console.log('共 ' + tabCount + ' 个账套页签');
});"""

    @staticmethod
    def _escape_html(text: str) -> str:
        """转义 HTML 特殊字符，防止 XSS 或页面渲染异常。"""
        if not text:
            return ""
        text = text.replace("&", "&")
        text = text.replace("<", "<")
        text = text.replace(">", ">")
        # 用字符串拼接构造 " 和 ' 避免 write_to_file 工具解析 HTML 实体
        text = text.replace('"', "&" + "quot;")
        text = text.replace("'", "&" + "#39;")
        return text

    # ================================================================
    # 文件保存
    # ================================================================

    def _save_html(self, html_content: str) -> str:
        """
        保存 HTML 内容到输出目录.
        文件名格式: 分店财务报表_2024_20260626_112233.html
        """
        output_dir = self.output_config.get("dir", "output")
        os.makedirs(output_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = self.output_config.get("filename_prefix", "分店财务报表_")
        filename = f"{prefix}{self.report_year}_{timestamp}.html"
        output_path = os.path.join(output_dir, filename)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        file_size = os.path.getsize(output_path)
        logger.info(f"HTML报表已保存: {output_path}")
        logger.info(f"  HTML文件大小: {file_size:,} 字节 ({file_size/1024:.1f} KB)")
        return os.path.abspath(output_path)
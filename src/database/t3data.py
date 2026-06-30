"""
用友T3财务数据提取模块
====================
从 T3 标准版数据库（UFDATA_XXX_YYYY）中提取财务报表所需的
科目余额、发生额等数据。

T3 核心表结构（SQL2008R2 版本）：
  - GL_AccVouch     : 凭证主表
  - GL_AccAss       : 辅助核算表
  - Code            : 科目表
  - GL_AccSum       : 科目总账（按科目+月份汇总）
"""

import logging
from typing import Optional
from src.database.connector import DatabaseConnector

logger = logging.getLogger(__name__)


class T3DataExtractor:
    """T3 财务数据提取器"""

    def __init__(self, connector: DatabaseConnector):
        self.connector = connector

    # ================================================================
    # 科目余额查询
    # ================================================================
    def get_subject_balance(self, db_name: str, subject_code: str,
                            year: int, month: int = 0) -> Optional[dict]:
        """
        查询指定科目某年月余额
        :param db_name: 账套数据库名，如 UFDATA_001_2024
        :param subject_code: 科目编码，如 "1001"（库存现金）
        :param year: 年份
        :param month: 月份（0=查询全年累计）
        :return: {code, name, begin_d, begin_c, end_d, end_c, ...}
        """
        if month == 0:
            sql = """
                SELECT cCode, cCode_Name,
                       mb=SUM(md), mc=SUM(mc),
                       md_f=SUM(md_f), mc_f=SUM(mc_f)
                FROM (
                    SELECT cCode, cCode_Name, md=md, mc=mc,
                           md_f=md_f, mc_f=mc_f
                    FROM GL_AccSum
                    WHERE cCode LIKE %s
                      AND iYear = %s
                      AND iPeriod BETWEEN 1 AND 12
                ) t
                GROUP BY cCode, cCode_Name
            """
            params = (f"{subject_code}%", year)
        else:
            sql = """
                SELECT cCode, cCode_Name, md, mc, md_f, mc_f
                FROM GL_AccSum
                WHERE cCode LIKE %s
                  AND iYear = %s
                  AND iPeriod = %s
            """
            params = (f"{subject_code}%", year, month)

        rows = self.connector.execute_query(sql, database=db_name,
                                            params=params)
        return rows[0] if rows else None

    def get_subject_balances(self, db_name: str, subject_code_like: str,
                             year: int) -> list[dict]:
        """
        查询科目族全年各月数据
        :param db_name: 账套数据库
        :param subject_code_like: 科目编码模糊匹配，如 "6001%" 或 "5%"
        :param year: 年份
        :return: [{cCode, cCode_Name, iPeriod, md, mc, ...}, ...]
        """
        sql = """
            SELECT cCode, cCode_Name, iPeriod,
                   md, mc, md_f, mc_f
            FROM GL_AccSum
            WHERE cCode LIKE %s
              AND iYear = %s
            ORDER BY cCode, iPeriod
        """
        return self.connector.execute_query(
            sql, database=db_name, params=(subject_code_like, year)
        )

    # ================================================================
    # 收入科目查询（常用的用友T3收入类科目）
    # ================================================================
    REVENUE_CODES = {
        "油菜花收入": "600101",
        "现金收入":   "600102",
        "美团收入":   "600103",
        "抖音收入":   "600104",
        "其他业务收入": "6051",
    }

    COST_CODE = "6401"       # 主营业务成本
    GROSS_CODE = None        # 毛利由计算得出

    # 费用科目
    EXPENSE_CODES = {
        # 营业费用 - 6601 销售费用
        "广告费":   "660101",
        "物料费":   "660102",
        "设备":     "660103",
        "折旧费":   "660104",
        "房租":     "660105",
        "物业费":   "660106",
        "电费":     "660107",
        "修配费":   "660108",
        "运杂费":   "660109",
        "其他":     "660110",
        # 管理费用 - 6602
        "工资":     "660201",
        "办公费":   "660202",
        "差旅费":   "660203",
        "业务招待费": "660204",
        "员工福利": "660205",
        "装修费":   "660206",
        "开办费":   "660207",
        "服务咨询费": "660208",
        "社保":     "660209",
        "管理公司费用分摊": "660210",
        "奖金":     "660211",
        "税费":     "660212",
        # 财务费用 - 6603
        "手续费":   "660301",
    }

    def get_monthly_revenue(self, db_name: str, year: int,
                            months: list[int] = None) -> dict:
        """
        按月获取各收入科目数据
        :return: {科目名: {1月: 值, 2月: 值, ...}, ...}
        """
        return self._get_subject_monthly(
            db_name, year, self.REVENUE_CODES, months
        )

    def get_monthly_cost(self, db_name: str, year: int,
                         months: list[int] = None) -> dict:
        """按月获取主营业务成本"""
        return self._get_subject_monthly(
            db_name, year, {"主营业务成本": self.COST_CODE}, months
        )

    def get_monthly_expenses(self, db_name: str, year: int,
                             months: list[int] = None) -> dict:
        """按月获取各项费用"""
        return self._get_subject_monthly(
            db_name, year, self.EXPENSE_CODES, months
        )

    def _get_subject_monthly(self, db_name: str, year: int,
                             code_map: dict,
                             months: list[int] = None) -> dict:
        """
        通用：按科目编码映射获取各月数据
        :param code_map: {显示名称: 科目编码}
        :param months: 指定月份列表，None=全部月份
        :return: {显示名称: {月份: 金额, ...}, ...}
        """
        result = {}
        for display_name, code in code_map.items():
            rows = self.get_subject_balances(db_name, f"{code}%", year)
            monthly = {}
            for row in rows:
                period = row["iPeriod"]
                if months and period not in months:
                    continue
                # 取贷方发生额 mc（收入类科目余额在贷方）
                monthly[period] = float(row.get("mc", 0) or 0)
            result[display_name] = monthly
        return result

    def get_monthly_sales_composition(self, db_name: str, year: int,
                                      months: list[int] = None) -> dict:
        """
        获取销售额构成（各收入渠道明细）
        用于填充模板中 (1)~(5) 各渠道收入行
        """
        return self._get_subject_monthly(
            db_name, year, self.REVENUE_CODES, months
        )

    def compute_monthly_gross_profit(self, db_name: str, year: int,
                                     months: list[int] = None) -> dict:
        """
        计算月毛利 = 总收入 - 成本
        """
        revenue = self.get_monthly_revenue(db_name, year, months)
        cost = self.get_monthly_cost(db_name, year, months)

        # 汇总总收入
        total_revenue = {}
        for name, mdata in revenue.items():
            for m, val in mdata.items():
                total_revenue[m] = total_revenue.get(m, 0) + val

        # 取成本
        cost_data = cost.get("主营业务成本", {})

        gross = {}
        for m in set(list(total_revenue.keys()) + list(cost_data.keys())):
            gross[m] = total_revenue.get(m, 0) - cost_data.get(m, 0)

        return gross

    # ================================================================
    # 科目表 Code 查询方法（用于动态行调整）
    # ================================================================
    def get_revenue_subjects_from_code(self, db_name: str) -> list[dict]:
        """
        从 code 表查询所有 ccode 以 5101 开头的 6 位科目。
        这些科目将作为"销售业绩"与"主营业务成本"之间的行显示。

        T3 Code 表结构：
          ccode       VARCHAR(40)   - 科目编码
          ccode_name  VARCHAR(60)   - 科目名称

        :param db_name: 账套数据库名，如 UFDATA_001_2024
        :return: [{ccode, ccode_name}, ...]，按 ccode 升序排列
        """
        # 注意：必须使用完全限定的表名 [{db_name}].dbo.Code，
        # 因为 connector.connect() 复用了已有连接，execute_query 的 database 参数
        # 在已有连接时不会切换到目标数据库。
        sql = f"""
            SELECT ccode, ccode_name
            FROM [{db_name}].dbo.Code
            WHERE ccode LIKE '5101%'
              AND LEN(ccode) = 6
              AND ccode NOT LIKE '%[^0-9]%'
            ORDER BY ccode
        """
        try:
            rows = self.connector.execute_query(sql, database=db_name)
            # 额外过滤：确保 ccode 确实是6位纯数字
            result = []
            for row in rows:
                code = str(row.get("ccode", "")).strip()
                name = str(row.get("ccode_name", "")).strip()
                if len(code) == 6 and code.isdigit() and code.startswith("5101"):
                    result.append({"ccode": code, "ccode_name": name})
            logger.info(f"  Code表查询 5101 科目: 共 {len(result)} 条: "
                        f"{[r['ccode_name'] for r in result]}")
            return result
        except Exception as e:
            logger.warning(f"  查询 Code 表（5101科目）失败: {e}")
            return []

    def get_expense_subjects_from_code(self, db_name: str) -> list[dict]:
        """
        从 code 表查询所有 ccode 以 5501 开头的 6 位科目。
        这些科目将作为"营业费用"与"管理费用"之间的行显示。

        :param db_name: 账套数据库名，如 UFDATA_001_2024
        :return: [{ccode, ccode_name}, ...]，按 ccode 升序排列
        """
        sql = f"""
            SELECT ccode, ccode_name
            FROM [{db_name}].dbo.Code
            WHERE ccode LIKE '5501%'
              AND LEN(ccode) = 6
              AND ccode NOT LIKE '%[^0-9]%'
            ORDER BY ccode
        """
        try:
            rows = self.connector.execute_query(sql, database=db_name)
            result = []
            for row in rows:
                code = str(row.get("ccode", "")).strip()
                name = str(row.get("ccode_name", "")).strip()
                if len(code) == 6 and code.isdigit() and code.startswith("5501"):
                    result.append({"ccode": code, "ccode_name": name})
            logger.info(f"  Code表查询 5501 科目: 共 {len(result)} 条: "
                        f"{[r['ccode_name'] for r in result]}")
            return result
        except Exception as e:
            logger.warning(f"  查询 Code 表（5501科目）失败: {e}")
            return []

    def get_manage_subjects_from_code(self, db_name: str) -> list[dict]:
        """
        从 code 表查询所有 ccode 以 5502 开头的 6 位科目。
        这些科目将作为"管理费用"与"财务费用"之间的行显示。

        :param db_name: 账套数据库名，如 UFDATA_001_2024
        :return: [{ccode, ccode_name}, ...]，按 ccode 升序排列
        """
        sql = f"""
            SELECT ccode, ccode_name
            FROM [{db_name}].dbo.Code
            WHERE ccode LIKE '5502%'
              AND LEN(ccode) = 6
              AND ccode NOT LIKE '%[^0-9]%'
            ORDER BY ccode
        """
        try:
            rows = self.connector.execute_query(sql, database=db_name)
            result = []
            for row in rows:
                code = str(row.get("ccode", "")).strip()
                name = str(row.get("ccode_name", "")).strip()
                if len(code) == 6 and code.isdigit() and code.startswith("5502"):
                    result.append({"ccode": code, "ccode_name": name})
            logger.info(f"  Code表查询 5502 科目: 共 {len(result)} 条: "
                        f"{[r['ccode_name'] for r in result]}")
            return result
        except Exception as e:
            logger.warning(f"  查询 Code 表（5502科目）失败: {e}")
            return []
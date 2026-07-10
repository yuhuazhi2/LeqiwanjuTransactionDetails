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

        注意：T3 SQL2008R2 版的 GL_AccSum 表无 iyear 和 cCode_Name 字段，
              年份由数据库名后缀（UFDATA_XXX_YYYY）承载，无需 WHERE 过滤；
              科目名称需通过 cCode 关联 Code 表获取。
              参见 get_bank_balance_from_gl_accsum() 注释。
        """
        if month == 0:
            sql = """
                SELECT cCode,
                       mb=SUM(md), mc=SUM(mc),
                       md_f=SUM(md_f), mc_f=SUM(mc_f)
                FROM GL_AccSum
                WHERE cCode LIKE %s
                  AND iPeriod BETWEEN 1 AND 12
                GROUP BY cCode
            """
            params = (f"{subject_code}%",)
        else:
            sql = """
                SELECT cCode, md, mc, md_f, mc_f
                FROM GL_AccSum
                WHERE cCode LIKE %s
                  AND iPeriod = %s
            """
            params = (f"{subject_code}%", month)

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

        注意：T3 SQL2008R2 版的 GL_AccSum 表无 iyear 和 cCode_Name 字段，
              年份由数据库名后缀（UFDATA_XXX_YYYY）承载，无需 WHERE 过滤；
              科目名称需通过 cCode 关联 Code 表获取。
        """
        sql = """
            SELECT cCode, iPeriod,
                   md, mc, md_f, mc_f
            FROM GL_AccSum
            WHERE cCode LIKE %s
            ORDER BY cCode, iPeriod
        """
        return self.connector.execute_query(
            sql, database=db_name, params=(subject_code_like,)
        )

    # ================================================================
    # 硬编码的科目代码已全部删除（不再使用6001/6401/6601科目体系）
    # 所有收入/费用科目均通过 Code 表动态查询 5101/5501/5502 开头的科目
    # 参见 get_revenue_subjects_from_code() 等方法
    # ================================================================


    # ================================================================
    # 科目表 Code 查询方法（用于动态行调整）
    # ================================================================
    def get_revenue_subjects_from_code(self, db_name: str) -> list[dict]:
        """
        从 code 表查询所有 ccode 以 5101 开头的科目。
        这些科目将作为"销售业绩"与"主营业务成本"之间的行显示。

        T3 Code 表结构：
          ccode       VARCHAR(40)   - 科目编码
          ccode_name  VARCHAR(60)   - 科目名称

        查询策略（V4.0 改进）：
          1. 优先查询 ccode LIKE '5101%' 且 LEN(ccode)=6 的明细科目
          2. 如果有6位明细科目，返回明细列表
          3. 如果没有6位明细科目，查询 5101 一级科目作为兜底返回

        :param db_name: 账套数据库名，如 UFDATA_001_2024
        :return: [{ccode, ccode_name}, ...]，按 ccode 升序排列
        """
        # 注意：必须使用完全限定的表名 [{db_name}].dbo.Code，
        # 因为 connector.connect() 复用了已有连接，execute_query 的 database 参数
        # 在已有连接时不会切换到目标数据库。
        def _query_subjects(like_pattern: str) -> list[dict]:
            """执行查询的辅助函数"""
            sql = f"""
                SELECT ccode, ccode_name
                FROM [{db_name}].dbo.Code
                WHERE ccode LIKE '{like_pattern}'
                  AND ccode NOT LIKE '%[^0-9]%'
                ORDER BY ccode
            """
            rows = self.connector.execute_query(sql, database=db_name)
            result = []
            for row in rows:
                code = str(row.get("ccode", "")).strip()
                name = str(row.get("ccode_name", "")).strip()
                if code.isdigit() and code.startswith("5101"):
                    result.append({"ccode": code, "ccode_name": name})
            return result

        try:
            # ---- 1. 优先查询6位明细科目 ----
            result = _query_subjects('5101%')
            result = [r for r in result if len(r['ccode']) == 6]

            if result:
                logger.info(f"  Code表查询 5101 明细科目: 共 {len(result)} 条: "
                            f"{[r['ccode_name'] for r in result]}")
                return result

            # ---- 2. 无6位明细，查询5101一级科目作为兜底 ----
            result = _query_subjects('5101')
            result = [r for r in result if r['ccode'] == '5101']
            if result:
                logger.info(f"  无5101明细科目，使用一级科目5101: {result[0]['ccode_name']}")
                return result

            # ---- 3. 连5101一级科目都没有，返回空列表 ----
            logger.info(f"  Code表查询 5101: 无任何5101科目记录")
            return []

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

    # ================================================================
    # GL_AccVouch 期间损益结转凭证查询（V2.6 新增）
    # ================================================================
    def get_period_transfer_vouchers(self, db_name: str) -> list[dict]:
        """
        从 GL_accvouch 表查询 cdigest='期间损益结转' 的所有凭证分录记录。
        这些记录用于在各店页签中填入期间损益结转后的实际发生额。

        查询字段:
          iperiod     - 会计期间（月份）
          ino_id      - 凭证编号
          inid        - 分录序号
          cdigest     - 摘要（期间损益结转）
          ccode       - 科目编码
          ccode_equal - 对方科目编码
          md          - 借方金额
          mc          - 贷方金额

        注意：SQL Server 2008R2 版 T3 中 GL_AccVouch 表无 iyear 字段，
              年份信息由数据库名 UFDATA_XXX_YYYY 承载。
              必须使用参数化查询 (%s) 传递中文条件，因为 pymssql
              通过 GBK 编码发送 SQL，硬编码在 SQL 中的中文字符
              可能被错误编码导致查询不到数据。

        :param db_name: 账套数据库名，如 UFDATA_007_2026
        :return: 字典列表，按 iperiod, ino_id, inid 升序排列
        """
        # 注意：必须使用 execute_query_odbc 而非 execute_query，
        # 因为 pymssql 在 charset=GBK 下对中文字符参数的编码有 Bug，
        # 会导致 cdigest = '期间损益结转' 条件匹配不到任何记录。
        # pyodbc 使用 ? 占位符能正确处理 Unicode 参数。
        sql = f"""
            SELECT iperiod, ino_id, inid, cdigest,
                   ccode, ccode_equal, md, mc
            FROM [{db_name}].dbo.GL_AccVouch
            WHERE cdigest = ?
            ORDER BY iperiod, ino_id, inid
        """
        try:
            rows = self.connector.execute_query_odbc(sql, database=db_name,
                                                     params=('期间损益结转',))
            logger.info(f"  查询期间损益结转凭证: 共 {len(rows)} 条分录")
            return rows
        except Exception as e:
            logger.warning(f"  查询 GL_AccVouch 期间损益结转失败: {e}")
            return []

    def get_code_subject_name(self, db_name: str,
                              ccode: str) -> str:
        """
        根据科目编码查询科目名称。

        :param db_name: 账套数据库名
        :param ccode: 科目编码
        :return: 科目名称（未找到返回空字符串）
        """
        sql = f"""
            SELECT ccode_name
            FROM [{db_name}].dbo.Code
            WHERE ccode = %s
        """
        try:
            rows = self.connector.execute_query(sql, database=db_name,
                                                params=(ccode,))
            if rows:
                return str(rows[0].get("ccode_name", "")).strip()
            return ""
        except Exception as e:
            logger.warning(f"  查询 Code 表科目名称失败 ccode={ccode}: {e}")
            return ""

    def get_code_subject_name_batch(self, db_name: str,
                                    ccode_list: list[str]) -> dict[str, str]:
        """
        批量根据科目编码查询科目名称。

        :param db_name: 账套数据库名
        :param ccode_list: 科目编码列表，如 ['510101','510102', ...]
        :return: {ccode: ccode_name, ...}
        """
        if not ccode_list:
            return {}
        # 去重
        unique_codes = list(set(ccode_list))
        placeholders = ','.join(['%s'] * len(unique_codes))
        sql = f"""
            SELECT ccode, ccode_name
            FROM [{db_name}].dbo.Code
            WHERE ccode IN ({placeholders})
        """
        try:
            rows = self.connector.execute_query(sql, database=db_name,
                                                params=tuple(unique_codes))
            result = {}
            for row in rows:
                code = str(row.get("ccode", "")).strip()
                name = str(row.get("ccode_name", "")).strip()
                if code:
                    result[code] = name
            return result
        except Exception as e:
            logger.warning(f"  批量查询 Code 表科目名称失败: {e}")
    # ================================================================
    # GL_AccSum 总账表查询 — 银行余额（V3.0 新增）
    # ================================================================
    # 功能说明：
    #   从 GL_accsum（科目总账）表中查询银行存款科目（1002）
    #   的期末余额 me 字段。用于填充报表最后一行"银行余额"的对应月份数据。
    #
    # GL_AccSum 表结构（用友T3标准版）：
    #   ccode       VARCHAR(40)   - 科目编码
    #   ccode_name  VARCHAR(60)   - 科目名称
    #   iyear       INT           - 会计年度
    #   iperiod     INT           - 会计期间（月份：1~12）
    #   mb          DECIMAL       - 期初余额
    #   md          DECIMAL       - 借方发生额
    #   mc          DECIMAL       - 贷方发生额
    #   me          DECIMAL       - 期末余额（核心字段，银行余额取此字段值）
    #   md_f        DECIMAL       - 借方累计
    #   mc_f        DECIMAL       - 贷方累计
    # ================================================================

    def get_bank_balance_from_gl_accsum(self, db_name: str,
                                        year: int,
                                        months: list[int] = None) -> dict[int, float]:
        """
        从 GL_accsum 总账表查询银行存款（1002科目）的各月期末余额 me。

        银行余额 = 银行存款（1002科目）的期末余额 me 字段。

        查询逻辑：
          1. 查询 ccode = '1002' 的 GL_accsum 记录（精确匹配一级科目）。
          2. 取 me（期末余额）字段，按 iperiod（月份）对应。
          3. 返回 {月份: 期末余额} 字典。

        注意：
          - 使用完全限定的表名 [{db_name}].dbo.GL_AccSum 进行查询，
            避免因连接复用导致的数据库切换问题。
          - 如果某月份查无记录，该月份不返回键值（由调用方处理）。
          - iperiod 字段的值与月份列直接对应：1月→1，2月→2，依此类推。

        :param db_name: 账套数据库名，如 UFDATA_007_2026
        :param year: 会计年度
        :param months: 需要查询的月份列表，如 [1,2,3,4,5]；
                       为 None 时查询全年 1~12 月
        :return: {月份(int): 期末余额(float)}，
                 如 {1: 12345.67, 2: 23456.78, ...}
        """
        if months is None:
            months = list(range(1, 13))

        # 银行余额 = 银行存款（1002 科目）的期末余额 me
        # 精确匹配一级科目 ccode = '1002'，不含下级明细科目
        # 注意：T3 SQL2008R2 版的 GL_AccSum 表无 iyear 字段，
        #       年份由数据库名后缀（UFDATA_XXX_YYYY）承载，无需 WHERE 过滤。
        #       该表可能跨年度包含记录，但通常各数据库只存储对应年度的数据。
        sql = f"""
            SELECT iperiod AS month_num, me
            FROM [{db_name}].dbo.GL_AccSum
            WHERE ccode = '1002'
              AND iperiod BETWEEN %s AND %s
            ORDER BY iperiod
        """

        try:
            min_month = min(months)
            max_month = max(months)
            params = (min_month, max_month)

            rows = self.connector.execute_query(sql, database=db_name,
                                                params=params)
            logger.info(f"  查询银行余额（GL_AccSum ccode=1002）: {db_name}, "
                        f"月份范围={min_month}~{max_month}, "
                        f"共 {len(rows)} 条记录")

            # 构建 {月份: 期末余额} 字典
            result = {}
            for row in rows:
                month_num = int(row["month_num"])
                balance = float(row["me"] or 0)
                result[month_num] = balance
                logger.debug(f"    月份 {month_num}: 期末余额={balance:.2f}")

            return result

        except Exception as e:
            logger.warning(f"  查询银行余额（GL_AccSum）失败: {db_name}, 错误={e}")
            return {}
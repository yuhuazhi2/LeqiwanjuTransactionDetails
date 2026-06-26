"""
UFSystem 账套信息查询模块
========================
从用友T3的公共数据库 UFSystem 中查询所有账套（核算单位）信息，
用于确定需要生成报表的门店列表及其对应的数据库名称。
"""

from typing import Optional
from src.database.connector import DatabaseConnector


class AccountInfo:
    """账套基本信息"""

    def __init__(self, cAcc_Id: str, cAcc_Name: str, iSysId: int,
                 cacc_path: str = "", db_name: str = ""):
        self.cAcc_Id = cAcc_Id          # 账套号，如 "001"
        self.cAcc_Name = cAcc_Name      # 账套名称，如 "三江店"
        self.iSysId = iSysId            # 系统ID
        self.cacc_path = cacc_path      # 账套路径
        self.db_name = db_name          # 对应的数据库名

    @property
    def sheet_name(self) -> str:
        """生成Excel工作表页签名（取账套名称前12位防止超长）"""
        name = self.cAcc_Name.strip()
        return name[:31] if name else f"账套{self.cAcc_Id}"

    def __repr__(self):
        return f"<Account {self.cAcc_Id}: {self.cAcc_Name}>"


class UFSystemQuerier:
    """
    UFSystem 账套查询器
    从 UFSystem 数据库的 UA_Account 表读取所有账套信息
    """

    # ================================================================
    # UFSystem 核心表结构（T3 11.2 标准版）
    # ----------------------------------------------------------------
    # UA_Account：账套主表
    #   cAcc_Id      VARCHAR(20)    - 账套号
    #   cAcc_Name    VARCHAR(100)   - 账套名称（即门店名称）
    #   iSysId       INT            - 系统内部ID
    #   cAcc_Path    VARCHAR(255)   - 账套物理路径
    # ================================================================

    TABLE_UA_ACCOUNT = "UA_Account"

    def __init__(self, connector: DatabaseConnector):
        self.connector = connector
        self._db_name = "UFSystem"

    def get_all_accounts(self) -> list[AccountInfo]:
        """
        获取UFSystem中所有启用的账套列表
        :return: AccountInfo 列表
        """
        sql = f"""
            SELECT cAcc_Id, cAcc_Name, iSysId, cAcc_Path
            FROM {self.TABLE_UA_ACCOUNT}
            ORDER BY iSysId
        """
        rows = self.connector.execute_query(sql, database=self._db_name)
        accounts = []
        for row in rows:
            acc = AccountInfo(
                cAcc_Id=str(row["cAcc_Id"]).strip(),
                cAcc_Name=str(row["cAcc_Name"]).strip(),
                iSysId=row["iSysId"],
                cacc_path=row.get("cAcc_Path", ""),
            )
            acc.db_name = self._build_db_name(acc.cAcc_Id)
            accounts.append(acc)
        return accounts

    def get_filtered_accounts(self, include_ids: list[str] = None,
                              exclude_ids: list[str] = None) -> list[AccountInfo]:
        """
        获取筛选后的账套列表
        :param include_ids: 只包含这些账套号（空列表=全部）
        :param exclude_ids: 排除这些账套号
        """
        all_accs = self.get_all_accounts()

        if include_ids:
            all_accs = [a for a in all_accs if a.cAcc_Id in include_ids]

        if exclude_ids:
            all_accs = [a for a in all_accs if a.cAcc_Id not in exclude_ids]

        return all_accs

    def get_account_by_id(self, acc_id: str) -> Optional[AccountInfo]:
        """根据账套号查询单个账套"""
        sql = f"""
            SELECT cAcc_Id, cAcc_Name, iSysId, cAcc_Path
            FROM {self.TABLE_UA_ACCOUNT}
            WHERE cAcc_Id = %s
        """
        rows = self.connector.execute_query(
            sql, database=self._db_name, params=(acc_id,)
        )
        if not rows:
            return None
        row = rows[0]
        acc = AccountInfo(
            cAcc_Id=str(row["cAcc_Id"]).strip(),
            cAcc_Name=str(row["cAcc_Name"]).strip(),
            iSysId=row["iSysId"],
            cacc_path=row.get("cAcc_Path", ""),
        )
        acc.db_name = self._build_db_name(acc.cAcc_Id)
        return acc

    @staticmethod
    def _build_db_name(acc_id: str) -> str:
        """
        根据账套号构建账套数据库名称
        用友T3标准命名规则：UFDATA_账套号_年度
        注：实际年度需从账套的年度信息中获取
        """
        return f"UFDATA_{acc_id.zfill(3)}"
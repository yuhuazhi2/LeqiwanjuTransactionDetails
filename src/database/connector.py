"""
SQL Server 数据库连接器（pyodbc 版）
==============================
自动尝试多种 ODBC 驱动连接 SQL Server 2008R2，
优先使用最新的可用驱动。
"""

import pyodbc
from typing import Optional
from contextlib import contextmanager
import logging

logger = logging.getLogger(__name__)

# ODBC 驱动候选列表（按优先顺序）
ODBC_DRIVERS_CANDIDATES = [
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "ODBC Driver 11 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server Native Client 10.0",
    "SQL Server",
]


def _get_available_drivers() -> list[str]:
    """获取系统中实际可用的 SQL Server ODBC 驱动列表（按优先顺序）"""
    installed = {d.strip() for d in pyodbc.drivers()}
    logger.debug(f"系统已安装的 ODBC 驱动: {installed}")

    available = []
    for candidate in ODBC_DRIVERS_CANDIDATES:
        if candidate in installed:
            available.append(candidate)

    # 如果候选列表一个都不在，仍尝试所有已安装的 SQL Server 驱动
    if not available:
        sql_drivers = [d for d in installed if "SQL" in d.upper() or "SERVER" in d.upper()]
        available.extend(sorted(sql_drivers, reverse=True))

    return available


def _build_connection_string(server: str, port: int, database: str,
                             user: str, password: str,
                             timeout: int) -> str:
    """构建 pyodbc 连接字符串，DRIVER 占位符由调用方替换"""
    conn_str = (
        f"DRIVER={{{{}}}};"
        f"SERVER={server},{port};"
        f"DATABASE={database};"
        f"UID={user};"
        f"PWD={password};"
        f"Connect Timeout={timeout};"
        f"Login Timeout={timeout};"
        f"Language=us_english;"
        f"APP=LeqiwanjuReport;"
        f"TrustServerCertificate=yes;"
        f"Encrypt=no;"
    )
    return conn_str


class DatabaseConnector:
    """SQL Server 数据库连接器（pyodbc / pymssql 实现，线程不安全，单次使用）"""

    def __init__(self, config: dict):
        """
        :param config: 数据库配置字典，需包含以下键：
            server, port, username, password, timeout, charset
        """
        self.config = config
        self._conn: Optional[object] = None
        self._drivers = _get_available_drivers()

    @property
    def is_connected(self) -> bool:
        if self._conn is None:
            return False
        try:
            # pyodbc.Connection 有 closed 属性
            if hasattr(self._conn, "closed"):
                return not self._conn.closed
            # pymssql.Connection 没有 closed 属性，尝试 ping
            try:
                cursor = self._conn.cursor()
                cursor.execute("SELECT 1")
                cursor.close()
                return True
            except Exception:
                return False
        except Exception:
            return False

    def connect(self, database: str = "") -> object:
        """建立数据库连接（自动重连 + 自动尝试多种驱动）"""
        if self.is_connected:
            try:
                cursor = self._conn.cursor()
                cursor.execute("SELECT 1")
                cursor.close()
                return self._conn
            except Exception:
                self.close()

        server = self.config["server"]
        port = self.config.get("port", 1433)
        user = self.config["username"]
        password = self.config["password"]
        timeout = self.config.get("timeout", 30)
        db = database or "master"

        # 先尝试用 pymssql（如果有的话）
        try:
            import pymssql
            pymssql_config = {
                "server": server,
                "port": port,
                "user": user,
                "password": password,
                "timeout": timeout,
                "charset": self.config.get("charset", "GBK"),
                "login_timeout": timeout,
                "tds_version": "7.0",
            }
            if database:
                pymssql_config["database"] = database
            logger.debug(f"正在尝试 pymssql 连接 {server}:{port}/{db}")
            conn = pymssql.connect(**pymssql_config)
            conn.autocommit(True)
            self._conn = conn
            logger.info(f"成功连接 pymssql -> {server}:{port}/{db}")
            return conn
        except ImportError:
            pass  # pymssql 未安装，跳过
        except Exception as e:
            logger.warning(f"pymssql 连接失败: {e}")

        # 用 pyodbc 自动尝试多种驱动
        # 注：旧驱动（SQL Server Native Client 10.0 / SQL Server）不支持
        # TrustServerCertificate 和 Encrypt 属性，必须去掉以免"无效的连接字符串属性"
        errors = []
        for driver in self._drivers:
            try:
                conn_str = (
                    f"DRIVER={{{driver}}};"
                    f"SERVER={server},{port};"
                    f"DATABASE={db};"
                    f"UID={user};"
                    f"PWD={password};"
                    f"Connect Timeout={timeout};"
                    f"Login Timeout={timeout};"
                )
                logger.debug(f"正在尝试驱动 [{driver}] 连接 {server}:{port}/{db}")
                conn = pyodbc.connect(conn_str, autocommit=True)
                self._conn = conn
                logger.info(f"成功连接 [{driver}] -> {server}:{port}/{db}")
                return conn

            except pyodbc.Error as e:
                error_text = str(e)
                logger.warning(f"驱动 [{driver}] 连接失败: {error_text}")
                errors.append(f"[{driver}] {error_text}")
                continue

        # 所有方式都失败
        error_summary = "\n".join(errors)
        raise ConnectionError(
            f"无法连接到数据库 {server}:{port}/{db}\n"
            f"已尝试 {len(self._drivers)} 个 ODBC 驱动:\n{error_summary}\n\n"
            f"请检查：\n"
            f"  1. SQL Server 服务是否运行\n"
            f"  2. 1433 端口是否可达\n"
            f"  3. 是否开启了 TCP/IP 协议\n"
            f"  4. 用户名/密码是否正确"
        )

    def close(self):
        """关闭连接"""
        if self._conn is None:
            return
        try:
            # 判断连接是否仍有效，避免重复关闭
            is_alive = False
            if hasattr(self._conn, "closed"):
                is_alive = not self._conn.closed
            else:
                # pymssql：尝试执行简单查询判断连接状态
                try:
                    cursor = self._conn.cursor()
                    cursor.execute("SELECT 1")
                    cursor.close()
                    is_alive = True
                except Exception:
                    is_alive = False

            if is_alive:
                self._conn.close()
        except Exception:
            pass
        finally:
            self._conn = None

    @contextmanager
    def cursor(self, database: str = ""):
        """获取游标的上下文管理器，自动释放资源"""
        conn = self.connect(database)
        cur = conn.cursor()
        try:
            yield cur
        finally:
            try:
                cur.close()
            except Exception:
                pass

    def execute_query(self, sql: str, database: str = "",
                      params: tuple = ()) -> list[dict]:
        """
        执行SQL查询，返回字典列表
        :param sql: SQL语句
        :param database: 目标数据库名
        :param params: SQL参数元组
        :return: 列名->值的字典列表
        """
        with self.cursor(database) as cursor:
            cursor.execute(sql, params)
            columns = [col[0] for col in cursor.description]
            rows = []
            for row in cursor.fetchall():
                rows.append(dict(zip(columns, row)))
            return rows

    def execute_many(self, sql: str, database: str = "",
                     params_list: list[tuple] = None):
        """批量执行SQL"""
        with self.cursor(database) as cursor:
            if params_list:
                cursor.executemany(sql, params_list)
            else:
                cursor.execute(sql)

    def get_closed_periods(self, db_name: str) -> list[int]:
        """
        查询指定账套数据库中已结账的月份（gl_mend 表 bflag=1 的 iperiod 值）
        :param db_name: 账套数据库名，如 "UFDATA_007_2026"
        :return: 已结账月份列表（升序），如 [1,2,3,4,5]
                 iperiod=0（年度结转）不返回

        注意：使用完全限定的表名 [{db_name}].dbo.gl_mend 而非依赖
        connect() 切换数据库，因为现有连接可能已固定在 UFSystem
        数据库上，execute_query 的 database 参数在连接已建立
        时不会生效（connect 方法会复用现有连接）。
        """
        sql = f"""
            SELECT iperiod FROM [{db_name}].dbo.gl_mend
            WHERE bflag = 1 AND iperiod > 0
            ORDER BY iperiod ASC
        """
        try:
            rows = self.execute_query(sql, database=db_name)
            return [row["iperiod"] for row in rows]
        except Exception as e:
            logger.warning(f"查询 {db_name}.gl_mend 失败: {e}")
            return []

    def execute_query_odbc(self, sql: str, database: str = "",
                           params: tuple = ()) -> list[dict]:
        """
        使用 pyodbc 直接执行 SQL 查询（绕过 pymssql 的 GBK 编码限制）。
        适用于 SQL 或参数中包含中文字符的场景，因为 pymssql 在
        charset=GBK 下对中文字符参数的编码可能导致查询无结果。

        注意：此方法会创建一个独立的 pyodbc 连接，不会影响主连接。
        每次执行后自动关闭该临时连接。

        :param sql: SQL 语句，参数占位符使用 ?（pyodbc 风格）
        :param database: 目标数据库名
        :param params: 参数元组
        :return: 列名->值的字典列表
        """
        server = self.config["server"]
        port = self.config.get("port", 1433)
        user = self.config["username"]
        password = self.config["password"]
        timeout = self.config.get("timeout", 30)
        db = database or "master"

        # 遍历驱动列表尝试连接
        last_error = None
        sql_safe = sql.replace("%s", "?")
        for driver in self._drivers:
            try:
                conn_str = (
                    f"DRIVER={{{driver}}};"
                    f"SERVER={server},{port};"
                    f"DATABASE={db};"
                    f"UID={user};"
                    f"PWD={password};"
                    f"Connect Timeout={timeout};"
                    f"Login Timeout={timeout};"
                )
                logger.debug(f"execute_query_odbc: 尝试驱动 [{driver}]")
                conn = pyodbc.connect(conn_str, autocommit=True)
                cursor = conn.cursor()
                cursor.execute(sql_safe, params)
                columns = [col[0] for col in cursor.description]
                rows = []
                for row in cursor.fetchall():
                    rows.append(dict(zip(columns, row)))
                cursor.close()
                conn.close()
                logger.debug(f"execute_query_odbc: [{driver}] 查询成功, {len(rows)} 行")
                return rows
            except Exception as e:
                last_error = e
                logger.debug(f"execute_query_odbc: 驱动 [{driver}] 失败: {e}")
                continue

        raise ConnectionError(
            f"execute_query_odbc: 无法连接到 {server}:{port}/{db}, "
            f"已尝试 {len(self._drivers)} 个驱动，最后错误: {last_error}"
        )

    def get_databases(self) -> list[str]:
        """获取服务器上所有数据库名称"""
        rows = self.execute_query(
            "SELECT name FROM sys.databases "
            "WHERE name NOT IN ('master','tempdb','model','msdb') "
            "ORDER BY name"
        )
        return [row["name"] for row in rows]

    def __del__(self):
        self.close()
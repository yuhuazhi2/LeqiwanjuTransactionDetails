#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
分店财务报表生成工具
==================
主程序入口，支持两种运行模式：
  1. 交互式GUI模式：双击运行 或 python main.py（默认启动）
  2. 命令行模式：python main.py --year 2024 --accounts 001 003 ...

GUI模式操作流程：
  1. 弹出对话界面，输入数据库连接信息（自动保留上次输入）
  2. 连接后从 UFSystem 查询所有账套列表
  3. 以多选勾选框展示账套（默认全选）
  4. 点击"生成合并汇总报表"执行生成
  5. 输出到 output/ 目录并提示打开

使用前请确保：
  - 已安装必要依赖（运行 install_deps.bat）
  - 模板文件 分店财务报表模板.xlsx 存在于项目根目录
"""

import os
import sys
import argparse
import traceback

# 确保项目根目录在 sys.path 中
_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)


def load_config(config_path: str = None) -> dict:
    """
    加载 YAML 配置文件
    :param config_path: 配置文件路径，默认 config/settings.yaml
    """
    import yaml

    if config_path is None:
        config_path = os.path.join(_PROJECT_ROOT,
                                   "config", "settings.yaml")

    if not os.path.exists(config_path):
        raise FileNotFoundError(
            f"配置文件不存在: {config_path}\n"
            f"请确保 config/settings.yaml 文件存在，"
            f"或使用 -c 参数指定配置文件路径。"
        )

    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    # 将配置文件路径中的相对路径转为绝对路径
    if not os.path.isabs(config.get("template", {}).get("filepath", "")):
        config["template"]["filepath"] = os.path.join(
            _PROJECT_ROOT, config["template"]["filepath"]
        )
    if not os.path.isabs(config.get("logging", {}).get("file", "")):
        log_file = config.get("logging", {}).get("file", "")
        if log_file:
            config["logging"]["file"] = os.path.join(_PROJECT_ROOT, log_file)
    if not os.path.isabs(config.get("output", {}).get("dir", "")):
        config["output"]["dir"] = os.path.join(
            _PROJECT_ROOT, config["output"]["dir"]
        )

    return config


def parse_args() -> argparse.Namespace:
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="分店财务报表生成工具 v2.0",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python main.py                         # 启动交互式GUI界面（默认）
  python main.py --cli ...               # 命令行模式（原有方式）
  python main.py --cli -c my_config.yaml
  python main.py --cli --year 2025 --months 1 2 3
  python main.py --cli --accounts 001 003
        """
    )
    parser.add_argument(
        "--cli", action="store_true",
        help="使用命令行模式运行（而非默认的GUI模式）"
    )
    parser.add_argument(
        "-c", "--config",
        default=os.path.join(_PROJECT_ROOT, "config", "settings.yaml"),
        help="配置文件路径（默认: config/settings.yaml）"
    )
    parser.add_argument(
        "--format", choices=["xlsx", "html", "both"], default=None,
        help="输出格式：xlsx（Excel）/ html（HTML）/ both（同时生成两种）"
    )
    parser.add_argument(
        "--year", type=int,
        help="报表年份（覆盖配置文件中的设置）"
    )
    parser.add_argument(
        "--months", type=int, nargs="+",
        help="月份范围，如 1 2 3（覆盖配置文件中的设置）"
    )
    parser.add_argument(
        "--accounts", type=str, nargs="+",
        help="指定账套号，如 001 003（覆盖配置文件中的筛选条件）"
    )
    parser.add_argument(
        "--no-open", action="store_true",
        help="生成后不自动打开文件"
    )
    parser.add_argument(
        "--output-dir", type=str,
        help="输出目录（覆盖配置文件中的设置）"
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="输出详细日志（DEBUG级别）"
    )
    return parser.parse_args()


def override_config_with_args(config: dict, args: argparse.Namespace) -> dict:
    """用命令行参数覆盖配置"""
    if args.year:
        config["report_year"] = args.year
    if args.months:
        config["report_months"] = list(set(args.months))
    if args.accounts:
        config["account_filter"]["include_ids"] = list(set(args.accounts))
    if args.format:
        config["output"]["format"] = args.format
    if args.no_open:
        config["output"]["open_when_done"] = False
    if args.output_dir:
        config["output"]["dir"] = args.output_dir
    if args.verbose:
        config["logging"]["level"] = "DEBUG"
    return config


def run_cli_mode(args: argparse.Namespace):
    """命令行模式：加载配置文件并执行报表生成"""
    config = load_config(args.config)
    config = override_config_with_args(config, args)

    # 初始化日志
    from src.utils.logger import setup_logging
    setup_logging(config["logging"])

    logger = __import__("logging").getLogger(__name__)

    try:
        # 创建报表生成器
        from src.report.builder import ReportBuilder
        from src.database.ufsystem import UFSystemQuerier
        from src.database.connector import DatabaseConnector

        builder = ReportBuilder(config)

        # 获取输出格式（从配置或命令行参数）
        output_format = config.get("output", {}).get("format", "xlsx")

        # 执行生成（使用 build_framework 替代已废弃的 build）
        # 查询账套列表
        accounts = builder._get_accounts()

        if not accounts:
            raise ValueError("未查询到任何账套，请检查数据库配置！")

        # 构建账套年份映射
        year = config.get("report_year", 2024)
        account_years = {acc.cAcc_Id: year for acc in accounts}

        output_paths = builder.build_framework(
            accounts=accounts,
            account_years=account_years,
            output_format=output_format
        )

        # 输出结果摘要
        print(f"\n{'=' * 60}")
        print(f"  ✅ 财务报表生成成功！")
        for path in output_paths:
            print(f"  📄 文件路径: {path}")
        print(f"{'=' * 60}")
        logger.info("程序执行完成")

    except FileNotFoundError as e:
        logger.error(str(e))
        print(f"\n❌ 错误: {e}")
        sys.exit(1)
    except ValueError as e:
        logger.error(str(e))
        print(f"\n❌ 错误: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"程序运行出错: {e}", exc_info=True)
        print(f"\n❌ 程序运行出错: {e}")
        traceback.print_exc()
        sys.exit(1)


def run_gui_mode():
    """GUI交互模式：启动图形界面"""
    try:
        from src.gui.app import MainWindow
        app = MainWindow()
        app.run()
    except Exception as e:
        print(f"\n❌ GUI启动失败: {e}")
        traceback.print_exc()
        sys.exit(1)


def main():
    """主函数 - 自动选择运行模式"""
    args = parse_args()

    # 判断是否使用命令行模式（--cli 参数或显式的配置/年份/账套等参数）
    # 注意：--format 参数在 CLI 模式下独立使用，不支持从 GUI 模式传递。
    # 当指定 --format 且没有同时指定 --cli 或账套等其他参数时，仍进入 GUI 模式。
    has_cli_args = any([
        args.config != os.path.join(_PROJECT_ROOT, "config", "settings.yaml"),
        args.year is not None,
        args.months is not None,
        args.accounts is not None,
        args.no_open,
        args.output_dir is not None,
        args.format is not None,
    ])

    if args.cli or has_cli_args:
        run_cli_mode(args)
    else:
        run_gui_mode()


if __name__ == "__main__":
    main()
"""
日志配置工具
==========
统一配置日志格式、输出级别和文件输出。
"""

import os
import logging
from logging.handlers import RotatingFileHandler


def setup_logging(config: dict):
    """
    根据配置初始化日志系统
    :param config: logging 配置子字典
    """
    log_config = config
    level = getattr(logging, log_config.get("level", "INFO").upper(), logging.INFO)
    log_file = log_config.get("file", "")
    console = log_config.get("console", True)

    # 根日志器
    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    # 清除已有处理器（防止重复添加）
    for h in root_logger.handlers[:]:
        root_logger.removeHandler(h)

    # 格式化器
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(filename)s:%(lineno)d | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # 文件处理器（带轮转，最大10MB保留5个备份）
    if log_file:
        log_dir = os.path.dirname(log_file)
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)
        file_handler = RotatingFileHandler(
            log_file, maxBytes=10 * 1024 * 1024,
            backupCount=5, encoding="utf-8"
        )
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)

    # 控制台处理器
    if console:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)

    logging.info("日志系统初始化完成，级别: %s", log_config.get("level", "INFO"))
# -*- coding: utf-8 -*-
"""
浏览器工具函数
=============
提供跨平台打开文件的能力，优先使用火狐浏览器打开 HTML 文件。
"""

import os
import subprocess
import logging

logger = logging.getLogger(__name__)


def _find_firefox_path() -> str | None:
    """
    查找火狐浏览器的可执行文件路径。

    搜索顺序：
      1. 注册表或常见安装路径（Windows）
      2. PATH 环境变量中的 firefox
      3. 回退 None

    :return: 火狐浏览器的完整路径，未找到则返回 None
    """
    # Windows 常见安装路径
    possible_paths = [
        os.path.join("C:\\", "Program Files", "Mozilla Firefox", "firefox.exe"),
        os.path.join("C:\\", "Program Files (x86)", "Mozilla Firefox", "firefox.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Mozilla Firefox", "firefox.exe"),
    ]

    for path in possible_paths:
        if os.path.isfile(path):
            logger.debug(f"找到火狐浏览器: {path}")
            return path

    # 尝试从 PATH 中查找
    try:
        # Windows 下 where 命令查找
        result = subprocess.run(
            ["where", "firefox"],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0 and result.stdout.strip():
            path = result.stdout.strip().splitlines()[0]
            if os.path.isfile(path):
                logger.debug(f"火狐浏览器(PATH): {path}")
                return path
    except Exception:
        pass

    logger.info("未找到火狐浏览器，将使用系统默认方式打开文件")
    return None


def open_file_with_firefox(filepath: str) -> bool:
    """
    使用火狐浏览器打开指定文件（HTML），
    找不到火狐时回退到系统默认打开方式。

    :param filepath: 文件路径（绝对路径）
    :return: 是否成功打开
    """
    if not os.path.isfile(filepath):
        logger.error(f"文件不存在: {filepath}")
        return False

    # 如果是 HTML 文件，优先使用火狐浏览器
    is_html = filepath.lower().endswith((".html", ".htm"))

    if is_html:
        firefox_path = _find_firefox_path()
        if firefox_path:
            try:
                # 使用 subprocess.Popen 非阻塞启动（与 os.startfile 行为一致）
                subprocess.Popen(
                    [firefox_path, filepath],
                    shell=False,
                    close_fds=True
                )
                logger.info(f"已使用火狐浏览器打开: {filepath}")
                return True
            except Exception as e:
                logger.warning(f"使用火狐浏览器打开失败: {e}，回退到系统默认方式")

    # 回退：使用系统默认方式打开
    try:
        if os.name == "nt":  # Windows
            os.startfile(filepath)
        elif os.name == "posix":  # macOS / Linux
            subprocess.Popen(
                ["open", filepath] if os.uname().sysname == "Darwin" else ["xdg-open", filepath],
                close_fds=True
            )
        logger.info(f"已使用系统默认方式打开: {filepath}")
        return True
    except Exception as e:
        logger.error(f"打开文件失败: {filepath}, 错误: {e}")
        return False
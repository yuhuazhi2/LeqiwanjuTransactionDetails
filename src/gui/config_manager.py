# -*- coding: utf-8 -*-
"""
配置管理器
=========
保存和加载上次数据库连接信息到 JSON 文件，
实现用户输入的自动持久化。

加载优先级：
1. 首先读取 connection_profile.json（上次GUI保存的配置）
2. 如果密码为空，回退到 settings.yaml 中的默认密码

PyInstaller 兼容处理：
- 写入操作（connection_profile.json）→ 写入可执行文件所在目录（可写）
- 读取操作（settings.yaml）→ 优先运行目录，其次资源目录
"""

import os
import sys
import json
import yaml


class ConnectionConfigManager:
    """管理数据库连接配置的持久化"""

    # ---- PyInstaller 兼容路径处理 ----
    @staticmethod
    def _get_run_dir():
        """获取可写配置目录（打包后为 exe 所在目录，开发模式为项目根目录）"""
        if getattr(sys, 'frozen', False):
            return os.path.join(os.path.dirname(sys.executable), "config")
        else:
            return os.path.join(
                os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                "config"
            )

    @staticmethod
    def _get_resource_dir():
        """获取资源配置目录（打包后为 _MEIPASS，开发模式同上）"""
        if getattr(sys, 'frozen', False):
            return os.path.join(sys._MEIPASS, "config")
        else:
            return ConnectionConfigManager._get_run_dir()

    @classmethod
    def get_config_dir(cls):
        """获取配置目录路径（可写）"""
        return cls._get_run_dir()

    @classmethod
    def get_config_file(cls):
        """获取 connection_profile.json 路径"""
        return os.path.join(cls._get_run_dir(), "connection_profile.json")

    @classmethod
    def get_settings_file(cls):
        """获取 settings.yaml 路径"""
        return cls._find_settings_file()

    @classmethod
    def _find_settings_file(cls):
        """查找 settings.yaml：优先运行目录，其次资源目录"""
        run_file = os.path.join(cls._get_run_dir(), "settings.yaml")
        if os.path.exists(run_file):
            return run_file
        res_file = os.path.join(cls._get_resource_dir(), "settings.yaml")
        if os.path.exists(res_file):
            return res_file
        # 最后尝试 settings.yaml.example（仅资源目录）
        example_file = os.path.join(cls._get_resource_dir(), "settings.yaml.example")
        if os.path.exists(example_file):
            return example_file
        return run_file  # fallback

    DEFAULT_CONFIG = {
        "server": "",
        "username": "sa",
        "password": "",
        "account_years": {},  # {账套号: 年份} — 上次选择的年份映射
    }

    @classmethod
    def _load_settings_defaults(cls) -> dict:
        """
        从 settings.yaml 读取默认数据库连接信息作为备用
        """
        settings_file = cls.get_settings_file()
        try:
            if os.path.exists(settings_file):
                with open(settings_file, "r", encoding="utf-8") as f:
                    settings = yaml.safe_load(f)
                db = settings.get("database", {})
                return {
                    "server": db.get("server", ""),
                    "username": db.get("username", "sa"),
                    "password": db.get("password", ""),
                }
        except Exception:
            pass
        return {}

    @classmethod
    def ensure_config_dir(cls):
        """确保配置目录存在"""
        os.makedirs(cls.get_config_dir(), exist_ok=True)

    @classmethod
    def load(cls) -> dict:
        """
        加载上次保存的连接配置
        如果密码为空，尝试从 settings.yaml 读取默认密码
        :return: 配置字典，包含 server, username, password, account_years
        """
        cls.ensure_config_dir()

        # 1. 基础默认值
        config = dict(cls.DEFAULT_CONFIG)

        # 2. 尝试从 settings.yaml 加载默认值
        settings_defaults = cls._load_settings_defaults()
        config.update(settings_defaults)

        # 3. 尝试读取 connection_profile.json（优先级最高）
        config_file = cls.get_config_file()
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    profile = json.load(f)
                # 合并：profile 中的非空值覆盖默认值
                for key in ("server", "username", "password"):
                    if profile.get(key):
                        config[key] = profile[key]
                # 加载上次保存的账套年份映射
                if profile.get("account_years"):
                    config["account_years"] = profile["account_years"]
            except (json.JSONDecodeError, IOError):
                pass

        return config

    @classmethod
    def save(cls, server: str, username: str, password: str,
             account_years: dict[str, int] = None):
        """
        保存连接配置到本地
        :param server: 服务器地址
        :param username: 登录用户名
        :param password: 登录密码
        :param account_years: {账套号: 年份} 映射（可选）
        """
        cls.ensure_config_dir()
        data = {
            "server": server,
            "username": username,
            "password": password,
        }
        # 始终写入 account_years 字段（即使是空字典），确保 JSON 结构完整
        data["account_years"] = account_years if account_years else {}
        try:
            config_file = cls.get_config_file()
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except IOError:
            pass  # 静默失败，不影响主流程

    @classmethod
    def clear(cls):
        """清除保存的配置"""
        cls.ensure_config_dir()
        try:
            config_file = cls.get_config_file()
            if os.path.exists(config_file):
                os.remove(config_file)
        except IOError:
            pass
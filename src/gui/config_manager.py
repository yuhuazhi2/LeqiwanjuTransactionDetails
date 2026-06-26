# -*- coding: utf-8 -*-
"""
配置管理器
=========
保存和加载上次数据库连接信息到 JSON 文件，
实现用户输入的自动持久化。

加载优先级：
1. 首先读取 connection_profile.json（上次GUI保存的配置）
2. 如果密码为空，回退到 settings.yaml 中的默认密码
"""

import os
import json
import yaml


class ConnectionConfigManager:
    """管理数据库连接配置的持久化"""

    CONFIG_DIR = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        "config"
    )
    CONFIG_FILE = os.path.join(CONFIG_DIR, "connection_profile.json")
    SETTINGS_FILE = os.path.join(CONFIG_DIR, "settings.yaml")

    DEFAULT_CONFIG = {
        "server": "",
        "username": "sa",
        "password": "",
    }

    @classmethod
    def _load_settings_defaults(cls) -> dict:
        """
        从 settings.yaml 读取默认数据库连接信息作为备用
        """
        try:
            if os.path.exists(cls.SETTINGS_FILE):
                with open(cls.SETTINGS_FILE, "r", encoding="utf-8") as f:
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
        os.makedirs(cls.CONFIG_DIR, exist_ok=True)

    @classmethod
    def load(cls) -> dict:
        """
        加载上次保存的连接配置
        如果密码为空，尝试从 settings.yaml 读取默认密码
        :return: 配置字典，包含 server, username, password
        """
        cls.ensure_config_dir()

        # 1. 基础默认值
        config = dict(cls.DEFAULT_CONFIG)

        # 2. 尝试从 settings.yaml 加载默认值
        settings_defaults = cls._load_settings_defaults()
        config.update(settings_defaults)

        # 3. 尝试读取 connection_profile.json（优先级最高）
        if os.path.exists(cls.CONFIG_FILE):
            try:
                with open(cls.CONFIG_FILE, "r", encoding="utf-8") as f:
                    profile = json.load(f)
                # 合并：profile 中的非空值覆盖默认值
                for key in ("server", "username", "password"):
                    if profile.get(key):
                        config[key] = profile[key]
            except (json.JSONDecodeError, IOError):
                pass

        return config

    @classmethod
    def save(cls, server: str, username: str, password: str):
        """
        保存连接配置到本地
        :param server: 服务器地址
        :param username: 登录用户名
        :param password: 登录密码
        """
        cls.ensure_config_dir()
        data = {
            "server": server,
            "username": username,
            "password": password,
        }
        try:
            with open(cls.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except IOError:
            pass  # 静默失败，不影响主流程

    @classmethod
    def clear(cls):
        """清除保存的配置"""
        cls.ensure_config_dir()
        try:
            if os.path.exists(cls.CONFIG_FILE):
                os.remove(cls.CONFIG_FILE)
        except IOError:
            pass
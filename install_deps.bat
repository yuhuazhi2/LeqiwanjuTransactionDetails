@echo off
chcp 65001 >nul
title 分店财务报表工具 - 依赖安装脚本
echo ============================================
echo  分店财务报表生成工具 - 依赖安装
echo ============================================
echo.

:: 检测 Python 环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [信息] Python 版本：
python --version
echo.

:: 检测是否在虚拟环境中
if "%VIRTUAL_ENV%"=="" (
    echo [信息] 未激活虚拟环境，将使用系统全局 Python
    echo.
    choice /C YN /M "是否创建并激活虚拟环境 .venv？"
    if errorlevel 1 (
        echo.
        echo 正在创建虚拟环境...
        python -m venv .venv
        call .venv\Scripts\activate.bat
        echo [OK] 虚拟环境已创建并激活
    )
) else (
    echo [信息] 当前虚拟环境: %VIRTUAL_ENV%
)

echo.
echo ============================================
echo  正在安装依赖包...
echo ============================================
echo.

:: 镜像源配置（国内镜像加速）
set MIRROR_URL=https://pypi.tuna.tsinghua.edu.cn/simple
set EXTRA_INDEX=https://mirrors.aliyun.com/pypi/simple

echo [配置] 主镜像源: %MIRROR_URL%
echo [配置] 备用源: %EXTRA_INDEX%
echo.

:: 升级 pip
echo [步骤 1/3] 升级 pip...
python -m pip install --upgrade pip -i %MIRROR_URL% --trusted-host pypi.tuna.tsinghua.edu.cn --trusted-host mirrors.aliyun.com
echo.

:: 安装依赖
echo [步骤 2/3] 安装依赖（openpyxl, pymssql, pyodbc, pyyaml）...
pip install openpyxl pymssql pyodbc pyyaml ^
    -i %MIRROR_URL% ^
    --trusted-host pypi.tuna.tsinghua.edu.cn ^
    --trusted-host mirrors.aliyun.com
echo.

:: 验证安装
echo [步骤 3/3] 验证依赖安装...
python -c "import openpyxl; print('openpyxl', openpyxl.__version__)" 2>nul || echo [警告] openpyxl 安装失败
python -c "import pymssql; print('pymssql', pymssql.__version__)" 2>nul || echo [警告] pymssql 安装失败
python -c "import pyodbc; print('pyodbc', pyodbc.version)" 2>nul || echo [警告] pyodbc 安装失败
python -c "import yaml; print('PyYAML', yaml.__version__)" 2>nul || echo [警告] PyYAML 安装失败

echo.
echo ============================================
echo  依赖安装完成！
echo ============================================
echo.
echo 使用说明:
echo   1. 编辑 config/settings.yaml 配置数据库连接信息
echo   2. 运行: python main.py
echo   3. 或在资源管理器中双击 main.py
echo.
echo 如果数据库连接失败，请确保:
echo   - SQL Server 2008R2 已启用 TCP/IP 协议
echo   - 防火墙允许 1433 端口通信
echo   - sa 账号密码配置正确
echo.

pausepause

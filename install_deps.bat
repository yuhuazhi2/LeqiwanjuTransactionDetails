@echo off
chcp 65001 >nul
title 分店财务报表工具 - 依赖安装脚本
echo ============================================
echo  分店财务报表生成工具 - 依赖安装
echo ============================================
echo.

:: 检测Python环境
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

:: ============================================================
:: 镜像源配置（优先使用国内镜像加速下载）
:: 清华大学 TUNA 源和阿里云源二选一
:: 如需切换源，请修改下面的 URL 即可
:: ============================================================
set MIRROR_URL=https://pypi.tuna.tsinghua.edu.cn/simple
:: set MIRROR_URL=https://mirrors.aliyun.com/pypi/simple

set EXTRA_INDEX=https://mirrors.aliyun.com/pypi/simple

echo [配置] 主镜像源: %MIRROR_URL%
echo [配置] 备用源: %EXTRA_INDEX%
echo.

:: 升级 pip
echo [步骤 1/4] 升级 pip...
python -m pip install --upgrade pip -i %MIRROR_URL% --trusted-host pypi.tuna.tsinghua.edu.cn --trusted-host mirrors.aliyun.com
echo.

:: 安装核心依赖
echo [步骤 2/4] 安装核心依赖（openpyxl, pymssql, pyyaml）...
pip install openpyxl pymssql pyyaml ^
    -i %MIRROR_URL% ^
    --trusted-host pypi.tuna.tsinghua.edu.cn ^
    --trusted-host mirrors.aliyun.com
echo.

:: 安装可选依赖（用于Web报表输出）
echo [步骤 3/4] 安装可选依赖（用于HTML报表输出）...
pip install jinja2 markupsafe ^
    -i %MIRROR_URL% ^
    --trusted-host pypi.tuna.tsinghua.edu.cn ^
    --trusted-host mirrors.aliyun.com
echo.

:: 验证安装
echo [步骤 4/4] 验证依赖安装...
python -c "import openpyxl; print('openpyxl', openpyxl.__version__)" 2>nul || echo [警告] openpyxl 安装失败
python -c "import pymssql; print('pymssql', pymssql.__version__)" 2>nul || echo [警告] pymssql 安装失败（如需连接SQL Server必须安装）
python -c "import yaml; print('PyYAML', yaml.__version__)" 2>nul || echo [警告] pyyaml 安装失败
python -c "import jinja2; print('jinja2', jinja2.__version__)" 2>nul || echo [信息] jinja2 未安装（不影响基本功能）

echo.
echo ============================================
echo  依赖安装完成！
echo ============================================
echo.
echo 使用说明:
echo   1. 编辑 config/settings.yaml 配置数据库连接
echo   2. 运行: python main.py
echo   3. 或在资源管理器中双击 main.py
echo.
echo 如果数据库连接失败，请确保:
echo   - SQL Server 2008R2 已启用 TCP/IP 协议
echo   - 防火墙允许 1433 端口通信
echo   - sa 账号密码配置正确
echo.

pause
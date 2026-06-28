# 分店财务报表生成工具

## 项目概述

本工具用于从**用友财务软件 T3 标准版 11.2** 的 SQL Server 2008R2 数据库中提取财务数据，按模板格式自动生成**多门店多页签**的财务报表 Excel 文件。

适用于连锁企业/多门店财务管理场景，支持同时核算**10+ 个门店**（账套），每个门店一个独立页签。

---

## 系统架构

```
LeqiwanjuTransactionDetails/
├── config/                        # 配置文件目录
│   └── settings.yaml              # 全局配置文件（数据库、模板、输出路径）
├── src/                           # 源代码
│   ├── database/                  # 数据库模块
│   │   ├── connector.py           # SQL Server 连接器（pymssql封装）
│   │   ├── ufsystem.py            # UFSystem 账套查询
│   │   └── t3data.py              # T3 财务数据提取（科目余额、发生额）
│   ├── template/                  # 模板模块
│   │   └── parser.py              # 模板解析器（读取.xlsx行列结构）
│   ├── report/                    # 报表生成模块
│   │   └── builder.py             # 报表生成器（核心：取数+填充+输出）
│   ├── utils/                     # 工具模块
│   │   └── logger.py              # 日志配置
│   └── __init__.py
├── output/                        # 报表输出目录
├── logs/                          # 日志目录
├── 分店财务报表模板.xlsx           # Excel模板文件
├── main.py                        # 主程序入口
├── install_deps.bat               # 依赖安装脚本（双击运行）
└── README.md                      # 本文件
```

---

## 运行环境要求

| 组件 | 要求 |
|------|------|
| **操作系统** | Windows 10（服务器/客户端均可） |
| **Python** | 3.8+（推荐 3.11） |
| **SQL Server** | 2008R2（用友T3标准版配套） |

### 支持组件（已集成到 install_deps.bat）

| 包名 | 用途 | 安装源 |
|------|------|--------|
| `pymssql` | SQL Server 数据库连接 | 清华 TUNA / 阿里云 |
| `openpyxl` | Excel 文件读写 | 清华 TUNA / 阿里云 |
| `PyYAML` | 配置文件解析 | 清华 TUNA / 阿里云 |
| `jinja2` | （可选）HTML报表输出 | 清华 TUNA / 阿里云 |

---

## 快速开始

### 第一步：安装依赖

**方式一（推荐）**：双击运行 `install_deps.bat`

**方式二**：命令行执行
```cmd
cd G:\JetBrains\PycharmProjects\LeqiwanjuTransactionDetails
pip install openpyxl pymssql pyyaml -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 第二步：修改配置

编辑 [`config/settings.yaml`](config/settings.yaml)，配置数据库连接信息：

```yaml
database:
  server: "192.168.1.100"    # SQL Server 服务器 IP
  port: 1433                  # 端口
  username: "sa"              # 登录用户名
  password: "your_password"   # 登录密码
  ufsystem_db: "UFSystem"     # 公共账套数据库（默认）
  charset: "GBK"              # 用友数据库字符集（不可修改）

template:
  filepath: "分店财务报表模板.xlsx"   # 模板文件路径
  source_sheet: "sheet"             # 模板页签名

report_year: 2024                     # 报表年份
```

### 第三步：运行

**方式一（推荐）**：在 VSCode 终端或 CMD 中执行
```cmd
python main.py
```

**方式二**：在资源管理器中双击 `main.py`

**方式三**：命令行参数控制
```cmd
# 指定年份和月份
python main.py --year 2024 --months 1 2 3

# 仅生成特定门店
python main.py --accounts 001 003 005

# 自定义配置文件
python main.py -c my_settings.yaml
```

---

## 模板说明

模板文件 [`分店财务报表模板.xlsx`](分店财务报表模板.xlsx) 定义了报表的格式结构：

| 区域 | 行内容 | 说明 |
|------|--------|------|
| **收入区** | 业绩完成率、目标预算、销售业绩（5个渠道） | 自动从收入科目取数 |
| **成本区** | 主营业务成本 | 从成本科目取数 |
| **毛利区** | 毛利、毛利率 | **自动计算**（总收入-总成本） |
| **费用区** | 营业费用（10项） | 从费用科目取数 |
| | 管理费用（12项） | 从费用科目取数 |
| | 财务费用（1项） | 从费用科目取数 |
| **利润区** | 利润、利润率 | **自动计算**（毛利-费用） |
| **结转区** | 分红、银行余额 | 预留手工填写区域 |

### 自定义模板
- 如需修改报表格式，直接编辑模板文件
- **保持 A 列行标签和 Header 行结构不变**
- 增加/删减科目行会被自动识别

---

## 数据库取数逻辑

本工具从用友 T3 标准版的以下数据库中取数：

### 1. UFSystem（公共数据库）
- **表**: `UA_Account`
- **作用**: 获取所有账套列表（账套号 → 门店名称映射）

### 2. UFDATA_XXX_YYYY（各账套数据库）
- **表**: `GL_AccSum`（科目汇总表）
- **取数字段**: `md`(借方发生额), `mc`(贷方发生额)
- **科目对照**:

| 报表行标签 | T3 科目编码 | 说明 |
|------------|------------|------|
| 油菜花收入 | 600101 | 收入类-贷方 |
| 现金收入 | 600102 | 收入类-贷方 |
| 美团收入 | 600103 | 收入类-贷方 |
| 抖音收入 | 600104 | 收入类-贷方 |
| 其他业务收入 | 6051 | 收入类-贷方 |
| 主营业务成本 | 6401 | 成本类-借方 |
| 广告费/物料费/... | 6601xx | 营业费用-借方 |
| 工资/办公费/... | 6602xx | 管理费用-借方 |
| 手续费 | 660301 | 财务费用-借方 |

> **提示**：如果实际使用的科目编码与默认值不同，请在 [`src/database/t3data.py`](src/database/t3data.py) 中修改 `REVENUE_CODES` / `EXPENSE_CODES` 映射。

---

## 命令行参数详解

```
python main.py [选项]

选项:
  -h, --help                 显示帮助信息
  -c, --config FILE          指定配置文件路径
  --year YEAR                指定报表年份
  --months M [M ...]         指定月份范围（如 1 2 3）
  --accounts ID [ID ...]     指定账套号（如 001 003）
  --no-open                  生成后不自动打开文件
  --output-dir DIR           指定输出目录
  -v, --verbose              输出详细调试日志
```

---

## 常见问题排查

### Q1: 数据库连接失败
- 确保 SQL Server 2008R2 **已启用 TCP/IP 协议**
- 检查 Windows 防火墙是否允许 **1433 端口**
- 在 SQL Server Management Studio 中确认 **sa 账号已启用**

### Q2: 提示 "pymssql 安装失败"
- 安装 Microsoft Visual C++ 14.0 运行库
- 或从清华源手动安装：`pip install pymssql -i https://pypi.tuna.tsinghua.edu.cn/simple`

### Q3: 生成的报表没有数据
- 检查 [`config/settings.yaml`](config/settings.yaml) 中的 `report_year` 设置
- 确认对应账套数据库中存在该年度的凭证数据
- 查看 `logs/app.log` 日志文件获取详细错误信息

### Q4: 科目编码不匹配
- 在 [`src/database/t3data.py`](src/database/t3data.py) 中修改相应科目编码映射
- 或在配置文件中自定义映射（后续版本支持外部科目映射表）

---

## 后续扩展方向

- [ ] **HTML/Web 报表输出**：支持直接在浏览器中查看
- [ ] **交互式配置界面**：GUI 配置数据库和模板设置
- [ ] **年度自动检测**：自动识别账套数据库中的最大年度
- [ ] **自定义科目映射**：通过外部配置文件管理科目编码映射
- [ ] **批量历史数据**：支持一次生成多年份报表
- [ ] **自动发送邮件**：生成后自动发送给指定收件人
- [ ] **数据库连接池**：支持高并发取数场景
- [ ] **数据校验**：对提取的财务数据进行勾稽校验

---

## 许可证

仅供内部使用。
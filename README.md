# 分店财务报表生成工具（乐其玩聚）

## 项目概述

本工具用于从**用友财务软件 T3 标准版 11.2** 的 SQL Server 2008R2 数据库中提取财务数据，按模板格式自动生成**多门店多页签**的财务报表 Excel 和 HTML 文件。

适用于连锁企业/多门店财务管理场景，支持同时核算**10+ 个门店**（账套），每个门店一个独立页签。

---

## 系统架构

```
LeqiwanjuTransactionDetails/
├── config/                            # 配置文件目录
│   ├── settings.yaml.example          # 全局配置示例（复制为 settings.yaml 使用）
│   └── ...                            # 实际配置（.gitignore 已排除）
├── src/                               # 源代码
│   ├── database/                      # 数据库模块
│   │   ├── connector.py               # SQL Server 连接器（pyodbc/pymssql）
│   │   ├── ufsystem.py                # UFSystem 账套查询（UA_Account）
│   │   └── t3data.py                  # T3 财务数据提取
│   ├── template/                      # 模板模块
│   │   └── parser.py                  # 模板解析器（读取 .xlsx 行列结构）
│   ├── report/                        # 报表生成模块
│   │   ├── builder.py                 # 报表生成器（核心：取数+填充+XLSX输出）
│   │   └── html_renderer.py           # HTML 报表渲染器
│   ├── utils/                         # 工具模块
│   │   ├── logger.py                  # 日志配置
│   │   └── browser_utils.py           # 浏览器打开工具
│   ├── gui/                           # GUI 桌面应用界面
│   │   ├── app.py                     # 主窗口（Tkinter）
│   │   └── config_manager.py          # 配置管理
│   └── __init__.py
├── output/                            # 报表输出目录（.gitignore 已排除）
├── logs/                              # 日志目录（.gitignore 已排除）
├── 分店财务报表模板.xlsx               # Excel 模板文件
├── main.py                            # 主程序入口
├── install_deps.bat                   # 依赖安装脚本（双击运行）
├── README.md                          # 本文件
├── .gitignore
└── .gitattributes
```

---

## 运行环境要求

| 组件 | 要求 |
|------|------|
| **操作系统** | Windows 10（服务器/客户端均可） |
| **Python** | 3.8+（推荐 3.11） |
| **SQL Server** | 2008R2（用友 T3 标准版配套） |

### 依赖包

| 包名 | 用途 | 安装源 |
|------|------|--------|
| `pymssql` | SQL Server 数据库连接 | 清华 TUNA / 阿里云 |
| `pyodbc` | SQL Server 数据库连接（支持 Unicode 参数） | 清华 TUNA / 阿里云 |
| `openpyxl` | Excel 文件读写 | 清华 TUNA / 阿里云 |
| `PyYAML` | 配置文件解析 | 清华 TUNA / 阿里云 |

---

## 快速开始

### 第一步：安装依赖

**方式一（推荐）**：双击运行 `install_deps.bat`

**方式二**：命令行执行
```cmd
cd LeqiwanjuTransactionDetails
pip install openpyxl pymssql pyodbc pyyaml -i https://pypi.tuna.tsinghua.edu.cn/simple
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

**方式二**：双击 `main.py`

**方式三**：命令行参数控制
```cmd
# 指定年份和月份
python main.py --year 2024 --months 1 2 3

# 仅生成特定门店
python main.py --accounts 001 003 005

# 指定输出格式（xlsx / html / both）
python main.py --format both

# 自定义配置文件
python main.py -c my_settings.yaml
```

---

## 模板说明

模板文件 [`分店财务报表模板.xlsx`](分店财务报表模板.xlsx) 定义了报表的行列结构：

### 行结构（A 列标签）

| 行号 | 行标签 | 行类型 | 数据来源/计算规则 |
|------|--------|--------|-------------------|
| 1 | 店名 | — | 标题行，自动填入账套名称 |
| 2 | — | header | 月份表头（1月、2月、...、年度合计） |
| 3 | **业绩完成率** | `performance` | 公式：=目标预算/销售业绩 |
| 4 | **目标预算（一档）** | `target` | 手工填列 |
| 5 | **销售业绩** | `sales` | 自动合计：区间内各渠道之和 |
| 6~N | 各收入渠道 | `revenue_detail` | 从 5101/5102 科目期间损益结转取数 |
| N+1 | **主营业务成本** | `cost` | 从 5401 科目期间损益结转取数 |
| N+2 | **毛利** | `gross_profit` | 自动计算：销售业绩 - 主营业务成本 |
| N+3 | **毛利率** | `gross_rate` | 自动计算：毛利 / 销售业绩 × 100% |
| N+4 | **费用支出** | `expense_total` | 自动计算：营业费用+管理费用+财务费用 |
| N+5 | **费用率** | `expense_rate` | 自动计算：费用支出 / 销售业绩 × 100% |
| N+6 | **营业费用** | `expense_header` | 自动合计：区间内明细行之和 |
| N+7~M | 营业费用明细（5501） | `expense_detail` | 从 5501 科目期间损益结转取数 |
| M+1 | **管理费用** | `manage_header` | 自动合计：区间内明细行之和 |
| M+2~K | 管理费用明细（5502） | `manage_detail` | 从 5502 科目期间损益结转取数 |
| K+1 | **财务费用** | `finance_header` | 自动合计：区间内明细行之和 |
| K+2 | 财务费用明细（5503） | `finance_detail` | 从 5503 科目期间损益结转取数 |
| K+3 | **利润** | `profit` | 自动计算：销售业绩-成本-费用支出 |
| K+4 | **利润率** | `profit_rate` | 自动计算：利润 / 销售业绩 × 100% |
| K+5 | **分红** | `dividend` | 手工填列 |
| K+6 | **总分红（往年分红万）** | `total_dividend` | 手工填列 |
| K+7 | **投资** | `invest` | 手工填列 |
| K+8 | **银行余额** | `bank_balance` | 从 GL_AccSum 1002 科目 me 字段取数 |

> **注意**：行号 N、M、K 是动态的，取决于各账套数据库中 5101/5501/5502 科目的明细数量。

### 列结构

- **月份列**：根据已结账月份动态生成（从 GL_Mend 表查询）
- **年度合计列**：固定最后一列，自动计算各月份之和（利润/利润率/银行余额除外）

### 自定义模板

- 如需修改报表格式，直接编辑模板文件
- **保持 A 列行标签和 Header 行结构不变**
- 增加/删减科目行会被自动识别

---

## 核心模块说明

### 1. 数据库模块 [`src/database/`](src/database/)

#### `connector.py` — 数据库连接器
- 封装 `pymssql` 和 `pyodbc` 两种连接方式
- `execute_query()` — 使用 pymssql 执行 SQL 查询（返回字典列表）
- `execute_query_odbc()` — 使用 pyodbc 执行 SQL 查询（支持 Unicode 参数，主要用于中文条件查询）
- `get_closed_periods()` — 查询已结账月份列表（从 GL_Mend 表）

#### `ufsystem.py` — 账套查询
- `UFSystemQuerier` 类
- 从 `UFSystem.dbo.UA_Account` 表查询所有账套
- `get_account_by_id()` — 按账套号查询
- `get_filtered_accounts()` — 批量查询（可排除 998/999 等非业务账套）
- `AccountInfo` 数据类：包含 `cAcc_Id`（账套号）、`cAcc_Name`（账套名称）、`sheet_name`（页签名）

#### `t3data.py` — 财务数据提取
- `T3DataExtractor` 类
- **科目余额查询**：
  - `get_subject_balance()` — 单科目单月余额
  - `get_subject_balances()` — 科目族全年各月数据（按科目前缀模糊匹配）
- **Code 表动态查询**（用于确定行结构）：
  - `get_revenue_subjects_from_code()` — 查询 5101/5102 收入科目
  - `get_expense_subjects_from_code()` — 查询 5501 营业费用科目
  - `get_manage_subjects_from_code()` — 查询 5502 管理费用科目
- **期间损益结转查询**：
  - `get_period_transfer_vouchers()` — 从 GL_AccVouch 表查询 cdigest='期间损益结转' 的凭证
  - `get_code_subject_name_batch()` — 批量查询科目名称
- **银行余额查询**：
  - `get_bank_balance_from_gl_accsum()` — 从 GL_AccSum 查询 1002 科目 me 字段

### 2. 模板模块 [`src/template/`](src/template/)

#### `parser.py` — 模板解析器
- `TemplateParser` 类
- 解析 Excel 模板的 A 列（行标签）和第一行（列标题）
- 返回 `TemplateLayout` 数据类：包含 `row_labels`、`columns`、`months` 等
- `RowLabel` 数据类：`row_index`、`label`、`level`、`skip`
- `ColumnHeader` 数据类：`col_index`、`label`、`is_total`、`month`

### 3. 报表生成模块 [`src/report/`](src/report/)

#### `builder.py` — Excel 报表生成器（核心）
- `ReportBuilder` 类
- **两个生成入口**：
  - `build()` — 旧版完整流程（已逐步被 `build_framework` 取代）
  - `build_framework()` — 新版主入口（V2.0+），包含全部步骤
- **关键步骤（build_framework 流程）**：

| 步骤 | 方法 | 功能 | 版本 |
|------|------|------|------|
| 1 | `_get_accounts()` | 查询合规账套列表 | V2.0 |
| 2 | `load_workbook()` | 加载模板工作簿 | V2.0 |
| 3 | `copy_worksheet()` | 为每个账套复制页签 | V2.0 |
| 4 | — | 删除原始模板页签 | V2.0 |
| 5 | — | A2 填入账套名称+年度（红色字体） | V2.0 |
| 5.5 | `_adjust_sheet_columns()` | 根据已结账月份调整列结构 | V2.1 |
| 5.6 | `_adjust_sheet_subject_rows()` | 根据 5101/5102 科目调整收入明细行 | V2.2 |
| 5.7 | `_adjust_sheet_expense_rows()` | 根据 5501 科目调整营业费用明细行 | V2.3 |
| 5.8 | `_adjust_sheet_manage_rows()` | 根据 5502 科目调整管理费用明细行 | V2.4 |
| 5.9 | `_apply_row_colors()` | 设置行背景色（分级着色） | V2.5 |
| 5.10 | `_fill_period_transfer_data()` | 从 GL_AccVouch 填列期间损益结转数据 | V2.6 |
| 5.11 | `_calculate_summary_sums()` | 计算汇总行合计和利润/利润率 | V2.8-V2.9 |
| 5.12 | `_fill_bank_balance()` | 从 GL_AccSum 填列银行余额 | V3.0 |
| 5.13 | `_fix_performance_formula()` | 修复业绩完成率行 Excel 公式 | V2.9 |
| 6 | `_widen_total_column()` | 调整合计列宽度 | V2.7 |
| 7 | — | 根据 output_format 输出文件 | V2.0+ |

- **行背景色方案（_apply_row_colors）**：

| 行标签 | 颜色 | 填充常量 |
|--------|------|----------|
| 业绩完成率 | 浅紫色 #E8D5F5 | `FILL_PERFORMANCE` |
| 目标预算（一档） | 浅绿色 #D5F5E3 | `FILL_TARGET` |
| 销售业绩 | 浅蓝色 #D6EAF8 | `FILL_SALES` |
| 主营业务成本 | 浅橙色 #FDEBD0 | `FILL_COST` |
| 毛利率/费用率 | 浅灰色 #F2F3F4 | `FILL_RATE` |
| 营业/管理/财务费用标题行 | 深黄色 #F9E79F | `FILL_EXPENSE_HEADER` |
| 费用明细行 | 浅黄色 #FCF3CF | `FILL_EXPENSE` |
| 利润/利润率 | 中浅蓝 #AED6F1 | `FILL_PROFIT` |
| 分红/投资 | 浅粉色 #F5B7B1 | `FILL_DIVIDEND` |
| 银行余额 | 淡紫色 #D2B4DE | `FILL_BANK` |

#### `html_renderer.py` — HTML 报表渲染器
- `HtmlRenderer` 类
- 与 Excel 流程**共用同一套数据提取器**（T3DataExtractor），不重复查询数据库
- 严格复刻 Excel 的行顺序、行标签、背景色方案
- 输出单一 HTML 文件，内嵌 CSS+JavaScript，不依赖外部资源
- 多账套通过 JavaScript Tab 切换展示
- **颜色常量**与 ReportBuilder 的 `FILL_*` 严格对应
- **行标签映射** `LABEL_COLOR_MAP` 控制每行的背景色
- **计算逻辑**与 Excel 版本完全一致（_calculate_summary_sums）
- 支持数字格式化（千分位、负数标红、零值灰显）

---

## 数据计算规则

### 汇总行合计

| 汇总行 | 区间 | 取数规则 |
|--------|------|----------|
| **销售业绩** | 销售业绩行 ↔ 主营业务成本行 | 取区间内**所有行**各列数值合计 |
| **营业费用** | 营业费用行 ↔ 管理费用行 | 仅取**有背景色**（非空白）的行各列数值合计 |
| **管理费用** | 管理费用行 ↔ 财务费用行 | 仅取**有背景色**（非空白）的行各列数值合计 |
| **财务费用** | 财务费用行 ↔ 利润行 | 仅取**有背景色**（非空白）的行各列数值合计 |

### 费用支出与费用率

- **费用支出**（V3.0 新增）：`费用支出 = 营业费用 + 管理费用 + 财务费用`（逐列计算）
- **费用率**（V3.0 新增）：`费用率 = 费用支出 / 销售业绩 × 100%`（逐列计算）

### 利润与利润率

- **利润**（逐列计算）：`利润 = 销售业绩 - 主营业务成本 - 费用支出`
- **利润率**（逐列计算）：`利润率 = 利润 / 销售业绩 × 100%`
- 覆盖所有月份列和年度合计列

### 年度合计列

- 明细行合计列 = 本行各月份数值之和
- 利润行合计列 = 按利润公式计算（非月份和）
- 利润率行合计列 = 按利润率公式计算（非月份和）
- **银行余额行不计算合计**（时点数不计年度合计）

---

## 输出格式

通过 `settings.yaml` 中的 `output.format` 或命令行 `--format` 参数配置：

| 格式 | 配置值 | 说明 |
|------|--------|------|
| **XLSX** | `xlsx` | Excel 文件，基于模板复刻行列结构 |
| **HTML** | `html` | 自包含 HTML 文件，可在线预览和打印 |
| **同时输出** | `both` | 同时生成 XLSX 和 HTML 两种格式 |

### HTML 报表特性

- 严格复刻 Excel 的行背景色、列结构、标签顺序
- 多账套通过 JavaScript Tab 页签切换
- 所有样式和脚本内嵌于单个 HTML 文件，无需网络依赖
- 支持数字格式化（千分位、负数标红、零值灰显）
- 行 CSS 类名自动映射：`row-{row_type.replace("_", "-")}`

---

## 数据库取数逻辑

本工具从用友 T3 标准版的以下数据库中取数：

### 1. UFSystem（公共数据库）
- **表**: `UA_Account`
- **作用**: 获取所有账套列表（账套号 → 门店名称映射）

### 2. UFDATA_XXX_YYYY（各账套数据库）
- **表**: `GL_AccSum`（科目汇总表）— 用于旧版 `build()` 流程和银行余额查询
- **表**: `GL_AccVouch`（凭证表）— 用于 `build_framework()` 流程，取期间损益结转数据
- **表**: `Code`（科目字典）— 查询科目编码与名称对应关系，动态确定行结构
- **表**: `GL_Mend`（结账表）— 查询已结账月份，确定列结构

### 数据方向

| 报表行标签 | T3 科目前缀 | 数据方向 | 数据源表 |
|------------|-------------|----------|----------|
| 各收入渠道 | 5101 / 5102 | 期间损益结转取 md（借方） | GL_AccVouch |
| 主营业务成本 | 5401 | 期间损益结转取 mc（贷方） | GL_AccVouch |
| 各营业费用 | 5501 | 期间损益结转取 mc（贷方） | GL_AccVouch |
| 各管理费用 | 5502 | 期间损益结转取 mc（贷方） | GL_AccVouch |
| 财务费用 | 5503 | 期间损益结转取 mc（贷方） | GL_AccVouch |
| 银行余额 | 1002 | 取 me（期末余额） | GL_AccSum |

> **提示**：所有科目均通过 Code 表动态查询，不再硬编码。如果实际使用的科目编码与默认值不同，查询逻辑会自动适应。

---

## 命令行参数详解

```
python main.py [选项]

选项:
  -h, --help                 显示帮助信息
  --cli                      使用命令行模式运行（默认启动 GUI 模式）
  -c, --config FILE          指定配置文件路径
  --year YEAR                指定报表年份
  --months M [M ...]         指定月份范围（如 1 2 3）
  --accounts ID [ID ...]     指定账套号（如 001 003）
  --format FORMAT            输出格式：xlsx / html / both
  --no-open                  生成后不自动打开文件
  --output-dir DIR           指定输出目录
  -v, --verbose              输出详细调试日志
```

---

## 配置说明

配置文件 [`config/settings.yaml`](config/settings.yaml) 包含以下主要配置项：

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `database.server` | SQL Server 主机地址 | `server03` |
| `database.port` | 端口号 | `1433` |
| `database.username` | 数据库登录名 | `sa` |
| `database.password` | 数据库密码 | — |
| `database.charset` | 用友数据库字符集 | `GBK` |
| `template.filepath` | 模板文件路径 | `分店财务报表模板.xlsx` |
| `template.source_sheet` | 模板页签名 | `sheet` |
| `output.format` | 输出格式 | `xlsx` |
| `output.dir` | 输出目录 | `output` |
| `output.open_when_done` | 生成后自动打开 | `true` |
| `account_filter.include_ids` | 指定账套号列表 | `[]`（取全部） |
| `account_filter.exclude_ids` | 排除账套号列表 | `[]` |
| `report_year` | 报表年份 | `2024` |
| `report_months` | 月份范围 | `[]`（取全年） |
| `logging.level` | 日志级别 | `INFO` |

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
- 本工具通过 Code 表动态查询 5101/5102/5501/5502/5503 科目体系
- 如果账套使用了不同的科目前缀，请在 [`src/database/t3data.py`](src/database/t3data.py) 中修改相关查询逻辑

### Q5: 部分账套收入数据为零
- 检查该账套数据库的 `Code` 表中 5101/5102 科目编码是否存在
- 确认该账套 `GL_AccVouch` 表中存在该年度的期间损益结转凭证
- 查看日志中该账套的"期间损益结转数据填列"记录数量

### Q6: HTML 报表缺少费用支出/费用率行
- 该问题已在 V3.0 修复
- 确认使用最新版本的 `src/report/html_renderer.py`
- 检查 `_build_row_labels()` 方法中是否包含 `expense_total` 和 `expense_rate` 行类型

---

## 版本历史

| 日期 | 版本 | 变更说明 |
|------|------|----------|
| 2026-07-10 | V3.0 | 新增费用支出/费用率行、HTML 同步修复、银行余额填列 |
| 2026-07-09 | V2.9 | 新增业绩完成率公式、利润/利润率计算优化 |
| 2026-07-08 | V2.8 | 新增汇总行合计（销售业绩/营业费用/管理费用） |
| 2026-07-07 | V2.7 | 合计列宽度自动调整 |
| 2026-07-06 | V2.6 | 新增期间损益结转数据填列 |
| 2026-07-05 | V2.5 | 新增行背景色（分级着色） |
| 2026-07-04 | V2.4 | 管理费用明细行动态调整 |
| 2026-07-03 | V2.3 | 营业费用明细行动态调整 |
| 2026-07-02 | V2.2 | 收入明细行动态调整（5101/5102 科目） |
| 2026-07-01 | V2.1 | 列结构调整（根据已结账月份） |
| 2026-06-30 | V2.0 | 全新 build_framework 流程框架雏形 |
| 2026-06-25 | V1.0 | 初始版本，基于 build() 流程生成完整报表 |

---

## 后续扩展方向

- [ ] **交互式配置界面**：GUI 配置数据库和模板设置（已有基础框架 [`src/gui/`](src/gui/)）
- [ ] **年度自动检测**：自动识别账套数据库中的最大年度
- [ ] **自定义科目映射**：通过外部配置文件管理科目编码映射
- [ ] **批量历史数据**：支持一次生成多年份报表
- [ ] **自动发送邮件**：生成后自动发送给指定收件人
- [ ] **数据库连接池**：支持高并发取数场景
- [ ] **数据校验**：对提取的财务数据进行勾稽校验（借贷平衡、科目完整性）
- [ ] **同比环比分析**：支持与历史同期数据的对比分析

---

## 许可证

仅供内部使用。
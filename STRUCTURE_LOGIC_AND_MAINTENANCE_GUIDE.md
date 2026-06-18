# 项目结构、运行逻辑与后期维护指南

本文档用于说明 `Invoice Ledger Automator` 的项目结构、核心运行流程、字段识别规则、Excel 写入逻辑、文件归档逻辑和后期维护方法。它比 `README.md` 更偏技术拆解，目标是让后续维护者能快速判断“问题在哪里、该改哪个函数、改完如何验证”。

建议阅读方式：

- 如果只是使用项目，先读 `README.md`。
- 如果要修改识别规则、排查失败发票、调整台账逻辑，读本文档。
- 如果要开源或提交代码，重点看“隐私与开源保护”和“维护检查清单”。

## 1. 项目目标

本项目用于自动处理可复制文字的 PDF 发票。脚本会从 PDF 中提取发票字段，追加写入 Excel 台账，并根据识别结果把 PDF 移动到不同文件夹。

系统解决的核心问题是：

- 批量读取 PDF 发票。
- 自动提取台账所需字段。
- 将结果追加写入 Excel。
- 新数据不覆盖旧数据。
- 成功和失败文件自动分流。
- 重复发票号自动去重。
- 没有发票号码的失败记录自动清理。

当前版本不包含 OCR，因此它只适用于文字可以被复制的 PDF。如果 PDF 是扫描件或图片版，需要额外接入 OCR 模块。

## 2. 根目录结构

项目根目录主要包含以下文件和文件夹：

```text
invoice_register.py
README.md
STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md
environment.yml
.gitignore
invoices_to_process/
archived_invoices/
failed_invoices/
invoice_ledger.xlsx
.conda/
```

各部分作用如下。

```text
invoice_register.py
```

主程序文件，包含 PDF 文本读取、字段解析、Excel 写入、文件归档、重复记录处理和失败记录清理等全部核心逻辑。

```text
README.md
```

项目对外说明文档，适合 GitHub 展示。内容包括项目简介、安装方法、使用方法、隐私说明等。

```text
STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md
```

当前文档，用于细致说明项目结构、运行逻辑和维护方法，适合排查问题、优化规则和二次开发时阅读。

```text
environment.yml
```

Conda 环境配置文件。用于在新电脑或新环境中重建项目运行环境。

```text
.gitignore
```

Git 忽略规则。它会忽略发票 PDF、Excel 台账、本地 Conda 环境、缓存文件和处理目录中的真实业务文件，降低误提交敏感数据的风险。

```text
invoices_to_process/
```

待识别发票目录。用户把需要处理的 PDF 发票放到这个文件夹中，脚本运行时只扫描这个目录下的 PDF。

```text
archived_invoices/
```

已归档发票目录。识别成功的 PDF 会被移动到这里。

```text
failed_invoices/
```

识别失败发票目录。只要必填字段识别失败，或在备注区域发现可选关键词但无法提取对应值，PDF 就会被移动到这里。

```text
invoice_ledger.xlsx
```

自动生成和持续追加的 Excel 台账。该文件属于业务数据，已被 `.gitignore` 忽略，不应提交到公开仓库。

```text
.conda/
```

项目本地 Conda 环境目录。该目录只用于本机运行，不应提交到 GitHub。

## 3. 运行入口

脚本入口位于 `invoice_register.py` 最底部：

```python
if __name__ == "__main__":
    main()
```

用户通常通过以下命令运行：

```bash
./.conda/bin/python invoice_register.py
```

如果只想预览识别结果，不写入 Excel，也不移动 PDF，可以运行：

```bash
./.conda/bin/python invoice_register.py --dry-run
```

`main()` 函数负责解析命令行参数，然后调用：

```python
process_invoices(dry_run=args.dry_run)
```

## 4. 总体运行流程

一次完整运行可以拆成以下步骤：

```text
1. 确保工作目录存在
2. 扫描 invoices_to_process/ 中的 PDF
3. 逐个读取 PDF 文本
4. 提取必填字段
5. 提取备注区域中的可选字段
6. 判断识别成功或失败
7. 写入或更新 Excel 台账
8. 自动清理无发票号码的失败行
9. 按发票号码去重
10. 移动 PDF 到归档目录或失败目录
11. 打印处理结果
```

对应主流程函数是：

```python
process_invoices()
```

这个函数会先确保以下目录存在：

```python
INPUT_DIR
ARCHIVE_DIR
FAILED_DIR
```

然后扫描 `invoices_to_process/` 下的 `.pdf` 和 `.PDF` 文件。

如果没有待处理 PDF，脚本会输出提示并结束：

```text
未发现待识别PDF：...
处理完成：成功 0 个，失败 0 个。
```

## 5. 路径常量

脚本顶部定义了所有关键路径：

```python
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "invoices_to_process"
ARCHIVE_DIR = BASE_DIR / "archived_invoices"
FAILED_DIR = BASE_DIR / "failed_invoices"
LEDGER_PATH = BASE_DIR / "invoice_ledger.xlsx"
```

这样设计的好处是：

- 所有路径都相对于脚本所在目录。
- 用户不需要手动配置绝对路径。
- 项目移动到其他电脑后仍能运行。
- 目录名称集中管理，后续改名更容易。

## 6. 数据模型

脚本使用 `InvoiceRecord` 表示一张发票的识别结果：

```python
@dataclass
class InvoiceRecord:
    invoice_number: str = ""
    invoice_date: str = ""
    buyer_name: str = ""
    seller_name: str = ""
    total_amount: str = ""
    usd_amount: str = "无"
    exchange_rate: str = "无"
    bill_of_lading: str = "无"
    source_file: str = ""
    status: str = "成功"
    failed_reason: str = ""
    processed_at: str = ...
```

字段含义如下：

```text
invoice_number   发票号码
invoice_date     开票日期
buyer_name       购买方名称
seller_name      销售方名称
total_amount     价税合计（小写）
usd_amount       美金金额
exchange_rate    汇率
bill_of_lading   提单号
source_file      源 PDF 文件名
status           识别状态
failed_reason    失败原因
processed_at     处理时间
```

`as_row()` 方法会把对象转换成 Excel 可以直接写入的一行数据。

## 7. Excel 表头

Excel 台账表头由 `HEADERS` 定义：

```python
HEADERS = [
    "发票号码",
    "开票日期",
    "购买方名称",
    "销售方名称",
    "价税合计（小写）",
    "美金金额",
    "汇率",
    "提单号",
    "源文件名",
    "识别状态",
    "失败原因",
    "处理时间",
]
```

虽然项目目录和文件名已经改为英文，但 Excel 表头仍保留中文。这是因为这些字段直接对应中国发票内容和原始台账业务习惯，中文字段更利于日常使用。

## 8. PDF 文本读取

PDF 文本读取由以下函数完成：

```python
extract_pdf_text(pdf_path)
```

它使用 `pdfplumber` 打开 PDF，并逐页调用：

```python
page.extract_text()
```

最后把所有页面文本拼接成一个字符串。

如果 PDF 是可复制文字的电子发票，通常能成功提取。如果 PDF 是扫描件、图片版或被加密限制读取，可能会失败或提取为空。

## 9. 文本预处理

脚本使用：

```python
compact_text(text)
```

将多余空白、换行压缩成单行文本。这样做是为了让正则表达式能跨行匹配发票字段。

例如 PDF 抽取后可能出现：

```text
购 名称：某公司 销 名称：某公司
买 售
方 方
```

压缩成一行后，更容易识别购买方和销售方。

## 10. 必填字段识别

必填字段定义在：

```python
REQUIRED_FIELDS = ["发票号码", "开票日期", "购买方名称", "销售方名称", "价税合计（小写）"]
```

对应识别函数如下：

```text
发票号码          extract_invoice_number()
开票日期          extract_invoice_date()
购买方和销售方    extract_names()
价税合计（小写）  extract_total_amount()
```

如果任意必填字段没有识别到，系统会把对应失败原因写入 `failed_reason`。

例如：

```text
发票号码识别失败
开票日期识别失败
购买方名称识别失败
销售方名称识别失败
价税合计（小写）识别失败
```

只要存在失败原因，`status` 就会从默认的 `成功` 改成 `失败`。

## 11. 发票号码识别

发票号码由：

```python
extract_invoice_number(text, one_line)
```

负责识别。

它支持几类情况：

- 标准格式：`发票号码：数字`
- 字段中间有空格：`发 票 号码：数字`
- 老版发票中文本顺序被 PDF 打散，号码不紧跟字段标签
- 老版电子普通发票中的 `发 票监` 附近出现号码

这个函数会先尝试常规正则，再使用老版发票兜底规则。

## 12. 开票日期识别

开票日期由：

```python
extract_invoice_date(text, one_line)
```

负责识别。

它支持：

- `2025年05月20日`
- `2025-05-20`
- `2025/05/20`
- 老版发票中散落的 `2024 07 30`

识别后会通过：

```python
normalize_date()
```

统一转换成：

```text
YYYY-MM-DD
```

这样 Excel 中的日期格式更规整，也方便排序和筛选。

## 13. 购买方和销售方识别

购买方和销售方由：

```python
extract_names(text, one_line)
```

负责识别。

它支持多种版式：

- `购 名称：... 销 名称：...`
- `购买方名称：... 销售方名称：...`
- `名 称：... 名 称：...`
- 老版发票中公司名散落在不同位置

如果常规规则失败，会调用：

```python
extract_legacy_buyer()
extract_legacy_seller()
```

这两个函数用于处理老版电子普通发票中 PDF 文本顺序较乱的情况。

## 14. 价税合计识别

价税合计由：

```python
extract_total_amount(text, one_line)
```

负责识别。

它主要寻找：

```text
（小写）¥数字
```

也兼容：

```text
(小写) ¥数字
```

以及大写金额后跟小写金额的老版发票格式。

金额会通过：

```python
normalize_amount()
```

去掉逗号，保留标准数字字符串。

## 15. 备注区域识别

`美金金额`、`汇率`、`提单号` 不是在整张发票全文中识别，而是只在备注区域中识别。

备注区域由：

```python
extract_remark_text(text)
```

负责提取。

这个设计非常重要，因为发票其他区域可能出现和业务字段无关的关键词。例如销售方银行账号里可能出现 `USD`，但那只是账号币种标识，不代表美金金额。

备注区域提取逻辑大致是：

```text
1. 找到价税合计所在行
2. 从价税合计后面的内容开始收集
3. 遇到开票人、收款人等结尾字段时停止
4. 排除开户行、银行账号、账号、账户等行
5. 返回清理后的备注文本
```

如果找不到价税合计后的备注区域，脚本还会尝试寻找 `备注`、`备 注`、`备`、`注` 等标记。

## 16. 可选字段识别

可选字段包括：

```text
美金金额
汇率
提单号
```

识别由：

```python
extract_keyword_number()
```

负责。

### 16.1 美金金额

只有当备注区域出现 `USD` 时，才尝试识别美金金额。

如果备注区域没有 `USD`：

```text
美金金额 = 无
```

如果备注区域出现 `USD` 但后面没有可识别数字：

```text
美金金额 = 识别失败
```

并且 PDF 会归入 `failed_invoices/`。

脚本还会忽略类似汇率表达中的 `USD1=`，避免把 `USD1=7.1CNY` 误识别成美金金额。

### 16.2 汇率

只有当备注区域出现 `汇率` 时，才尝试识别汇率数字。

如果没有 `汇率`：

```text
汇率 = 无
```

如果有 `汇率` 但后面的数字识别失败：

```text
汇率 = 识别失败
```

并且 PDF 会归入 `failed_invoices/`。

### 16.3 提单号

只有当备注区域出现 `提单号` 时，才尝试识别提单号。

如果没有 `提单号`：

```text
提单号 = 无
```

如果有 `提单号` 但后面的编号识别失败：

```text
提单号 = 识别失败
```

并且 PDF 会归入 `failed_invoices/`。

## 17. 分隔符兼容

脚本定义了统一的分隔符规则：

```python
SEPARATORS = r"[\s:：,，;；、为]*"
```

它允许关键词和值之间存在：

```text
空格
英文冒号
中文冒号
英文逗号
中文逗号
分号
顿号
“为”字
```

因此这些写法都可以被兼容：

```text
USD31825.00
USD 31825.00
USD:31825.00
汇率为7.2818
汇率：7.2818
提单号：ABC123
提单号ABC123
```

## 18. 识别结果判断

核心解析函数是：

```python
parse_invoice(pdf_path)
```

它会返回一个 `InvoiceRecord`。

默认情况下，记录状态是：

```text
成功
```

只要出现任何失败原因，就改为：

```text
失败
```

失败原因会被拼接成中文说明，写入 Excel 的 `失败原因` 列。

## 19. Excel 创建与校验

Excel 工作簿由：

```python
ensure_workbook(path)
```

负责创建或打开。

如果 `invoice_ledger.xlsx` 不存在，脚本会新建它，并写入表头。

如果文件已存在，脚本会检查第一行表头是否与 `HEADERS` 完全一致。如果不一致，会抛出错误并停止运行。

这个校验可以避免用户手动修改表头后，脚本把数据写错列。

## 20. Excel 写入流程

写入函数是：

```python
append_record(record, ledger_path)
```

它的顺序是：

```text
1. 打开或创建 Excel
2. 清理历史中“发票号码识别失败”的失败行
3. 追加当前记录
4. 再次清理“发票号码识别失败”的失败行
5. 按发票号码去重
6. 设置单元格换行和顶部对齐
7. 保存 Excel
```

之所以清理两次，是为了同时覆盖：

- 历史已经存在的无发票号码失败行
- 本次刚识别出来的无发票号码失败行

## 21. 自动清理失败行

自动清理逻辑由以下函数完成：

```python
row_should_be_removed(row_values)
remove_failed_rows_without_invoice_number(sheet)
```

如果某行满足：

```text
识别状态 = 失败
失败原因包含 发票号码识别失败
```

系统会删除整行。

原因是没有发票号码的失败记录无法参与后续去重，也很难和真实发票建立稳定对应关系。保留它们容易造成台账堆积，因此脚本会在下一次运行时自动清理。

## 22. 重复发票号处理

重复发票号处理由：

```python
deduplicate_sheet(sheet)
```

负责。

它会扫描 Excel 中所有记录，以 `发票号码` 作为唯一键。

如果发现同一个发票号码出现多次，系统会只保留一条记录。

保留规则由：

```python
row_completeness_score(row_values)
```

计算。

评分逻辑大致是：

- 识别状态为 `成功` 的记录加高分。
- 发票号码、开票日期、购买方名称、销售方名称、金额、美金金额、汇率、提单号越完整，分数越高。
- `识别失败` 和空值不加分。
- `无` 代表字段确实不存在，会给少量分数。

这样可以实现：

```text
同一发票号码，只保留识别更完整的一条
```

## 23. PDF 文件归档

PDF 移动由：

```python
move_pdf(pdf_path, record)
```

负责。

如果识别状态是：

```text
成功
```

PDF 会移动到：

```text
archived_invoices/
```

如果识别状态是：

```text
失败
```

PDF 会移动到：

```text
failed_invoices/
```

如果目标目录中已经有同名文件，系统不会覆盖旧文件，而是通过：

```python
unique_destination(directory, filename)
```

生成新的文件名，例如在文件名后追加序号。

## 24. dry-run 模式

`--dry-run` 模式用于只看识别结果，不修改任何业务文件。

在 dry-run 模式下：

- 会读取 PDF。
- 会打印识别成功或失败。
- 不会写入 Excel。
- 不会移动 PDF。

这个模式适合在优化识别规则前后进行对比验证。

## 25. 隐私与开源保护

项目通过 `.gitignore` 忽略敏感数据：

```text
*.pdf
*.PDF
*.xlsx
*.xls
*.csv
invoices_to_process/*
archived_invoices/*
failed_invoices/*
.conda/
__pycache__/
```

同时通过 `.gitkeep` 保留空目录结构：

```text
invoices_to_process/.gitkeep
archived_invoices/.gitkeep
failed_invoices/.gitkeep
```

这样 GitHub 仓库中会保留必要目录，但不会上传真实发票和台账。

开源前建议执行：

```bash
git status --ignored
git ls-files
```

确认没有 PDF、Excel、CSV 或其他敏感业务数据被跟踪。

## 26. 后续维护建议

维护这个项目时，最重要的是先判断失败属于哪一层，再去改对应函数。不要一上来大范围重写规则，否则容易修好一种发票、影响另一种发票。

### 26.1 问题定位速查表

| 现象或失败原因 | 优先检查位置 | 主要函数 |
| --- | --- | --- |
| `发票号码识别失败` | 发票号码标签、老版发票号码位置、PDF 文本是否乱序 | `extract_invoice_number()` |
| `开票日期识别失败` | 日期格式是否新增、日期是否被 PDF 拆开 | `extract_invoice_date()` |
| `购买方名称识别失败` | 购买方/销售方是否同一行、是否有空格拆字 | `extract_names()`、`extract_legacy_buyer()` |
| `销售方名称识别失败` | 销售方是否在备注或开户行附近、公司名是否被拆分 | `extract_names()`、`extract_legacy_seller()` |
| `价税合计（小写）识别失败` | 小写金额标签是否变化、金额是否换行 | `extract_total_amount()` |
| `检测到USD关键词，但美金金额识别失败` | USD 是否真的在备注区、后面数字格式是否新增 | `extract_remark_text()`、`extract_keyword_number()` |
| `检测到汇率关键词，但汇率数字识别失败` | 汇率表达方式是否新增，例如 `USD1=...` | `extract_keyword_number()` |
| `检测到提单号关键词，但提单号识别失败` | 提单号是否含新的字符类型或分隔符 | `extract_keyword_number()` |
| 重复发票号保留了不理想的行 | 完整度评分是否需要调整 | `row_completeness_score()`、`deduplicate_sheet()` |
| 没有发票号码的失败行仍保留 | 清理条件是否覆盖当前失败原因 | `row_should_be_removed()` |
| PDF 被移动到了错误目录 | `status` 判断是否正确 | `parse_invoice()`、`move_pdf()` |

### 26.2 推荐维护流程

如果遇到新的发票版式识别失败，可以按以下顺序排查：

```text
1. 先使用 --dry-run 运行脚本
2. 查看终端输出的失败原因
3. 如果是必填字段失败，定位对应 extract_* 函数
4. 如果是 USD、汇率、提单号失败，优先检查备注区域提取是否正确
5. 调整规则后再次 dry-run
6. 确认无误后正式运行
```

建议每次只改一个问题点。例如这次只处理“汇率识别失败”，就尽量只改 `extract_keyword_number()` 或 `extract_remark_text()`，不要同时改购销方、金额、去重等无关逻辑。

### 26.3 常见扩展点

常见扩展点如下：

```text
新增发票号码规则      修改 extract_invoice_number()
新增日期规则          修改 extract_invoice_date()
新增购销方规则        修改 extract_names()
新增金额规则          修改 extract_total_amount()
调整备注范围          修改 extract_remark_text()
调整可选字段识别      修改 extract_keyword_number()
调整台账去重策略      修改 row_completeness_score() 或 deduplicate_sheet()
调整失败清理策略      修改 row_should_be_removed()
```

### 26.4 修改前检查清单

改代码前建议先确认：

- 失败 PDF 是否在 `failed_invoices/` 或测试目录中。
- 能否用 `--dry-run` 复现问题。
- 失败原因具体是哪一项。
- 需要修改的是必填字段、备注字段、Excel 写入，还是文件移动逻辑。
- 这次修改是否可能影响已有成功发票。

### 26.5 修改后验证清单

改代码后建议至少执行：

```bash
./.conda/bin/python -m py_compile invoice_register.py
./.conda/bin/python invoice_register.py --dry-run
```

如果改动涉及 Excel 写入、去重或失败行清理，建议额外用临时 Excel 测试，不要直接污染正式 `invoice_ledger.xlsx`。

如果改动涉及文件移动，先用少量测试 PDF 验证：

```text
1. 放入 invoices_to_process/
2. 运行 --dry-run
3. 确认输出符合预期
4. 再正式运行
5. 检查 PDF 是否进入正确目录
6. 检查 Excel 是否追加正确
```

### 26.6 回归测试建议

每次优化识别规则后，最好保留几类非敏感测试样本或脱敏样本：

- 新版电子发票样本。
- 老版电子普通发票样本。
- 有 USD 和汇率的样本。
- 有提单号的样本。
- 没有 USD、汇率、提单号的样本。
- 需要失败归档的样本。
- 重复发票号样本。

这些样本不建议使用真实业务发票直接提交到 GitHub。如果要开源测试样本，应先脱敏或人工构造。

### 26.7 高风险修改点

以下位置改动时要特别小心：

```text
HEADERS
append_record()
deduplicate_sheet()
remove_failed_rows_without_invoice_number()
move_pdf()
.gitignore
```

原因如下：

- `HEADERS` 影响 Excel 列顺序，改错会导致历史台账无法继续写入。
- `append_record()` 影响写表、清理、去重的整体顺序。
- `deduplicate_sheet()` 影响重复发票号保留结果。
- `remove_failed_rows_without_invoice_number()` 会删除 Excel 行，条件不能过宽。
- `move_pdf()` 会移动真实 PDF 文件，逻辑错误会导致归档混乱。
- `.gitignore` 关系到开源安全，不能误删 PDF 和 Excel 忽略规则。

### 26.8 安全提交检查

提交到 GitHub 前建议执行：

```bash
git -c core.quotePath=false status --short --ignored
git -c core.quotePath=false ls-files
```

确认以下内容没有出现在已跟踪文件中：

```text
*.pdf
*.PDF
*.xlsx
*.xls
*.csv
invoice_ledger.xlsx
真实发票文件名
真实客户或供应商敏感信息
```

如果误把敏感文件加入暂存区，应先取消跟踪：

```bash
git rm --cached path/to/sensitive-file
```

如果敏感文件已经进入提交历史，不要直接推送公开仓库，应先处理 Git 历史。

## 27. 当前系统边界

当前版本有以下边界：

- 不支持图片版或扫描版 PDF。
- 不联网查询汇率。
- 不校验发票真伪。
- 不自动识别所有可能的备注业务字段。
- 不提供图形界面。
- 不支持多人同时写入同一个 Excel。

这些限制是有意保留的。当前项目定位是一个轻量、可本地运行、可持续优化的发票识别和台账生成脚本。

## 28. 一句话总结

本项目的核心逻辑可以概括为：

```text
从 invoices_to_process/ 读取 PDF
提取发票字段
追加并整理 invoice_ledger.xlsx
成功则归档到 archived_invoices/
失败则移动到 failed_invoices/
```

它的设计重点是稳定、可解释、可维护，而不是追求一次性覆盖所有发票版式。每当出现新失败样本，都可以根据失败原因逐步优化规则。

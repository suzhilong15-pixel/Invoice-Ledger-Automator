# Invoice Ledger Automator

## English

Invoice Ledger Automator is a Python tool for extracting structured data from text-based PDF invoices and appending the results to an Excel ledger. It supports batch processing, automatic archiving, failure tracking, and duplicate invoice handling.

The tool is designed for PDF invoices whose text can be selected and copied. Scanned or image-based PDFs require OCR, which is not included in this version.

For a detailed breakdown of the project structure, execution flow, and maintenance guidance, see [STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md](STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md).

### Key Features

- Extracts invoice number, invoice date, buyer name, seller name, total amount, USD amount, exchange rate, and bill of lading number.
- Appends new records to `invoice_ledger.xlsx` without overwriting existing data.
- Moves successfully processed PDFs to `archived_invoices/`.
- Moves failed PDFs to `failed_invoices/` with failure reasons recorded in the ledger.
- Extracts `USD`, `汇率`, and `提单号` only from the invoice remarks area to avoid false matches from bank account information.
- Deduplicates records by invoice number and keeps the more complete result.
- Automatically removes failed ledger rows when the invoice number could not be extracted.
- Provides a dedicated maintenance guide for future rule updates and troubleshooting.

### Project Structure

```text
invoices_to_process/  Put PDF invoices here before running the script
archived_invoices/    Successfully processed invoices are moved here
failed_invoices/      Failed invoices are moved here for review
invoice_ledger.xlsx   Generated Excel ledger
invoice_register.py   Main script
environment.yml       Conda environment configuration
STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md
                      Structure, logic, and maintenance guide
```

### Ledger Fields

The generated Excel ledger contains these columns:

```text
发票号码
开票日期
购买方名称
销售方名称
价税合计（小写）
美金金额
汇率
提单号
源文件名
识别状态
失败原因
处理时间
```

The column names remain in Chinese because they match the invoice fields and the original business workflow.

### Extraction Rules

These fields are required:

```text
发票号码
开票日期
购买方名称
销售方名称
价税合计（小写）
```

If any required field cannot be extracted, the invoice is moved to `failed_invoices/`.

The optional fields `美金金额`, `汇率`, and `提单号` are extracted only from the invoice remarks area:

- If `USD` is not found in the remarks area, `美金金额` is set to `无`.
- If `USD` is found but the following amount cannot be extracted, `美金金额` is set to `识别失败`, and the PDF is moved to `failed_invoices/`.
- If `汇率` is not found in the remarks area, `汇率` is set to `无`.
- If `汇率` is found but the following number cannot be extracted, `汇率` is set to `识别失败`, and the PDF is moved to `failed_invoices/`.
- If `提单号` is not found in the remarks area, `提单号` is set to `无`.
- If `提单号` is found but the following code cannot be extracted, `提单号` is set to `识别失败`, and the PDF is moved to `failed_invoices/`.

If `USD` appears in seller bank account information, such as bank name, account number, or account fields, it is ignored and is not treated as a USD amount keyword.

### Duplicate Handling

If a newly extracted invoice has the same invoice number as an existing ledger record, the script keeps only one record.

The retained record is selected by these rules:

- Prefer records with `成功` status.
- If the status is the same, keep the record with more complete extracted fields.

Failed ledger rows are also cleaned automatically. If a row has `识别状态` set to `失败` and its `失败原因` contains `发票号码识别失败`, that entire row is removed during the next script run because it cannot be matched or deduplicated by invoice number.

### Supported Keyword Formats

The script supports spaces, colons, Chinese colons, commas, and Chinese commas between keywords and values:

```text
USD31825.00
USD 31825.00
USD:31825.00
汇率 7.2813
汇率：7.2813
提单号：YMJAM236512123
提单号986-92700554
```

### Installation

Create the local Conda environment in the project root:

```bash
conda env create -p ./.conda -f environment.yml
```

### Usage

Put PDF invoices into `invoices_to_process/`, then run:

```bash
./.conda/bin/python invoice_register.py
```

After processing:

- Successful invoices are moved to `archived_invoices/`.
- Failed invoices are moved to `failed_invoices/`.
- Results are appended to `invoice_ledger.xlsx`.

To preview extraction results without writing to Excel or moving files, use dry-run mode:

```bash
./.conda/bin/python invoice_register.py --dry-run
```

### Privacy

Invoice PDFs, generated Excel ledgers, local Conda environments, and processing folder contents are ignored by `.gitignore` to reduce the risk of committing confidential data.

## 中文

Invoice Ledger Automator 是一个基于 Python 的发票台账自动化工具，用于从可复制文字的 PDF 发票中提取结构化信息，并追加写入 Excel 台账。它支持批量处理、自动归档、失败记录和重复发票号处理。

本工具适用于文字可以被选中和复制的 PDF 发票。扫描件或图片版 PDF 需要 OCR，本版本暂未包含 OCR 功能。

如果需要了解项目结构、运行流程和后期维护方法，请阅读 [STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md](STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md)。

### 核心功能

- 自动提取发票号码、开票日期、购买方名称、销售方名称、价税合计、美金金额、汇率和提单号。
- 将新识别结果追加写入 `invoice_ledger.xlsx`，不会覆盖旧数据。
- 识别成功的 PDF 自动移动到 `archived_invoices/`。
- 识别失败的 PDF 自动移动到 `failed_invoices/`，并在台账中记录失败原因。
- `USD`、`汇率`、`提单号` 只在发票备注区域识别，避免把银行账号等位置的 `USD` 误判为美金金额。
- 按发票号码自动去重，并保留字段更完整的识别结果。
- 如果失败记录中包含 `发票号码识别失败`，下次运行时会自动删除该台账行。
- 提供专门的维护指南，方便后续优化识别规则和排查问题。

### 项目结构

```text
invoices_to_process/  运行脚本前，把需要识别的 PDF 发票放到这里
archived_invoices/    识别成功的发票会自动移动到这里
failed_invoices/      识别失败的发票会自动移动到这里，方便人工复核
invoice_ledger.xlsx   自动生成和追加更新的 Excel 台账
invoice_register.py   主程序
environment.yml       Conda 环境配置
STRUCTURE_LOGIC_AND_MAINTENANCE_GUIDE.md
                      项目结构、运行逻辑与后期维护指南
```

### 台账字段

生成的 Excel 台账包含以下字段：

```text
发票号码
开票日期
购买方名称
销售方名称
价税合计（小写）
美金金额
汇率
提单号
源文件名
识别状态
失败原因
处理时间
```

字段名保留中文，是为了和发票字段及原始业务流程保持一致。

### 识别规则

以下字段属于必填项：

```text
发票号码
开票日期
购买方名称
销售方名称
价税合计（小写）
```

只要有任意必填项识别失败，PDF 就会进入 `failed_invoices/`。

`美金金额`、`汇率`、`提单号` 只在发票备注区域内识别：

- 如果备注区域没有找到 `USD`，美金金额填写 `无`。
- 如果找到 `USD`，但无法识别后面的数字，美金金额填写 `识别失败`，并归入 `failed_invoices/`。
- 如果备注区域没有找到 `汇率`，汇率填写 `无`。
- 如果找到 `汇率`，但无法识别后面的数字，汇率填写 `识别失败`，并归入 `failed_invoices/`。
- 如果备注区域没有找到 `提单号`，提单号填写 `无`。
- 如果找到 `提单号`，但无法识别后面的编号，提单号填写 `识别失败`，并归入 `failed_invoices/`。

如果 `USD` 出现在销售方开户行、银行账号、账号、账户等信息中，脚本会忽略它，不会把它当作美金金额关键词。

### 重复发票处理

如果新识别的发票和台账中已有记录的发票号码相同，脚本会自动去重，只保留一条记录。

保留规则如下：

- 优先保留识别状态为 `成功` 的记录。
- 如果状态相同，保留字段识别更完整的记录。

脚本还会自动清理失败台账行。如果某一行的 `识别状态` 是 `失败`，并且 `失败原因` 包含 `发票号码识别失败`，那么系统会在下一次运行时删除整行记录，因为这类记录没有发票号码，无法用于后续匹配和去重。

### 支持的关键词格式

脚本兼容关键词和数值之间存在空格、冒号、中文冒号、逗号、中文逗号等情况：

```text
USD31825.00
USD 31825.00
USD:31825.00
汇率 7.2813
汇率：7.2813
提单号：YMJAM236512123
提单号986-92700554
```

### 安装

在项目根目录创建本地 Conda 环境：

```bash
conda env create -p ./.conda -f environment.yml
```

### 使用方法

把 PDF 发票放入 `invoices_to_process/`，然后运行：

```bash
./.conda/bin/python invoice_register.py
```

运行完成后：

- 识别成功的 PDF 会移动到 `archived_invoices/`。
- 识别失败的 PDF 会移动到 `failed_invoices/`。
- 识别结果会追加写入 `invoice_ledger.xlsx`。

如果只想预览识别结果，不写入 Excel，也不移动 PDF，可以使用试运行模式：

```bash
./.conda/bin/python invoice_register.py --dry-run
```

### 隐私说明

发票 PDF、生成的 Excel 台账、本地 Conda 环境和处理目录内容已通过 `.gitignore` 忽略，以降低误提交机密数据的风险。

# Invoice Ledger Automator

A Python-based automation tool for extracting structured data from text-based PDF invoices and appending the results to an Excel ledger. Each run appends new records to `发票台账.xlsx` without overwriting existing data.

一个基于 Python 的自动化工具，用于从可复制文字的 PDF 发票中提取结构化信息，并追加写入 Excel 台账。每次运行都会把新记录追加到 `发票台账.xlsx`，不会覆盖旧数据。

## Project Structure

## 项目结构

```text
待识别发票/      PDF invoices waiting to be processed / 放入需要识别的 PDF 发票
已归档发票/      Successfully processed invoices / 识别成功后自动移动到这里
识别失败发票/    Failed invoices for manual review / 识别失败后自动移动到这里
发票台账.xlsx    Generated Excel ledger / 自动生成和追加更新的 Excel 台账
invoice_register.py Main script / 主程序
environment.yml  Conda environment file / Conda 环境配置
```

## Ledger Fields

## 台账字段

The script writes the following fields to the Excel ledger:

脚本会写入以下字段：

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

## Extraction Rules

## 识别规则

The following fields are required. If any of them cannot be extracted, the PDF will be moved to `识别失败发票/`.

以下字段属于必填项。只要有任意一项识别失败，PDF 就会进入 `识别失败发票/`。

```text
发票号码
开票日期
购买方名称
销售方名称
价税合计（小写）
```

`美金金额`, `汇率`, and `提单号` are only extracted from the invoice remarks area.

`美金金额`、`汇率`、`提单号` 只在发票备注区域内按关键词判断。

- If `USD` is not found in the remarks area, `美金金额` is set to `无`.
- 如果备注区域没有找到 `USD`，美金金额填写 `无`。
- If `USD` is found but the following amount cannot be extracted, `美金金额` is set to `识别失败`, and the PDF is moved to the failed folder.
- 如果找到 `USD`，但无法识别后面的数字，美金金额填写 `识别失败`，并归入失败文件夹。
- If `汇率` is not found in the remarks area, `汇率` is set to `无`.
- 如果备注区域没有找到 `汇率`，汇率填写 `无`。
- If `汇率` is found but the following number cannot be extracted, `汇率` is set to `识别失败`, and the PDF is moved to the failed folder.
- 如果找到 `汇率`，但无法识别后面的数字，汇率填写 `识别失败`，并归入失败文件夹。
- If `提单号` is not found in the remarks area, `提单号` is set to `无`.
- 如果备注区域没有找到 `提单号`，提单号填写 `无`。
- If `提单号` is found but the following code cannot be extracted, `提单号` is set to `识别失败`, and the PDF is moved to the failed folder.
- 如果找到 `提单号`，但无法识别后面的编号，提单号填写 `识别失败`，并归入失败文件夹。

If `USD` appears in seller bank account information, such as bank name, account number, or account fields, it is not treated as a USD amount keyword.

如果 `USD` 出现在销售方开户行、银行账号、账号、账户等信息中，脚本不会把它当作美金金额关键词。

## Duplicate Handling

## 重复发票处理

If a newly extracted invoice has the same invoice number as an existing ledger record, the script keeps only one record.

如果新识别的发票和台账中已有记录的发票号码相同，脚本会自动去重，只保留一条记录。

- Records with `成功` status are preferred.
- 优先保留识别状态为 `成功` 的记录。
- If the status is the same, the record with more complete extracted fields is kept.
- 如果状态相同，保留发票号码、开票日期、购买方名称、销售方名称、价税合计、美金金额、汇率、提单号这些字段中识别更完整的记录。

## Supported Keyword Formats

## 支持的关键词格式

The script supports spaces, colons, Chinese colons, commas, and Chinese commas between keywords and values.

脚本兼容关键词和数值之间存在空格、冒号、中文冒号、逗号、中文逗号等情况。

```text
USD31825.00
USD 31825.00
USD:31825.00
汇率 7.2813
汇率：7.2813
提单号：YMJAM236512123
提单号986-92700554
```

## Installation

## 安装

Create the local Conda environment in the project root:

在项目根目录创建本地 Conda 环境：

```bash
conda env create -p ./.conda -f environment.yml
```

If the environment already exists, you can use it directly:

如果环境已经存在，可以直接使用：

```bash
./.conda/bin/python invoice_register.py --dry-run
```

## Usage

## 使用方法

1. Put PDF invoices into `待识别发票/`.
2. Run the script from the project root.
3. Open `发票台账.xlsx` after processing.

1. 把需要识别的 PDF 发票放入 `待识别发票/`。
2. 在项目根目录运行脚本。
3. 运行完成后查看 `发票台账.xlsx`。

```bash
./.conda/bin/python invoice_register.py
```

Successfully processed PDFs are moved to `已归档发票/`. Failed PDFs are moved to `识别失败发票/`.

识别成功的 PDF 会移动到 `已归档发票/`，识别失败的 PDF 会移动到 `识别失败发票/`。

## Dry Run

## 试运行

Use dry-run mode to preview extraction results without writing to Excel or moving PDF files.

如果只想查看识别结果，不写入 Excel，也不移动 PDF，可以运行试运行模式。

```bash
./.conda/bin/python invoice_register.py --dry-run
```

## Notes

## 注意事项

- This script is designed for text-based PDFs where text can be copied.
- 当前脚本适用于 PDF 内文字可以复制的发票。
- Scanned or image-based PDFs require OCR, which is not included in this version.
- 如果 PDF 是扫描件或图片版，需要另加 OCR 功能。
- If a new invoice format fails, use the `失败原因` field in the ledger to improve extraction rules.
- 如果某类新发票进入失败文件夹，可以根据台账中的 `失败原因` 优化识别规则。

## Privacy

## 隐私说明

Invoice PDFs, generated Excel ledgers, local Conda environments, and processing folders are ignored by `.gitignore` to reduce the risk of committing confidential data.

发票 PDF、生成的 Excel 台账、本地 Conda 环境和处理目录内容已通过 `.gitignore` 忽略，以降低误提交机密数据的风险。

# xlsx2txt

Excel 转 TXT 工具，将 Excel 工作表转换为制表符分隔的文本文件。

## 功能特性

- 自动转换所有工作表（或指定工作表）
- 处理单元格内的特殊字符（换行符、制表符）
- 支持短参数和长参数
- 支持自定义输出分隔符

## 使用方法

```bash
# 基本用法
xlsx2txt -input data.xlsx -prefix output

# 使用短参数
xlsx2txt -i data.xlsx -p output

# 指定输出分隔符
xlsx2txt -i data.xlsx -p output -s ","

# 只转换指定工作表
xlsx2txt -i data.xlsx -p output -sheet Sheet1,Sheet2
```

## 参数说明

| 参数 | 短参数 | 说明 | 默认值 |
|------|--------|------|--------|
| `-input` | `-i` | 输入 Excel 文件路径 | 无（必填） |
| `-prefix` | `-p` | 输出文件前缀 | 输入文件名 |
| `-sep` | `-s` | 输出列分隔符 | `\t`（制表符） |
| `-sheet` | - | 工作表名称列表（逗号分隔） | 空（转换所有工作表） |

## 特殊字符处理

- 换行符 `\r\n` 或 `\n` 转换为 `<br/>`
- 制表符 `\t` 转换为 `&#9;`

## 示例

```bash
# 转换所有工作表
xlsx2txt -i data.xlsx

# 使用逗号作为分隔符
xlsx2txt -i data.xlsx -s ","

# 只转换特定工作表
xlsx2txt -i data.xlsx -sheet Sheet1,Sheet3 -p result
```

## 输出格式

输出文件格式：`{prefix}.{sheetName}.txt`
# xlsx2txt

Excel 转 TXT 工具，将 Excel 工作表转换为制表符分隔的文本文件。

## 功能特性

- 自动转换所有工作表
- 处理单元格内的特殊字符（换行符、制表符）
- 支持自定义输出分隔符

## 使用方法

```bash
xlsx2txt -input input.xlsx -prefix output
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-input` | 输入 Excel 文件路径 | 无（必填） |
| `-prefix` | 输出文件前缀 | 输入文件名 |
| `-sep` | 输出列分隔符 | `\t`（制表符） |

## 特殊字符处理

- 换行符 `\r\n` 或 `\n` 转换为 `<br/>`
- 制表符 `\t` 转换为 `&#9;`

## 示例

```bash
xlsx2txt -input data.xlsx -prefix output
```

输出文件格式：`output.SheetName.txt`
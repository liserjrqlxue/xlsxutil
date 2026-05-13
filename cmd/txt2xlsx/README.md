# txt2xlsx

TXT 文件转 XLSX 工具，支持特殊字符处理。

## 功能特性

- 将 TXT 文件转换为 Excel 格式
- 处理特殊字符（如换行符标记 `[\n]`）
- 支持自定义分隔符和工作表名称

## 使用方法

```bash
txt2xlsx -input data.txt -excel output.xlsx
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-input` | 输入 TXT 文件路径 | 无（必填） |
| `-excel` | 输出 Excel 文件路径 | 无（必填） |
| `-sheet` | 输出工作表名称 | `data` |
| `-sep` | 列分隔符 | `\t`（制表符） |

## 示例

```bash
txt2xlsx -input data.txt -excel output.xlsx -sheet mydata
```
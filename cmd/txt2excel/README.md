# txt2excel

TXT 文件转 Excel 工具，支持多个 TXT 文件合并到一个 Excel 的不同工作表。

## 功能特性

- 支持多个输入 TXT 文件
- 自动创建对应工作表
- 支持自定义分隔符

## 使用方法

```bash
txt2excel -input data1.txt,data2.txt -output result
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-input` | 输入 TXT 文件路径（多个用逗号分隔） | 无（必填） |
| `-output` | 输出文件名（自动添加 `.xlsx` 后缀） | 第一个输入文件名 |
| `-sheet` | 工作表名称（多个用逗号分隔） | 输入文件名 |
| `-sep` | 列分隔符 | `\t`（制表符） |

## 示例

```bash
txt2excel -input data1.txt,data2.txt -output merged -sheet Sheet1,Sheet2
```
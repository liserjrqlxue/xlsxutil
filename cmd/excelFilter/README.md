# excelFilter

Excel 基因过滤工具，根据基因列表筛选 Excel 表格中的变异数据。

## 功能特性

- 读取基因列表文件进行过滤
- 仅保留指定基因的变异记录
- 支持自定义工作表名称

## 使用方法

```bash
excelFilter -input input.xlsx -gene gene.list -output output.xlsx
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-input` | 输入 Excel 文件路径 | 无（必填） |
| `-output` | 输出 Excel 文件路径 | 输入文件名 + `.filter.xlsx` |
| `-gene` | 基因列表文件路径（每行一个基因） | 无（必填） |
| `-sheet` | 需要过滤的工作表名称 | `filter_variants` |

## 示例

```bash
excelFilter -input variants.xlsx -gene gene.list -output filtered.xlsx
```
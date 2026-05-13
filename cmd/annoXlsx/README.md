# annoXlsx

Excel 基因注释工具，用于在 Excel 表格中添加基因是否在 Lancet-PAGE 研究基因集中的标记。

## 功能特性

- 读取输入的 Excel 文件
- 加载基因列表数据库
- 在指定的工作表中添加基因注释列
- 输出带注释的 Excel 文件

## 使用方法

```bash
annoXlsx -input input.xlsx -output output.xlsx
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-input` | 输入 Excel 文件路径 | 无（必填） |
| `-output` | 输出 Excel 文件路径 | 输入文件名 + `.anno.xlsx` |
| `-genelist` | 基因列表 Excel 文件路径 | 内置 Lancet-PAGE 基因集 |
| `-genesheet` | 基因列表工作表名称 | `1621个基因` |
| `-annosheet` | 需要注释的工作表名称 | `filter_variants` |
| `-annotitle` | 注释列标题 | `是否lancet记录基因` |

## 示例

```bash
annoXlsx -input variants.xlsx -output variants_anno.xlsx
```
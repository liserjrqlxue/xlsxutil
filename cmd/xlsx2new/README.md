# xlsx2new

Excel 变异注释工具，用于对基因变异数据进行全面注释和处理。

## 功能特性

- 支持 SNV 变异注释（`filter_variants` 工作表）
- 支持 CNV 变异注释（`exon_cnv` 工作表）
- 支持 GnomAD 频率注释
- 支持 ACMG 标准变异分类
- 支持基因疾病数据库注释
- 支持自动化突变判定

## 使用方法

```bash
xlsx2new -input input.xlsx -output output.xlsx
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-input` | 输入 Excel 文件路径 | 无（必填） |
| `-output` | 输出 Excel 文件路径 | 无（必填） |
| `-acmg` | ACMG 基因数据库文件 | 内置 ACMG59 基因集 |
| `-acmgSheet` | ACMG 数据库工作表名称 | `ACMG推荐59个基因` |
| `-geneDb` | 基因库文件路径 | 内置基因特征谱 |
| `-title` | 输出列标题列表文件 | `etc/title.txt` |
| `-annoCnv` | 是否注释 CNV | `false` |
| `-annoGnomAD` | 是否更新 GnomAD 信息 | `false` |
| `-gnomAD` | GnomAD VCF 文件路径 | 内置路径 |
| `-annoACMG` | 是否进行 ACMG 分类 | `false` |
| `-gender` | 先证者性别 | 空 |

## 示例

```bash
# 基础注释
xlsx2new -input variants.xlsx -output annotated.xlsx

# 完整注释（包含 GnomAD 和 ACMG）
xlsx2new -input variants.xlsx -output annotated.xlsx -annoGnomAD -annoACMG -gender M
```
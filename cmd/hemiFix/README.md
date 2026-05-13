# hemiFix

半合子修复工具，用于修正男性样本在 X/Y 染色体非 PAR 区域的基因型表示。

## 功能特性

- 自动识别 X/Y 染色体上的变异
- 根据样本性别修正半合子基因型（将 Hom 改为 Hemi）
- 正确处理 PAR（拟常染色体区域）

## 使用方法

```bash
hemiFix -input input.xlsx -output output.xlsx -gender M
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-input` | 输入 Excel 文件路径 | 无（必填） |
| `-output` | 输出 Excel 文件路径 | 输入文件名 + `.fixHemi.xlsx` |
| `-gender` | 样本性别（M/F，多个用逗号分隔） | 无（必填） |

## 示例

```bash
hemiFix -input variants.xlsx -output fixed.xlsx -gender M
```
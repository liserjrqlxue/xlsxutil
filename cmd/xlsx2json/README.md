# xlsx2json

Excel 转 JSON 工具，支持将 Excel 工作表转换为 JSON 格式。

## 功能特性

- 支持单个或多个工作表转换
- 支持按列名作为键值的映射格式
- 支持 AES 加密输出
- 支持自定义合并分隔符

## 使用方法

```bash
xlsx2json -xlsx input.xlsx -prefix output
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-xlsx` | 输入 Excel 文件路径 | 无（必填） |
| `-sheet` | 工作表名称（不指定则转换所有工作表） | 空（全部转换） |
| `-prefix` | 输出文件前缀 | 输入文件名 |
| `-key` | 作为行键的列名（生成 Map 格式） | 空（生成数组格式） |
| `-sep` | 合并行的分隔符 | `\n` |
| `-aes` | 是否进行 AES 加密 | `false` |
| `-codeKey` | AES 加密密钥 | `c3d112d6a47a0a04aad2b9d2d2cad266` |

## 示例

```bash
# 转换所有工作表
xlsx2json -xlsx data.xlsx

# 转换指定工作表并按列名作为键
xlsx2json -xlsx data.xlsx -sheet Sheet1 -key ID

# 转换并加密输出
xlsx2json -xlsx data.xlsx -aes -prefix encrypted
```
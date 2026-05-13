# xlsxutil

一个用于处理 Excel 文件的 Go 工具集，包含多个独立的命令行工具。

## 项目结构

```
xlsxutil/
├── cmd/                    # 可执行命令目录
│   ├── annoXlsx/           # Excel 基因注释工具
│   ├── excelFilter/        # Excel 基因过滤工具
│   ├── hemiFix/            # 半合子修复工具
│   ├── txt2excel/          # TXT 转 Excel 工具
│   ├── txt2xlsx/           # TXT 转 XLSX 工具
│   ├── xlsx2json/          # Excel 转 JSON 工具
│   ├── xlsx2new/           # Excel 变异注释工具
│   └── xlsx2txt/           # Excel 转 TXT 工具
├── go.mod                  # Go 模块依赖
├── go.sum                  # Go 模块校验和
└── LICENSE                 # 许可证文件
```

## 工具列表

| 工具 | 功能描述 |
|------|----------|
| [annoXlsx](cmd/annoXlsx/README.md) | 在 Excel 表格中添加基因是否在 Lancet-PAGE 研究基因集中的标记 |
| [excelFilter](cmd/excelFilter/README.md) | 根据基因列表筛选 Excel 表格中的变异数据 |
| [hemiFix](cmd/hemiFix/README.md) | 修正男性样本在 X/Y 染色体非 PAR 区域的基因型表示 |
| [txt2excel](cmd/txt2excel/README.md) | 将多个 TXT 文件合并到一个 Excel 的不同工作表 |
| [txt2xlsx](cmd/txt2xlsx/README.md) | 将 TXT 文件转换为 Excel 格式，支持特殊字符处理 |
| [xlsx2json](cmd/xlsx2json/README.md) | 将 Excel 工作表转换为 JSON 格式 |
| [xlsx2new](cmd/xlsx2new/README.md) | 对基因变异数据进行全面注释和处理 |
| [xlsx2txt](cmd/xlsx2txt/README.md) | 将 Excel 工作表转换为制表符分隔的文本文件 |

## 安装

```bash
# 克隆项目
git clone https://github.com/liserjrqlxue/xlsxutil.git
cd xlsxutil

# 安装所有工具
go install ./cmd/...
```

## 构建

```bash
# 构建所有工具
go build ./cmd/...

# 构建指定工具
go build -o bin/xlsx2txt ./cmd/xlsx2txt
```

## 使用示例

```bash
# 将 Excel 转换为 TXT
xlsx2txt -input data.xlsx -prefix output

# 将 TXT 转换为 Excel
txt2xlsx -input data.txt -excel output.xlsx

# 注释 Excel 中的基因
annoXlsx -input variants.xlsx -output annotated.xlsx

# 根据基因列表过滤
excelFilter -input variants.xlsx -gene gene.list -output filtered.xlsx
```

## 依赖

- [github.com/tealeg/xlsx](https://github.com/tealeg/xlsx) - Excel 文件读写
- [github.com/xuri/excelize/v2](https://github.com/xuri/excelize) - Excel 文件处理
- [github.com/liserjrqlxue/simple-util](https://github.com/liserjrqlxue/simple-util) - 工具函数库

## 许可证

MIT License
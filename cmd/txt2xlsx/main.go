package main

import (
	"flag"
	"os"
	"regexp"

	"github.com/liserjrqlxue/simple-util"
	"github.com/xuri/excelize/v2"
)

var (
	outputExcel = flag.String(
		"excel",
		"",
		"output excel file name",
	)
	sheetName = flag.String(
		"sheet",
		"data",
		"output sheet name",
	)
	inputData = flag.String(
		"input",
		"",
		"inout txt file",
	)
	sep = flag.String(
		"sep",
		"\t",
		"sep for split txt column",
	)
)

var (
	newLine = regexp.MustCompile(`\[\n\]`)
)

var newLineList = []string{
	"PP_interpretation", "PP_mutation information",
	"PP_突变定义", "PP_突变详情",
	"PP_References", "PP_References other1", "PP_References other2",
	"PP_disGroup",
	"中文-疾病名称", "中文-疾病背景", "中文-治疗与干预", "中文-突变判定",
	"遗传模式",
	"发病率",
	"发病率-EN",
	"中文-突变详情",
	"中文-疾病简介",
	"英文-疾病简介",
	"参考文献-原有", "参考文献-新增",
	"自动化突变判定", "证据分类",
	"英文-疾病名称", "英文-疾病背景", "英文-治疗与干预", "英文-突变判定",
	"英文-突变详情",
	"Reference", "Evidence Classification",
	"Reference-final-Info",
	"备注",
	"note2",
	"Database",
}

func main() {
	flag.Parse()
	if *inputData == "" || *outputExcel == "" {
		flag.Usage()
		os.Exit(0)
	}
	outputXlsx := excelize.NewFile()
	outputXlsx.NewSheet(*sheetName)

	dataArray, titleList := simple_util.File2MapArray(*inputData, *sep, nil)

	for j, title := range titleList {
		axis, _ := excelize.CoordinatesToCellName(j+1, 1)
		outputXlsx.SetCellValue(*sheetName, axis, title)
	}

	for i, dataHash := range dataArray {
		for _, title := range newLineList {
			dataHash[title] = newLine.ReplaceAllString(dataHash[title], "\n")
		}
		for j, title := range titleList {
			axis, _ := excelize.CoordinatesToCellName(j+1, i+2)
			outputXlsx.SetCellValue(*sheetName, axis, dataHash[title])
		}
	}

	err := outputXlsx.SaveAs(*outputExcel)
	simple_util.CheckErr(err)
}
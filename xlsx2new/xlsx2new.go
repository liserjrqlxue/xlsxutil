package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"sort"
)

func main() {
	var inputXlsxPath, outputPrefix string
	if len(os.Args) > 2 {
		inputXlsxPath = os.Args[1]
		outputPrefix = os.Args[2]
	} else {
		fmt.Println(os.Args[0], "input.xlsx", "outputPrefix")
		os.Exit(1)
	}

	// 读取input.xlsx
	inputXlsx, err := excelize.OpenFile(inputXlsxPath)
	checkError(err)

	// 生成新excel
	outputXlsx := excelize.NewFile()

	// 复制工作表
	// 便利input.xlsx的工作表
	sheetMap := inputXlsx.GetSheetMap()
	var keys []int
	for k := range sheetMap {
		keys = append(keys, k)
	}
	sort.Ints(keys)
	for _, k := range keys {
		name := inputXlsx.GetSheetName(k)
		fmt.Printf("Copy sheet %d [%s]\n", k, name)
		inputRows := inputXlsx.GetRows(name)
		outputXlsx.NewSheet(name)
		for i, row := range inputRows {
			for j, cell := range row {
				axis := positionToAxis(i, j)
				outputXlsx.SetCellValue(name, axis, cell)
			}
		}
	}
	// 保存到 outputPrefix.xlsx
	outputXlsx.DeleteSheet("Sheet1")
	err = outputXlsx.SaveAs(outputPrefix + ".xlsx")
	checkError(err)
}

func checkError(e error) {
	if e != nil {
		panic(e)
	}
}

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
	/*  中文乱码问题待解决，用title.xlsx替代
	file,err:=os.Open("title.txt")
	checkError(err)
	defer file.Close()

	var titleList []string
	scanner:=bufio.NewScanner(file)
	for scanner.Scan(){
		fmt.Println(scanner.Text())
		titleList=append(titleList,scanner.Text())
	}
	checkError(scanner.Err())
	*/
	titleXlsx, err := excelize.OpenFile("title.xlsx")
	checkError(err)
	titleList := titleXlsx.GetRows("title")[0]

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
		sheetName := inputXlsx.GetSheetName(k)
		fmt.Printf("Copy sheet %d [%s]\n", k, sheetName)
		if sheetName == "filter_variants" {
			annoSheet(inputXlsx, outputXlsx, sheetName, titleList)
		} else {
			copySheet(inputXlsx, outputXlsx, sheetName)
		}
	}
	// 保存到 outputPrefix.xlsx
	outputXlsx.DeleteSheet("Sheet1")
	fmt.Printf("sheetName:%s, sheetIndex:%d",
		sheetMap[1], outputXlsx.GetSheetIndex(sheetMap[1]))
	outputXlsx.SetActiveSheet(1)
	err = outputXlsx.SaveAs(outputPrefix + ".xlsx")
	checkError(err)
}

func annoSheet(inputXlsx, outputXlsx *excelize.File, sheetName string, titleList []string) error {
	inputRows := inputXlsx.GetRows(sheetName)
	outputXlsx.NewSheet(sheetName)
	var keysList []string
	for i, row := range inputRows {
		if i == 0 {
			for _, cell := range row {
				keysList = append(keysList, cell)
			}
			for j, title := range titleList {
				axis := positionToAxis(i, j)
				outputXlsx.SetCellValue(sheetName, axis, title)
			}
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				dataHash[keysList[j]] = cell
			}
			for j, title := range titleList {
				axis := positionToAxis(i, j)
				outputXlsx.SetCellValue(sheetName, axis, dataHash[title])
			}
		}
	}
	return nil
}

func copySheet(inputXlsx, outputXlsx *excelize.File, sheetName string) error {
	inputRows := inputXlsx.GetRows(sheetName)
	outputXlsx.NewSheet(sheetName)
	for i, row := range inputRows {
		for j, cell := range row {
			axis := positionToAxis(i, j)
			outputXlsx.SetCellValue(sheetName, axis, cell)
		}
	}
	return nil
}

func checkError(e error) {
	if e != nil {
		panic(e)
	}
}

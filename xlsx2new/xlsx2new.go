package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	xlsx := excelize.NewFile()
	// 创建一个工作表
	index := xlsx.NewSheet("Sheet2")
	// 设置单元格的值
	xlsx.SetCellValue("Sheet2", "A2", "Hello word.")
	xlsx.SetCellValue("Sheet1", "B2", 100)
	// 设置工作簿的默认工作表
	xlsx.SetActiveSheet(index)
	// 根据指定路径保存文件
	err := xlsx.SaveAs("./Book1.xlsx")
	checkError(err)

	inputXlsx, err := excelize.OpenFile("./Book1.xlsx")
	checkError(err)
	// 获取工作表中指定单元格的值
	cell := inputXlsx.GetCellValue("Sheet1", "B2")
	fmt.Println(cell)
	// 获取Sheet1上所有单元格
	rows := inputXlsx.GetRows("Sheet1")
	for i, row := range rows {

		fmt.Print("row:", i, "\t[")
		for _, cell := range row {
			fmt.Print(cell, "]\t[")
		}
		fmt.Println("]")
	}
}

func checkError(e error) {
	if e != nil {
		panic(e)
	}
}

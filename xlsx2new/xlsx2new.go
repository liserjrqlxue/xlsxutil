package main

import (
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
}

func checkError(e error) {
	if e != nil {
		panic(e)
	}
}

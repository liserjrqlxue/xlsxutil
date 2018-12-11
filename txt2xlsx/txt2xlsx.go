package main

import (
	"flag"
	"github.com/liserjrqlxue/simple-util"
	"github.com/tealeg/xlsx"
	"os"
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

func main() {
	flag.Parse()
	if *inputData == "" || *outputExcel == "" {
		flag.Usage()
		os.Exit(0)
	}
	outputXlsx := xlsx.NewFile()
	outputSheet, err := outputXlsx.AddSheet(*sheetName)
	simple_util.CheckErr(err)

	mapDb := simple_util.File2MapDb(*inputData, *sep)
	outputRow := outputSheet.AddRow()
	for _, title := range mapDb.Title {
		outputRow.AddCell().SetString(title)
	}

	for _, dataHash := range mapDb.Data {
		outputRow := outputSheet.AddRow()
		for _, title := range mapDb.Title {
			outputRow.AddCell().SetString(dataHash[title])
		}
	}

	err = outputXlsx.Save(*outputExcel)
	simple_util.CheckErr(err)
}

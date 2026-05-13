package main

import (
	"flag"
	"os"
	"path/filepath"

	"github.com/liserjrqlxue/simple-util"
	"github.com/xuri/excelize/v2"
)

var (
	ex, _  = os.Executable()
	exPath = filepath.Dir(ex)
	pSep   = string(os.PathSeparator)
)

var (
	input = flag.String(
		"input",
		"",
		"input excel",
	)
	output = flag.String(
		"output",
		"",
		"output excel",
	)
	genelist = flag.String(
		"genelist",
		exPath+pSep+"Lancet-PAGE研究1621个基因集.xlsx",
		"gene list ",
	)
	geneSheet = flag.String(
		"genesheet",
		"1621个基因",
		"sheet name of gene list",
	)
	annoSheet = flag.String(
		"annosheet",
		"filter_variants",
		"anno sheet of input",
	)
	annoTitle = flag.String(
		"annotitle",
		"是否lancet记录基因",
		"anno title of sheet",
	)
)

func main() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		os.Exit(0)
	}
	if *output == "" {
		*output = *input + ".anno.xlsx"
	}

	geneXlsx, err := excelize.OpenFile(*genelist)
	simple_util.CheckErr(err)
	var inGeneList = make(map[string]bool)
	rows, err := geneXlsx.GetRows(*geneSheet)
	simple_util.CheckErr(err)
	for _, row := range rows {
		if len(row) > 0 {
			inGeneList[row[0]] = true
		}
	}

	inputXlsx, err := excelize.OpenFile(*input)
	simple_util.CheckErr(err)

	outputXlsx := excelize.NewFile()

	for _, sheetName := range inputXlsx.GetSheetList() {
		if sheetName == *annoSheet {
			updateSheet(inputXlsx, outputXlsx, sheetName, inGeneList)
		} else {
			copySheet(inputXlsx, outputXlsx, sheetName)
		}
	}
	simple_util.CheckErr(outputXlsx.SaveAs(*output))
}

func copySheet(inputXlsx, outputXlsx *excelize.File, sheetName string) {
	rows, err := inputXlsx.GetRows(sheetName)
	simple_util.CheckErr(err)
	outputXlsx.NewSheet(sheetName)
	for i, row := range rows {
		for j, cell := range row {
			axis, _ := excelize.CoordinatesToCellName(j+1, i+1)
			outputXlsx.SetCellValue(sheetName, axis, cell)
		}
	}
}

func updateSheet(inputXlsx *excelize.File, outputXlsx *excelize.File, sheetName string, inGeneList map[string]bool) {
	rows, err := inputXlsx.GetRows(sheetName)
	simple_util.CheckErr(err)

	nrow := len(rows)
	if nrow < 1 {
		return
	}

	outputXlsx.NewSheet(sheetName)
	var keysList []string
	for j, cell := range rows[0] {
		keysList = append(keysList, cell)
		axis, _ := excelize.CoordinatesToCellName(j+1, 1)
		outputXlsx.SetCellValue(sheetName, axis, cell)
	}
	axis, _ := excelize.CoordinatesToCellName(len(keysList)+1, 1)
	outputXlsx.SetCellValue(sheetName, axis, *annoTitle)
	keysList = append(keysList, *annoTitle)

	if nrow > 1 {
		for i := 1; i < nrow; i++ {
			var item = make(map[string]string)
			row := rows[i]
			for j, cell := range row {
				if j < len(keysList) {
					item[keysList[j]] = cell
				}
			}
			if inGeneList[item["Gene Symbol"]] {
				item[*annoTitle] = "是"
			}
			for j, key := range keysList {
				axis, _ := excelize.CoordinatesToCellName(j+1, i+1)
				outputXlsx.SetCellValue(sheetName, axis, item[key])
			}
		}
	}
}
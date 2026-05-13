package main

import (
	"flag"
	_ "net/http/pprof"
	"os"
	"strings"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/textUtil"
	"github.com/xuri/excelize/v2"
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
	geneList = flag.String(
		"gene",
		"",
		"gene list to filter",
	)
	sheetName = flag.String(
		"sheet",
		"filter_variants",
		"sheet to be filter")
)

var inGene = make(map[string]bool)

func main() {
	flag.Parse()
	if *input == "" || *geneList == "" {
		flag.Usage()
		os.Exit(0)
	}
	if *output == "" {
		*output = *input + ".filter.xlsx"
	}
	for _, gene := range textUtil.File2Array(*geneList) {
		inGene[gene] = true
	}

	inputXlsx, err := excelize.OpenFile(*input)
	simpleUtil.CheckErr(err)
	outputXlsx := excelize.NewFile()
	for _, sheet := range inputXlsx.GetSheetList() {
		switch sheet {
		case *sheetName:
			filterSheet(inputXlsx, outputXlsx, sheet, inGene)
		default:
		}
	}
	simpleUtil.CheckErr(outputXlsx.SaveAs(*output))
}

func filterSheet(inputXlsx *excelize.File, outputXlsx *excelize.File, sheetName string, inGene map[string]bool) {
	rows, err := inputXlsx.GetRows(sheetName)
	simpleUtil.CheckErr(err)

	nrow := len(rows)
	if nrow < 1 {
		return
	}

	outputXlsx.NewSheet(sheetName)
	var keysList []string
	outputRow := 1
	for j, cell := range rows[0] {
		text := strings.Split(cell, "*(")[0]
		keysList = append(keysList, text)
		axis, _ := excelize.CoordinatesToCellName(j+1, outputRow)
		outputXlsx.SetCellValue(sheetName, axis, cell)
	}
	outputRow++

	if nrow > 1 {
		for i := 1; i < nrow; i++ {
			var item = make(map[string]string)
			row := rows[i]
			for j, cell := range row {
				if j < len(keysList) {
					item[keysList[j]] = cell
				}
			}
			gene := item["Gene Symbol"]
			if inGene[gene] {
				for j, key := range keysList {
					axis, _ := excelize.CoordinatesToCellName(j+1, outputRow)
					outputXlsx.SetCellValue(sheetName, axis, item[key])
				}
				outputRow++
			}
		}
	}
}

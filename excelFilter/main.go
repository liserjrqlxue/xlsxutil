package main

import (
	"errors"
	"flag"
	"github.com/liserjrqlxue/simple-util"
	"github.com/tealeg/xlsx"
	"os"
	"strings"
)

// flag
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
	for _, gene := range simple_util.File2Array(*geneList) {
		inGene[gene] = true
	}

	inputXlsx, err := xlsx.OpenFile(*input)
	simple_util.CheckErr(err)
	outputXlsx := xlsx.NewFile()
	for _, sheet := range inputXlsx.Sheets {
		switch sheet.Name {
		case *sheetName:
			simple_util.CheckErr(filterSheet(*sheet, outputXlsx, sheet.Name, inGene))
		default:
			//outputXlsx.AppendSheet(*sheet, sheet.Name)
		}
	}
	simple_util.CheckErr(outputXlsx.Save(*output))
}

func filterSheet(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string, inGene map[string]bool) error {
	outputSheet, err := outputXlsx.AddSheet(sheetName)
	simple_util.CheckErr(err)

	nrow := len(sheet.Rows)
	if nrow < 1 {
		return errors.New("error sheet")
	}

	var keysList []string
	var outputRow = outputSheet.AddRow()
	for _, cell := range sheet.Rows[0].Cells {
		text, _ := cell.FormattedValue()
		keysList = append(keysList, strings.Split(text, "*(")[0])
		outputRow.AddCell().SetString(text)
	}
	if nrow > 1 {
		for i := 1; i < nrow; i++ {
			var item = make(map[string]string)
			row := sheet.Rows[i]
			for j, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				item[keysList[j]] = text
			}
			gene := item["Gene Symbol"]
			if inGene[gene] {
				var outputRow = outputSheet.AddRow()
				for _, key := range keysList {
					outputRow.AddCell().SetString(item[key])
				}
			}
		}
	}
	return nil
}

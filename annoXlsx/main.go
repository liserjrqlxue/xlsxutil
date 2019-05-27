package main

import (
	"errors"
	"flag"
	"github.com/liserjrqlxue/simple-util"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
)

// os
var (
	ex, _  = os.Executable()
	exPath = filepath.Dir(ex)
	pSep   = string(os.PathSeparator)
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

	geneXlsx, err := xlsx.OpenFile(*genelist)
	var inGeneList = make(map[string]bool)
	rows := geneXlsx.Sheet[*geneSheet].Rows
	for _, row := range rows {
		inGeneList[row.Cells[0].Value] = true
	}

	inputXlsx, err := xlsx.OpenFile(*input)
	simple_util.CheckErr(err)

	outputXlsx := xlsx.NewFile()

	for _, sheet := range inputXlsx.Sheets {
		switch sheet.Name {
		case *annoSheet:
			simple_util.CheckErr(updateSheet(*sheet, outputXlsx, inGeneList))
		default:
			_, err := outputXlsx.AppendSheet(*sheet, sheet.Name)
			simple_util.CheckErr(err)
		}
	}
	simple_util.CheckErr(outputXlsx.Save(*output))
}

func updateSheet(sheet xlsx.Sheet, outputXlsx *xlsx.File, inGeneList map[string]bool) error {
	outputSheet, err := outputXlsx.AddSheet(sheet.Name)
	simple_util.CheckErr(err)

	nrow := len(sheet.Rows)
	if nrow < 1 {
		return errors.New("error sheet")
	}

	var keysList []string
	var outputRow = outputSheet.AddRow()
	for _, cell := range sheet.Rows[0].Cells {
		text, _ := cell.FormattedValue()
		keysList = append(keysList, text)
		outputRow.AddCell().SetString(text)
	}
	outputRow.AddCell().SetString(*annoTitle)
	keysList = append(keysList, *annoTitle)

	if nrow > 1 {
		for i := 1; i < nrow; i++ {
			var outputRow = outputSheet.AddRow()
			var item = make(map[string]string)
			row := sheet.Rows[i]
			for j, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				item[keysList[j]] = text
			}
			if inGeneList[item["Gene Symbol"]] {
				item[*annoTitle] = "是"
			}
			for _, key := range keysList {
				outputRow.AddCell().SetString(item[key])
			}
		}
	}
	return nil
}

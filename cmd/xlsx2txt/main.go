package main

import (
	"flag"
	"os"
	"regexp"

	"github.com/liserjrqlxue/goUtil/fmtUtil"
	"github.com/liserjrqlxue/goUtil/osUtil"
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/xuri/excelize/v2"
)

var (
	input = flag.String(
		"input",
		"",
		"input excel",
	)
	prefix = flag.String(
		"prefix",
		"",
		"output prefix, default is -xlsx",
	)
	sep = flag.String(
		"sep",
		"\t",
		"output sep",
	)
)
var (
	reg1 = regexp.MustCompile("\r\n")
	reg2 = regexp.MustCompile("\n")
	reg3 = regexp.MustCompile("\t")
)

func main() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		os.Exit(1)
	}
	if *prefix == "" {
		*prefix = *input
	}
	var xlsxF = simpleUtil.HandleError(excelize.OpenFile(*input)).(*excelize.File)
	for _, sheetName := range xlsxF.GetSheetMap() {
		var w = osUtil.Create(*prefix + "." + sheetName + ".txt")
		var rows = simpleUtil.HandleError(xlsxF.GetRows(sheetName)).([][]string)
		for _, row := range rows {
			var rowV []string
			for _, cell := range row {
				cell = reg1.ReplaceAllString(cell, "<br/>")
				cell = reg2.ReplaceAllString(cell, "<br/>")
				cell = reg3.ReplaceAllString(cell, "&#9;")
				rowV = append(rowV, cell)
			}
			fmtUtil.FprintStringArray(w, rowV, *sep)
		}
		simpleUtil.DeferClose(w)
	}
}

package main

import (
	"flag"
	"os"
	"regexp"
	"strings"

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
	inputShort = flag.String(
		"i",
		"",
		"input excel (short)",
	)
	prefix = flag.String(
		"prefix",
		"",
		"output prefix, default is input filename",
	)
	prefixShort = flag.String(
		"p",
		"",
		"output prefix, default is input filename (short)",
	)
	sep = flag.String(
		"sep",
		"\t",
		"output sep",
	)
	sepShort = flag.String(
		"s",
		"\t",
		"output sep (short)",
	)
	sheetList = flag.String(
		"sheet",
		"",
		"sheet names, comma separated, default is all sheets",
	)
)

var (
	reg1 = regexp.MustCompile("\r\n")
	reg2 = regexp.MustCompile("\n")
	reg3 = regexp.MustCompile("\t")
)

func main() {
	flag.Parse()

	inputFile := *input
	if inputFile == "" {
		inputFile = *inputShort
	}
	if inputFile == "" {
		flag.Usage()
		os.Exit(1)
	}

	outputPrefix := *prefix
	if outputPrefix == "" {
		outputPrefix = *prefixShort
	}
	if outputPrefix == "" {
		outputPrefix = inputFile
	}

	outputSep := *sep
	if outputSep == "" {
		outputSep = *sepShort
	}
	if outputSep == "" {
		outputSep = "\t"
	}

	var sheets []string
	var sheetSet map[string]bool
	if *sheetList != "" {
		sheets = strings.Split(*sheetList, ",")
		sheetSet = make(map[string]bool)
		for _, sheet := range sheets {
			sheetSet[strings.TrimSpace(sheet)] = true
		}
	}

	xlsxF := simpleUtil.HandleError(excelize.OpenFile(inputFile))
	allSheets := xlsxF.GetSheetList()

	for _, sheetName := range allSheets {
		if sheetSet != nil && !sheetSet[sheetName] {
			continue
		}
		var w = osUtil.Create(outputPrefix + "." + sheetName + ".txt")
		var rows = simpleUtil.HandleError(xlsxF.GetRows(sheetName))
		for _, row := range rows {
			var rowV []string
			for _, cell := range row {
				cell = reg1.ReplaceAllString(cell, "<br/>")
				cell = reg2.ReplaceAllString(cell, "<br/>")
				cell = reg3.ReplaceAllString(cell, "&#9;")
				rowV = append(rowV, cell)
			}
			fmtUtil.FprintStringArray(w, rowV, outputSep)
		}
		simpleUtil.DeferClose(w)
	}
}

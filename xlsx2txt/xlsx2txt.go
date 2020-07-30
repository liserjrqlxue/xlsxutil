package main

import (
	"flag"
	"os"
	"regexp"

	"github.com/liserjrqlxue/goUtil/fmtUtil"
	"github.com/liserjrqlxue/goUtil/osUtil"
	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/xlsxUtil"
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
	var xlsxF = xlsxUtil.OpenFile(*input)
	for sheetName, sheet := range xlsxF.Sheet {
		var w = osUtil.Create(*prefix + "." + sheetName + ".xlsx")
		defer simpleUtil.DeferClose(w)
		for _, row := range sheet.Rows {
			var rowV []string
			for _, cell := range row.Cells {
				var value = cell.Value
				value = reg1.ReplaceAllString(value, "<br/>")
				value = reg2.ReplaceAllString(value, "<br/>")
				value = reg3.ReplaceAllString(value, "&#9;")
				rowV = append(rowV, value)
			}
			fmtUtil.FprintStringArray(w, rowV, *sep)
		}
	}
}

func check(e error) {
	if e != nil {
		panic(e)
	}
}

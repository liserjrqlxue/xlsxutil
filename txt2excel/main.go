package main

import (
	"flag"
	"os"
	"path/filepath"
	"strings"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
	"github.com/liserjrqlxue/goUtil/textUtil"
	"github.com/liserjrqlxue/goUtil/xlsxUtil"
	"github.com/tealeg/xlsx/v2"
)

var (
	input = flag.String(
		"input",
		"",
		"input txt, comma as sep",
	)
	output = flag.String(
		"output",
		"",
		"output name, .xlsx as suffix, default is first input",
	)
	sheetName = flag.String(
		"sheet",
		"",
		"sheet names, comma as sep, default is basename",
	)
	sep = flag.String(
		"sep",
		"\t",
		"sep for load input to slice",
	)
)

func main() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		os.Exit(1)
	}
	var inputList = strings.Split(*input, ",")
	if *output == "" {
		*output = inputList[0]
	}
	var sheetNames []string
	var sheetNamesMap = make(map[string]bool)
	if *sheetName == "" {
		for _, path := range inputList {
			path = filepath.Base(path)
			sheetNames = append(sheetNames, path)
			sheetNamesMap[path] = true
		}
	}
	if len(sheetNamesMap) != len(sheetNames) || len(sheetNames) != len(inputList) {
		panic("sheetNames error!")
	}
	var excel = xlsx.NewFile()
	for i := range inputList {
		var sheet = xlsxUtil.AddSheet(excel, sheetNames[i])
		xlsxUtil.AddSlice2Sheet(textUtil.File2Slice(inputList[i], *sep), sheet)
	}
	simpleUtil.CheckErr(excel.Save(*output + ".xlsx"))
}

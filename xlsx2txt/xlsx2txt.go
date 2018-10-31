package main

import (
	"encoding/csv"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"regexp"
	"strconv"
)

func main() {
	var xlsxPath, prefix string
	if len(os.Args) > 2 {
		xlsxPath = os.Args[1]
		prefix = os.Args[2]
	} else {
		fmt.Println(os.Args[0], ".xlsx", "prefix.i.tsv")
		os.Exit(1)
	}
	mySlice, e := xlsx.FileToSlice(xlsxPath)
	check(e)

	reg1, e := regexp.Compile("\r\n")
	reg2, e := regexp.Compile("\n")
	reg3, e := regexp.Compile("\t")
	check(e)
	for i, myTsv := range mySlice {
		if len(myTsv) == 0 {
			continue
		}
		file, e := os.Create(prefix + "." + strconv.Itoa(i) + ".tsv")
		check(e)
		defer file.Close()

		writer := csv.NewWriter(file)
		writer.Comma = '\t'
		defer writer.Flush()

		for _, value := range myTsv {
			for k, cell := range value {
				value[k] = reg1.ReplaceAllString(cell, "[n]")
				value[k] = reg2.ReplaceAllString(value[k], "[n]")
				value[k] = reg3.ReplaceAllString(value[k], "[\\t]")
			}
			e := writer.Write(value)
			check(e)
		}
	}
}

func check(e error) {
	if e != nil {
		panic(e)
	}
}

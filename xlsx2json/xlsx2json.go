package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/liserjrqlxue/simple-util"
	"os"
)

var (
	xlsx = flag.String(
		"xlsx",
		"",
		"input excel",
	)
	sheet = flag.String(
		"sheet",
		"",
		"input excel sheet name",
	)
	prefix = flag.String(
		"prefix",
		"",
		"output prefix: output prefix.sheet.json",
	)
	key = flag.String(
		"key",
		"",
		"column name as key of rows",
	)
	sep = flag.String(
		"sep",
		"\n",
		"sep of merge rows",
	)
	aes = flag.Bool(
		"aes",
		false,
		"if aes encode",
	)
	codeKey = flag.String(
		"codeKey",
		"c3d112d6a47a0a04aad2b9d2d2cad266",
		"codeKey for aes",
	)
)

var err error

func main() {
	flag.Parse()
	if *xlsx == "" {
		flag.Usage()
		fmt.Print("-xlsx as input is required")
		os.Exit(0)
	}
	if *prefix == "" {
		*prefix = *xlsx
	}

	if *sheet == "" {
		xlsxFh, err := excelize.OpenFile(*xlsx)
		simple_util.CheckErr(err)
		for _, sheetName := range xlsxFh.GetSheetMap() {
			fileName := *prefix + "." + sheetName + ".json"
			rows := xlsxFh.GetRows(sheetName)
			var d []byte
			if *key == "" {
				_, data := simple_util.Slice2MapArray(rows)
				d, err = json.MarshalIndent(data, "", "  ")
				simple_util.CheckErr(err)
			} else {
				_, data := simple_util.Slice2MapMapMerge(rows, *key, *sep)
				d, err = json.MarshalIndent(data, "", "  ")
				simple_util.CheckErr(err)
			}
			if *aes {
				simple_util.Encode2file(fileName+".aes", d, []byte(*codeKey))
			} else {
				err = simple_util.Json2file(d, fileName)
				simple_util.CheckErr(err)
			}
		}
	} else {
		fileName := *prefix + "." + *sheet + ".json"
		var d []byte
		if *key == "" {
			_, data := simple_util.Sheet2MapArray(*xlsx, *sheet)
			d, err = json.MarshalIndent(data, "", "  ")
			simple_util.CheckErr(err)
		} else {
			_, data := simple_util.Sheet2MapMapMerge(*xlsx, *sheet, *key, *sep)
			d, err = json.MarshalIndent(data, "", "  ")
			simple_util.CheckErr(err)
		}
		if *aes {
			simple_util.Encode2file(fileName+".aes", d, []byte(*codeKey))
		} else {
			err = simple_util.Json2file(d, fileName)
			simple_util.CheckErr(err)
		}
	}
}

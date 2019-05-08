package main

import (
	"errors"
	"flag"
	"github.com/liserjrqlxue/simple-util"
	"github.com/tealeg/xlsx"
	"log"
	"os"
	"regexp"
	"strconv"
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
	gender = flag.String(
		"gender",
		"",
		"gender list, comma as sep",
	)
)

func main() {
	flag.Parse()
	if *input == "" {
		flag.Usage()
		os.Exit(0)
	}
	if *output == "" {
		*output = *input + ".fixHemi.xlsx"
	}
	inputXlsx, err := xlsx.OpenFile(*input)
	simple_util.CheckErr(err)
	outputXlsx := xlsx.NewFile()
	sheetMap := inputXlsx.Sheet
	for sheetName, sheet := range sheetMap {
		simple_util.CheckErr(updateSheet(*sheet, outputXlsx, sheetName, *gender))
	}
	simple_util.CheckErr(outputXlsx.Save(*output))
}

func updateSheet(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName, gender string) error {
	outputSheet, err := outputXlsx.AddSheet(sheetName)
	simple_util.CheckErr(err)

	nrow := len(sheet.Rows)
	if nrow < 3 {
		return errors.New("error sheet")
	}

	for i := 0; i < 2; i++ {
		var outputRow = outputSheet.AddRow()
		for _, cell := range sheet.Rows[i].Cells {
			text, _ := cell.FormattedValue()
			outputRow.AddCell().SetString(text)
		}
	}

	var keysList []string
	var outputRow = outputSheet.AddRow()
	for _, cell := range sheet.Rows[2].Cells {
		text, _ := cell.FormattedValue()
		keysList = append(keysList, strings.Split(text, "*(")[0])
		outputRow.AddCell().SetString(text)
	}
	if nrow > 3 {
		for i := 3; i < nrow; i++ {
			var outputRow = outputSheet.AddRow()
			var item = make(map[string]string)
			row := sheet.Rows[i]
			for j, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				item[keysList[j]] = text
			}
			updateHemi(item, gender)
			for _, key := range keysList {
				outputRow.AddCell().SetString(item[key])
			}
		}
	}
	return nil
}

func copySheet(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string) (err error) {
	_, err = outputXlsx.AppendSheet(sheet, sheetName)
	return
}

var (
	isChrX  = regexp.MustCompile(`X`)
	isChrY  = regexp.MustCompile(`Y`)
	isChrXY = regexp.MustCompile(`[XY]`)
	isMale  = regexp.MustCompile(`M`)
	withHom = regexp.MustCompile(`Hom`)
)

func updateHemi(item map[string]string, gender string) {
	zygosityKey := "杂合性"
	chr := item["染色体号"]
	if isChrXY.MatchString(chr) && isMale.MatchString(gender) {
		start, err := strconv.Atoi(item["起始位置"])
		simple_util.CheckErr(err)
		stop, err := strconv.Atoi(item["终止位置"])
		simple_util.CheckErr(err)
		if !inPAR(chr, start, stop) && withHom.MatchString(item[zygosityKey]) {
			zygosity := strings.Split(item[zygosityKey], ";")
			genders := strings.Split(gender, ",")
			if len(genders) <= len(zygosity) {
				for i := range genders {
					if isMale.MatchString(genders[i]) && withHom.MatchString(zygosity[i]) {
						zygosity[i] = strings.Replace(zygosity[i], "Hom", "Hemi", 1)
					}
				}
				item[zygosityKey] = strings.Join(zygosity, ";")
			} else {
				log.Fatalf("conflict gender[%s]and Zygosity[%s]\n", gender, item[zygosityKey])
			}
		}
	}
}

var xparReg = [][]int{
	{60000, 2699520},
	{154931043, 155260560},
}
var yparReg = [][]int{
	{10000, 2649520},
	{59034049, 59363566},
}

func inPAR(chr string, start, end int) bool {
	if isChrX.MatchString(chr) {
		for _, par := range xparReg {
			if start < par[1] && end > par[0] {
				return true
			}
		}
	} else if isChrY.MatchString(chr) {
		for _, par := range yparReg {
			if start < par[1] && end > par[0] {
				return true
			}
		}
	}
	return false
}

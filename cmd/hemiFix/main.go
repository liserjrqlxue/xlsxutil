package main

import (
	"flag"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/simple-util"
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
	inputXlsx, err := excelize.OpenFile(*input)
	simple_util.CheckErr(err)
	outputXlsx := excelize.NewFile()
	for _, sheetName := range inputXlsx.GetSheetList() {
		updateSheet(inputXlsx, outputXlsx, sheetName, *gender)
	}
	simple_util.CheckErr(outputXlsx.SaveAs(*output))
}

func updateSheet(inputXlsx *excelize.File, outputXlsx *excelize.File, sheetName, gender string) {
	rows, err := inputXlsx.GetRows(sheetName)
	simple_util.CheckErr(err)

	nrow := len(rows)
	if nrow < 3 {
		return
	}

	outputXlsx.NewSheet(sheetName)

	for i := 0; i < 2; i++ {
		for j, cell := range rows[i] {
			axis, _ := excelize.CoordinatesToCellName(j+1, i+1)
			outputXlsx.SetCellValue(sheetName, axis, cell)
		}
	}

	var keysList []string
	for j, cell := range rows[2] {
		text := strings.Split(cell, "*(")[0]
		keysList = append(keysList, text)
		axis, _ := excelize.CoordinatesToCellName(j+1, 3)
		outputXlsx.SetCellValue(sheetName, axis, cell)
	}

	if nrow > 3 {
		for i := 3; i < nrow; i++ {
			var item = make(map[string]string)
			row := rows[i]
			for j, cell := range row {
				if j < len(keysList) {
					item[keysList[j]] = cell
				}
			}
			updateHemi(item, gender)
			for j, key := range keysList {
				axis, _ := excelize.CoordinatesToCellName(j+1, i+1)
				outputXlsx.SetCellValue(sheetName, axis, item[key])
			}
		}
	}
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
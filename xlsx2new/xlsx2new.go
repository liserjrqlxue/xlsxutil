package main

import (
	"bufio"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

var (
	ex, _  = os.Executable()
	exPath = filepath.Dir(ex)
)

var (
	help = flag.Bool(
		"help",
		false,
		"this help",
	)
	inputExcel = flag.String(
		"input",
		"",
		"input Excel",
	)
	outputExcel = flag.String(
		"output",
		"",
		"output Excel",
	)
	acmgExcel = flag.String(
		"acmg",
		exPath+string(os.PathSeparator)+"崔淑歌 文献 ACMG推荐59个基因更新-20181030.xlsx",
		"database of ACMG",
	)
	acmgSheet = flag.String(
		"acmgSheet",
		"ACMG推荐59个基因",
		"sheet name of ACMG database in excel",
	)
	geneDbExcel = flag.String(
		"geneDb",
		exPath+string(os.PathSeparator)+"基因库0906（最终版）.xlsx",
		"database of 突变频谱",
	)
	geneDbSheet = flag.String(
		"geneDbSheet",
		"基因-疾病（隐藏线粒体基因组）",
		"sheet name of 突变频谱 database in excel",
	)
	titleTxt = flag.String(
		"title",
		exPath+string(os.PathSeparator)+"etc"+string(os.PathSeparator)+"title.txt",
		"output title list",
	)
)

var long2short = map[string]string{
	"Pathogenic":             "P",
	"Likely Pathogenic":      "LP",
	"Uncertain Significance": "VUS",
	"Likely Benign":          "LB",
	"Benign":                 "B",
}

var LoF = map[string]int{
	"splice-3":   1,
	"splice-5":   1,
	"inti-loss":  1,
	"alt-start":  1,
	"frameshift": 1,
	"nonsense":   1,
	"stop-gain":  1,
	"span":       1,
}

var geneDbHash = map[string]string{
	"OMIM":                  "Phenotype MIM number",
	"DiseaseNameEN":         "Disease NameEN",
	"DiseaseNameCH":         "Disease NameCH",
	"AliasEN":               "Alternative Disease NameEN",
	"Location":              "Location",
	"Gene":                  "Gene/Locus",
	"Gene/Locus MIM number": "Gene/Locus MIM number",
	"ModeInheritance":       "Inheritance",
	"GeneralizationEN":      "GeneralizationEN",
	"GeneralizationCH":      "GeneralizationCH",
	//"SystemSort":"SystemSort",
}

var (
	isHgmd    = regexp.MustCompile("DM")
	isClinvar = regexp.MustCompile("Pathogenic|Likely_pathogenic")
	indexReg  = regexp.MustCompile(`\d+\.\s+`)
)

//var leftBracket = regexp.MustCompile("(")
var geneDb = make(map[string]string)

var acmgDb = make(map[string]map[string]string)

func main() {
	flag.Parse()
	if *help || *outputExcel == "" {
		flag.Usage()
		os.Exit(0)
	}

	inputXlsx, err := xlsx.OpenFile(*inputExcel)
	checkError(err)

	// 读取输出title list
	titleList := tsv2array(*titleTxt)

	// ACMG推荐基因数据库
	acmgDbXlsx, err := excelize.OpenFile(*acmgExcel)
	checkError(err)
	acmgDbRows := acmgDbXlsx.GetRows(*acmgSheet)
	var acmgDbTitle []string
	for i, row := range acmgDbRows {
		if i == 0 {
			acmgDbTitle = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				dataHash[acmgDbTitle[j]] = cell
			}
			acmgDb[dataHash["Gene/Locus"]] = dataHash
		}
	}
	//acmgDb2:=sheet2mapHash(acmgDbXlsx,*acmgSheet,"Gene/Locus")
	//fmt.Println(reflect.DeepEqual(acmgDb2,acmgDb))

	// 突变频谱数据库
	geneDbXlsx, err := excelize.OpenFile(*geneDbExcel)
	checkError(err)
	geneDbRows := geneDbXlsx.GetRows(*geneDbSheet)
	var geneDbTitle []string

	for i, row := range geneDbRows {
		if i == 0 {
			geneDbTitle = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				dataHash[geneDbTitle[j]] = cell
			}
			geneDb[dataHash["基因名称"]] = dataHash["突变类型"]
		}
	}

	// 生成新excel
	outputXlsx := xlsx.NewFile()

	// 复制工作表
	// 遍历input.xlsx的工作表
	sheetMap := inputXlsx.Sheet
	for sheetName, sheet := range sheetMap {
		fmt.Printf("Copy sheet [%s]\n", sheetName)
		if sheetName == "filter_variants" {
			annoSheet3(*sheet, outputXlsx, sheetName, titleList)
		} else {
			copySheet4(*sheet, outputXlsx, sheetName)
		}
	}
	// 保存到 outputExcel
	err = outputXlsx.Save(*outputExcel)
	checkError(err)
}

func tsv2array(tsv string) []string {
	file, err := os.Open(tsv)
	checkError(err)
	defer file.Close()
	var array []string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		array = append(array, scanner.Text())
	}
	checkError(scanner.Err())
	return array
}

func sheet2mapArray(excel *excelize.File, sheetName string) []map[string]string {
	rows := excel.GetRows(sheetName)
	var mapArray []map[string]string
	var title []string
	for i, row := range rows {
		if i == 0 {
			title = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				dataHash[title[j]] = cell
			}
			mapArray = append(mapArray, dataHash)
		}
	}
	return mapArray
}

func sheet2mapHash(excel *excelize.File, sheetName, key string) map[string]map[string]string {
	rows := excel.GetRows(sheetName)
	var mapHash map[string]map[string]string
	var title []string
	for i, row := range rows {
		if i == 0 {
			title = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				dataHash[title[j]] = cell
			}
			mapHash[dataHash[key]] = dataHash
		}
	}
	return mapHash
}

func annoSheet(inputXlsx, outputXlsx *excelize.File, sheetName string, titleList []string) error {
	inputRows := inputXlsx.GetRows(sheetName)
	outputXlsx.NewSheet(sheetName)
	var keysList []string
	for i, row := range inputRows {
		if i == 0 {
			for _, cell := range row {
				keysList = append(keysList, cell)
			}
			for j, title := range titleList {
				axis := positionToAxis(i, j)
				outputXlsx.SetCellValue(sheetName, axis, title)
			}
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				dataHash[keysList[j]] = cell
			}
			// pHGVS= pHGVS1+"|"+pHGVS3
			dataHash["pHGVS"] = dataHash["pHGVS1"] + " | " + dataHash["pHGVS3"]

			score, err := strconv.ParseFloat(dataHash["dbscSNV_ADA_SCORE"], 32)
			if err != nil {
				dataHash["dbscSNV_ADA_pred"] = dataHash["dbscSNV_ADA_SCORE"]
			} else {
				if score >= 0.6 {
					dataHash["dbscSNV_ADA_pred"] = "D"
				} else {
					dataHash["dbscSNV_ADA_pred"] = "P"
				}
			}
			score, err = strconv.ParseFloat(dataHash["dbscSNV_RF_SCORE"], 32)
			if err != nil {
				dataHash["dbscSNV_RF_pred"] = dataHash["dbscSNV_RF_SCORE"]
			} else {
				if score >= 0.6 {
					dataHash["dbscSNV_RF_pred"] = "D"
				} else {
					dataHash["dbscSNV_RF_pred"] = "P"
				}
			}

			score, err = strconv.ParseFloat(dataHash["GERP++_RS"], 32)
			if err != nil {
				dataHash["GERP++_RS_pred"] = dataHash["GERP++_RS"]
			} else {
				if score >= 2 {
					dataHash["GERP++_RS_pred"] = "D"
				} else {
					dataHash["GERP++_RS_pred"] = "P"
				}
			}

			dataHash["烈性突变"] = "N"
			if LoF[dataHash["Function"]] == 1 {
				dataHash["烈性突变"] = "Y"
			}

			for j, title := range titleList {
				axis := positionToAxis(i, j)
				outputXlsx.SetCellValue(sheetName, axis, dataHash[title])
			}
		}
	}
	return nil
}

func annoSheet2(sheet *xlsx.Sheet, outputXlsx *excelize.File, sheetName string, titleList []string) error {
	outputXlsx.NewSheet(sheetName)
	var keysList []string
	for i, row := range sheet.Rows {
		if i == 0 {
			for _, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				keysList = append(keysList, text)
			}
			for j, title := range titleList {
				axis := positionToAxis(i, j)
				outputXlsx.SetCellValue(sheetName, axis, title)
			}
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				dataHash[keysList[j]] = text
			}
			// pHGVS= pHGVS1+"|"+pHGVS3
			dataHash["pHGVS"] = dataHash["pHGVS1"] + " | " + dataHash["pHGVS3"]

			score, err := strconv.ParseFloat(dataHash["dbscSNV_ADA_SCORE"], 32)
			if err != nil {
				dataHash["dbscSNV_ADA_pred"] = dataHash["dbscSNV_ADA_SCORE"]
			} else {
				if score >= 0.6 {
					dataHash["dbscSNV_ADA_pred"] = "D"
				} else {
					dataHash["dbscSNV_ADA_pred"] = "P"
				}
			}
			score, err = strconv.ParseFloat(dataHash["dbscSNV_RF_SCORE"], 32)
			if err != nil {
				dataHash["dbscSNV_RF_pred"] = dataHash["dbscSNV_RF_SCORE"]
			} else {
				if score >= 0.6 {
					dataHash["dbscSNV_RF_pred"] = "D"
				} else {
					dataHash["dbscSNV_RF_pred"] = "P"
				}
			}

			score, err = strconv.ParseFloat(dataHash["GERP++_RS"], 32)
			if err != nil {
				dataHash["GERP++_RS_pred"] = dataHash["GERP++_RS"]
			} else {
				if score >= 2 {
					dataHash["GERP++_RS_pred"] = "D"
				} else {
					dataHash["GERP++_RS_pred"] = "P"
				}
			}

			dataHash["烈性突变"] = "N"
			if LoF[dataHash["Function"]] == 1 {
				dataHash["烈性突变"] = "Y"
			}

			dataHash["HGMDorClinvar"] = "N"
			if isHgmd.MatchString(dataHash["HGMD Pred"]) {
				dataHash["HGMDorClinvar"] = "Y"
			}
			if isClinvar.MatchString(dataHash["ClinVar Significance"]) {
				dataHash["HGMDorClinvar"] = "Y"
			}

			dataHash["纯合，半合"] = dataHash["GnomAD homo"] + "|" + dataHash["GnomAD hemi"]

			for j, title := range titleList {
				axis := positionToAxis(i, j)
				outputXlsx.SetCellValue(sheetName, axis, dataHash[title])
			}
		}
	}
	return nil
}

func annoSheet3(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string, titleList []string) error {
	outputSheet, err := outputXlsx.AddSheet(sheetName)
	checkError(err)
	var keysList []string
	for i, row := range sheet.Rows {
		var outputRow = outputSheet.AddRow()
		if i == 0 {
			for _, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				keysList = append(keysList, strings.Split(text, "(")[0])
			}
			for _, title := range titleList {
				//axis := positionToAxis(i, j)
				outputCell := outputRow.AddCell()
				outputCell.SetString(title)
			}
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				dataHash[keysList[j]] = text
			}

			geneSymbol := dataHash["Gene Symbol"]

			// pHGVS= pHGVS1+"|"+pHGVS3
			dataHash["pHGVS"] = dataHash["pHGVS1"] + " | " + dataHash["pHGVS3"]

			score, err := strconv.ParseFloat(dataHash["dbscSNV_ADA_SCORE"], 32)
			if err != nil {
				dataHash["dbscSNV_ADA_pred"] = dataHash["dbscSNV_ADA_SCORE"]
			} else {
				if score >= 0.6 {
					dataHash["dbscSNV_ADA_pred"] = "D"
				} else {
					dataHash["dbscSNV_ADA_pred"] = "P"
				}
			}
			score, err = strconv.ParseFloat(dataHash["dbscSNV_RF_SCORE"], 32)
			if err != nil {
				dataHash["dbscSNV_RF_pred"] = dataHash["dbscSNV_RF_SCORE"]
			} else {
				if score >= 0.6 {
					dataHash["dbscSNV_RF_pred"] = "D"
				} else {
					dataHash["dbscSNV_RF_pred"] = "P"
				}
			}

			score, err = strconv.ParseFloat(dataHash["GERP++_RS"], 32)
			if err != nil {
				dataHash["GERP++_RS_pred"] = dataHash["GERP++_RS"]
			} else {
				if score >= 2 {
					dataHash["GERP++_RS_pred"] = "D"
				} else {
					dataHash["GERP++_RS_pred"] = "P"
				}
			}

			dataHash["烈性突变"] = "否"
			if LoF[dataHash["Function"]] == 1 {
				dataHash["烈性突变"] = "是"
			}

			dataHash["HGMDorClinvar"] = "否"
			if isHgmd.MatchString(dataHash["HGMD Pred"]) {
				dataHash["HGMDorClinvar"] = "是"
			}
			if isClinvar.MatchString(dataHash["ClinVar Significance"]) {
				dataHash["HGMDorClinvar"] = "是"
			}

			dataHash["GnomAD homo"] = dataHash["GnomAD HomoAlt Count"]
			dataHash["GnomAD hemi"] = dataHash["GnomAD HemiAlt Count"]
			dataHash["纯合，半合"] = dataHash["GnomAD HomoAlt Count"] + "|" + dataHash["GnomAD HemiAlt Count"]

			dataHash["突变频谱"] = geneDb[geneSymbol]

			dataHash["历史样本检出个数"] = dataHash["sampleMut"] + "/" + dataHash["sampleAll"]

			// remove index
			for _, k := range [2]string{"GeneralizationEN", "GeneralizationCH"} {
				sep := "\n\n"
				keys := strings.Split(dataHash[k], sep)
				for i, _ := range keys {
					keys[i] = indexReg.ReplaceAllLiteralString(keys[i], "")
				}
				dataHash[k] = strings.Join(keys, sep)
			}
			// add acmg
			if acmgDb[geneSymbol] != nil {
				acmgDbGene := acmgDb[geneSymbol]
				if dataHash["Gene"] == "." {
					for k, v := range geneDbHash {
						dataHash[k] = acmgDbGene[v]
					}
					dataHash["SystemSort"] = "ACMG"
				} else {
					for k, v := range geneDbHash {
						if k == "GeneralizationEN" || k == "GeneralizationCH" {
							sep := "\n\n"
							dataHash[k] = dataHash[k] + sep + acmgDbGene[v]
						} else {
							sep := "\n"
							dataHash[k] = dataHash[k] + sep + acmgDbGene[v]
						}
					}
					sep := "\n"
					dataHash["SystemSort"] = dataHash["SystemSort"] + sep + "ACMG"
				}
			}

			// 自动化判断
			dataHash["自动化判断"] = long2short[dataHash["ACMG"]]

			for _, title := range titleList {
				outputCell := outputRow.AddCell()
				outputCell.SetString(dataHash[title])
			}
		}
	}
	return nil
}

func copySheet(inputXlsx, outputXlsx *excelize.File, sheetName string) error {
	inputRows := inputXlsx.GetRows(sheetName)
	outputXlsx.NewSheet(sheetName)
	for i, row := range inputRows {
		for j, cell := range row {
			axis := positionToAxis(i, j)
			outputXlsx.SetCellValue(sheetName, axis, cell)
		}
	}
	return nil
}

func copySheet2(sheet *xlsx.Sheet, outputXlsx *excelize.File, sheetName string) error {
	outputXlsx.NewSheet(sheetName)
	for i, row := range sheet.Rows {
		for j, cell := range row.Cells {
			text, _ := cell.FormattedValue()
			axis := positionToAxis(i, j)
			outputXlsx.SetCellValue(sheetName, axis, text)
		}
	}
	return nil
}

func copySheet3(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string) error {
	_, err := outputXlsx.AppendSheet(sheet, sheetName)
	checkError(err)
	return err
}

func copySheet4(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string) error {
	outputSheet, err := outputXlsx.AddSheet(sheetName)
	checkError(err)
	for _, row := range sheet.Rows {
		var outputRow = outputSheet.AddRow()
		for _, cell := range row.Cells {
			text, _ := cell.FormattedValue()
			outputCell := outputRow.AddCell()
			outputCell.SetString(text)
		}
	}
	return err
}

func checkError(e error) {
	if e != nil {
		panic(e)
	}
}

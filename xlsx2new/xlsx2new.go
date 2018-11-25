package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tealeg/xlsx"
	"os"
	"regexp"
	"strconv"
)

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

var isHgmd, e1 = regexp.Compile("DM")
var isClinvar, e2 = regexp.Compile("Pathogenic|Likely_pathogenic")
var geneDb = make(map[string]string)

func main() {
	checkError(e1)
	checkError(e2)
	var inputXlsxPath, outputPrefix string
	if len(os.Args) > 2 {
		inputXlsxPath = os.Args[1]
		outputPrefix = os.Args[2]
	} else {
		fmt.Println(os.Args[0], "input.xlsx", "outputPrefix")
		os.Exit(1)
	}

	// 读取输出title list
	/*  中文乱码问题待解决，用title.xlsx替代
	file,err:=os.Open("title.txt")
	checkError(err)
	defer file.Close()

	var titleList []string
	scanner:=bufio.NewScanner(file)
	for scanner.Scan(){
		fmt.Println(scanner.Text())
		titleList=append(titleList,scanner.Text())
	}
	checkError(scanner.Err())
	*/
	titleXlsx, err := excelize.OpenFile("title.xlsx")
	checkError(err)
	titleList := titleXlsx.GetRows("title")[0]

	// 突变频谱数据库
	geneDbXlsx, err := excelize.OpenFile("基因库0906（最终版）.xlsx")
	checkError(err)
	geneDbRows := geneDbXlsx.GetRows("基因-疾病（隐藏线粒体基因组）")
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

	// 读取input.xlsx
	inputXlsx, err := xlsx.OpenFile(inputXlsxPath)
	checkError(err)

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
			copySheet3(*sheet, outputXlsx, sheetName)
		}
	}
	// 保存到 outputPrefix.xlsx
	//outputXlsx.DeleteSheet("Sheet1")
	//outputXlsx.SetActiveSheet(1)
	err = outputXlsx.Save(outputPrefix + ".xlsx")
	checkError(err)
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
				keysList = append(keysList, text)
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

			dataHash["突变频谱"] = geneDb[dataHash["Gene Symbol"]]

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

func checkError(e error) {
	if e != nil {
		panic(e)
	}
}

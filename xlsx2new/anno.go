package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/liserjrqlxue/annogo/GnomAD"
	"github.com/liserjrqlxue/simple-util"
	"github.com/tealeg/xlsx"
	"regexp"
	"strconv"
	"strings"
)

var (
	chrPrefix = regexp.MustCompile("^chr")
)

func getJson(url string, target interface{}) error {
	r, err := myClient.Get(url)
	simple_util.CheckErr(err)
	defer simple_util.DeferClose(r.Body)

	return json.NewDecoder(r.Body).Decode(target)
}

// anno cnv
func annoExonCnv(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string, anno bool) error {
	outputSheet, err := outputXlsx.AddSheet(sheetName)
	simple_util.CheckErr(err)
	var keysList []string
	var keysHash = make(map[string]bool)
	for i, row := range sheet.Rows {
		var outputRow = outputSheet.AddRow()
		if i == 0 {
			for _, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				key := strings.Split(text, "(")[0]
				keysList = append(keysList, key)
				keysHash[key] = true
			}
			// change CopyNum to Copy_Num
			if !keysHash["Copy_Num"] && keysHash["CopyNum"] {
				keysHash["Copy_Num"] = true
				for i, title := range keysList {
					if title == "CopyNum" {
						keysList[i] = "Copy_Num"
					}
				}
			}
			for _, title := range exonCnvAdd {
				if !keysHash[title] {
					keysList = append(keysList, title)
				}
			}
			for _, title := range keysList {
				outputCell := outputRow.AddCell()
				outputCell.SetString(title)
			}
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row.Cells {
				text, _ := cell.FormattedValue()
				dataHash[keysList[j]] = text
			}

			if anno {
				dataHash = updateExonCnv(dataHash)
			}

			for _, title := range keysList {
				outputCell := outputRow.AddCell()
				outputCell.SetString(dataHash[title])
			}
		}
	}
	return nil
}

func updateExonCnv(dataHash map[string]string) map[string]string {
	geneArray := strings.Split(dataHash["OMIM_Gene"], ";")
	dataHash["OMIM_Gene"] = strings.Join(geneArray, "\n")
	var diseaseInfo []string
	var mergeInfo = make(map[string][]string)
	for _, gene := range geneArray {
		if gene == "" || gene == "-" || gene == "." {
			continue
		}
		err := getJson(host+"/OMIM_CN?query="+gene, &diseaseInfo)
		simple_util.CheckErr(err)
		//fmt.Println(len(diseaseInfo))
		if len(diseaseInfo) == 11 {
			for i, k := range exonCnvAdd {
				var sep = "\n"
				if k == "GeneralizationEN" || k == "GeneralizationCH" {
					sep = "\n\n"
				}
				//fmt.Println("["+k+"]\t["+sep+"]")
				mergeInfo[k] = append(mergeInfo[k], strings.Join(strings.Split(diseaseInfo[i], "\n"), sep))
			}
		}
	}
	for _, k := range exonCnvAdd {
		var sep = "\n"
		if k == "GeneralizationEN" || k == "GeneralizationCH" {
			sep = "\n\n"
		}
		dataHash[k] = strings.Join(mergeInfo[k], sep)
	}
	return dataHash
}

type empty interface{}

// anno snv
func annoSheet3(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string, titleList []string) error {

	outputSheet, err := outputXlsx.AddSheet(sheetName)
	simple_util.CheckErr(err)

	nrow := len(sheet.Rows)
	if nrow < 1 {
		return errors.New("empty sheet!")
	}

	// title
	var keysList []string
	var outputRow = outputSheet.AddRow()
	for _, cell := range sheet.Rows[0].Cells {
		text, _ := cell.FormattedValue()
		keysList = append(keysList, strings.Split(text, "(")[0])
	}
	for _, title := range titleList {
		//axis := positionToAxis(i, j)
		outputCell := outputRow.AddCell()
		outputCell.SetString(title)
	}

	if nrow > 1 {
		sem := make(chan empty, nrow-1)
		var dataHashArray = make([]map[string]string, nrow-1)
		for i := 1; i < nrow; i++ {
			go func(i int) {
				var dataHash = make(map[string]string)
				row := sheet.Rows[i]
				for j, cell := range row.Cells {
					text, _ := cell.FormattedValue()
					dataHash[keysList[j]] = text
				}
				if *annoGnomAD {
					dataHash = addGnomAD(tbx, dataHash)
				}
				dataHash = updateSnv(dataHash)
				dataHashArray[i-1] = dataHash
				sem <- new(empty)
			}(i)
		}
		for i := 0; i < nrow-1; i++ {
			<-sem
		}
		for i := 0; i < nrow-1; i++ {
			var outputRow = outputSheet.AddRow()
			dataHash := dataHashArray[i]
			for _, title := range titleList {
				outputCell := outputRow.AddCell()
				outputCell.SetString(dataHash[title])
			}
		}
	}
	fmt.Printf("anno %d count\n", nrow)
	return nil
}

func updateSnv(dataHash map[string]string) map[string]string {
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

	// 0-0.6 不保守  0.6-2.5 保守 ＞2.5 高度保守
	score, err = strconv.ParseFloat(dataHash["PhyloP Vertebrates"], 32)
	if err != nil {
		dataHash["PhyloP Vertebrates Pred"] = dataHash["PhyloP Vertebrates"]
	} else {
		if score >= 2.5 {
			dataHash["PhyloP Vertebrates Pred"] = "高度保守"
		} else if score > 0.6 {
			dataHash["PhyloP Vertebrates Pred"] = "保守"
		} else {
			dataHash["PhyloP Vertebrates Pred"] = "不保守"
		}
	}
	score, err = strconv.ParseFloat(dataHash["PhyloP Placental Mammals"], 32)
	if err != nil {
		dataHash["PhyloP Placental Mammals Pred"] = dataHash["PhyloP Placental Mammals"]
	} else {
		if score >= 2.5 {
			dataHash["PhyloP Placental Mammals Pred"] = "高度保守"
		} else if score > 0.6 {
			dataHash["PhyloP Placental Mammals Pred"] = "保守"
		} else {
			dataHash["PhyloP Placental Mammals Pred"] = "不保守"
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
	dataHash["纯合，半合"] = dataHash["GnomAD HomoAlt Count"] // + "|" + dataHash["GnomAD HemiAlt Count"]
	dataHash["MutationNameLite"] = dataHash["Transcript"] + ":" + strings.Split(dataHash["MutationName"], ":")[1]

	dataHash["突变频谱"] = geneDb[geneSymbol]

	dataHash["历史样本检出个数"] = dataHash["sampleMut"] + "/" + dataHash["sampleAll"]

	// remove index
	for _, k := range [2]string{"GeneralizationEN", "GeneralizationCH"} {
		sep := "\n\n"
		keys := strings.Split(dataHash[k], sep)
		for i := range keys {
			keys[i] = indexReg.ReplaceAllLiteralString(keys[i], "")
		}
		dataHash[k] = strings.Join(keys, sep)
	}

	// add acmg
	if acmgDb[geneSymbol] != nil {
		acmgDbGene := acmgDb[geneSymbol]
		var sep = "\n"
		systemSort := strings.Split(acmgDbGene["SystemSort"], sep)
		for i := range systemSort {
			systemSort[i] = "ACMG"
		}

		if dataHash["Gene"] == "." {
			for k, v := range geneDbHash {
				dataHash[k] = acmgDbGene[v]
			}
		} else {
			for k, v := range geneDbHash {
				sep = "\n"
				vArray := strings.Split(acmgDbGene[v], sep)
				if k == "GeneralizationEN" || k == "GeneralizationCH" {
					sep = "\n\n"
				}
				for _, str := range strings.Split(dataHash[k], sep) {
					vArray = append(vArray, str)
				}
				dataHash[k] = strings.Join(vArray, sep)
			}
			sep = "\n"
			for _, str := range strings.Split(dataHash["SystemSort"], sep) {
				systemSort = append(systemSort, str)
			}
		}
		dataHash["SystemSort"] = strings.Join(systemSort, sep)
	}

	// 自动化判断
	dataHash = addACMG2015(dataHash)
	dataHash["自动化判断"] = long2short[dataHash["ACMG"]]
	return dataHash
}

// copy sheet without change
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
	simple_util.CheckErr(err)
	return err
}
func copySheet4(sheet xlsx.Sheet, outputXlsx *xlsx.File, sheetName string) error {
	outputSheet, err := outputXlsx.AddSheet(sheetName)
	simple_util.CheckErr(err)
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
	var mapHash = make(map[string]map[string]string)
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

func excel2MapMap(excelPath, sheetName, key string) map[string]map[string]string {
	inXlsx, err := excelize.OpenFile(excelPath)
	simple_util.CheckErr(err)
	var db = make(map[string]map[string]string)
	rows := inXlsx.GetRows(sheetName)
	var title []string
	for i, row := range rows {
		if i == 0 {
			title = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				dataHash[title[j]] = cell
			}
			db[dataHash[key]] = dataHash
		}
	}
	return db
}

func addGnomAD(tbx *GnomAD.Tbx, inputData map[string]string) map[string]string {
	chr := inputData["#Chr"]
	chr = chrPrefix.ReplaceAllString(chr, "")
	start, err := strconv.Atoi(inputData["Start"])
	simple_util.CheckErr(err)
	stop, err := strconv.Atoi(inputData["Stop"])
	qStart := start
	if start == stop {
		qStart -= 1
	}
	vals := tbx.Query(chr, start-1, stop)
	if vals == nil {
		return inputData
	}

	ref := inputData["Ref"]
	if ref == "." {
		ref = ""
	}
	alt := inputData["Call"]
	if alt == "." {
		alt = ""
	}

	hit := tbx.Hit(chr, start, stop, ref, alt, vals)
	if hit.Info == nil {
		return inputData
	}
	if hit.Info["AF"] == nil {
		inputData["GnomAD AF"] = ""
	} else {
		inputData["GnomAD AF"] = strconv.FormatFloat(float64(hit.Info["AF"].(float32)), 'f', -1, 32)
	}
	if hit.Info["AF_eas"] == nil {
		inputData["GnomAD EAS AF"] = ""
	} else {
		inputData["GnomAD EAS AF"] = strconv.FormatFloat(float64(hit.Info["AF_eas"].(float32)), 'f', -1, 32)
	}
	inputData["GnomAD HomoAlt Count"] = strconv.Itoa(hit.Info["nhomalt"].(int))
	inputData["GnomAD EAS HomoAlt Count"] = strconv.Itoa(hit.Info["nhomalt"].(int))
	return inputData
	//fmt.Println(chr,"\t",start,"\t",stop,"\t",ref,"\t",alt,"\t:\t",hit.Info["AF"],hit.Info["AF_eas"])
	//fmt.Println(hit)//.Chrom,hit.Start,hit.End,hit.Ref,hit.Alt)
}

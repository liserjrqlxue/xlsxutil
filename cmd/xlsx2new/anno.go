package main

import (
	"encoding/json"
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/liserjrqlxue/acmg2015"
	"github.com/liserjrqlxue/anno2xlsx/v2/anno"
	"github.com/liserjrqlxue/annogo/GnomAD"
	"github.com/liserjrqlxue/simple-util"
	"github.com/xuri/excelize/v2"
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

func annoExonCnv(inputXlsx *excelize.File, outputXlsx *excelize.File, sheetName string, annoFlag bool) {
	rows, err := inputXlsx.GetRows(sheetName)
	simple_util.CheckErr(err)

	nrow := len(rows)
	if nrow < 1 {
		return
	}

	outputXlsx.NewSheet(sheetName)
	var keysList []string
	var keysHash = make(map[string]bool)

	for j, cell := range rows[0] {
		key := strings.Split(cell, "(")[0]
		keysList = append(keysList, key)
		keysHash[key] = true
		axis, _ := excelize.CoordinatesToCellName(j+1, 1)
		outputXlsx.SetCellValue(sheetName, axis, cell)
	}

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
			axis, _ := excelize.CoordinatesToCellName(len(keysList), 1)
			outputXlsx.SetCellValue(sheetName, axis, title)
		}
	}

	if nrow > 1 {
		for i := 1; i < nrow; i++ {
			var dataHash = make(map[string]string)
			row := rows[i]
			for j, cell := range row {
				if j < len(keysList) {
					dataHash[keysList[j]] = cell
				}
			}

			if annoFlag {
				dataHash = updateExonCnv(dataHash)
			}

			for j, title := range keysList {
				axis, _ := excelize.CoordinatesToCellName(j+1, i+1)
				outputXlsx.SetCellValue(sheetName, axis, dataHash[title])
			}
		}
	}
}

func updateExonCnv(dataHash map[string]string) map[string]string {
	dataHash["OMIM_Gene"] = newlineReg.ReplaceAllLiteralString(dataHash["OMIM_Gene"], ";")
	geneArray := strings.Split(dataHash["OMIM_Gene"], ";")
	dataHash["OMIM_Gene"] = strings.Join(geneArray, "\n")
	var diseaseInfo interface{}
	var mergeInfo = make(map[string][]string)
	for _, gene := range geneArray {
		if gene == "" || gene == "-" || gene == "." {
			continue
		}
		err := getJson(host+"/OMIM_CN?query="+gene, &diseaseInfo)
		simple_util.CheckErr(err)
		disInfo, ok := diseaseInfo.([]interface{})
		if ok && len(disInfo) == 11 {
			for i, k := range exonCnvAdd {
				var sep = "\n"
				if k == "GeneralizationEN" || k == "GeneralizationCH" {
					sep = "\n\n"
				}
				mergeInfo[k] = append(mergeInfo[k], strings.Join(strings.Split(disInfo[i].(string), "\n"), sep))
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

func annoSheet3(inputXlsx *excelize.File, outputXlsx *excelize.File, sheetName, gender string, titleList []string) error {
	rows, err := inputXlsx.GetRows(sheetName)
	simple_util.CheckErr(err)

	nrow := len(rows)
	if nrow < 1 {
		return fmt.Errorf("empty sheet")
	}

	outputXlsx.NewSheet(sheetName)
	var keysList []string

	for _, cell := range rows[0] {
		text := strings.Split(cell, "(")[0]
		keysList = append(keysList, text)
	}

	for j, title := range titleList {
		axis, _ := excelize.CoordinatesToCellName(j+1, 1)
		outputXlsx.SetCellValue(sheetName, axis, title)
	}

	if nrow > 1 {
		sem := make(chan empty, nrow-1)
		var dataHashArray = make([]map[string]string, nrow-1)
		for i := 1; i < nrow; i++ {
			go func(i int) {
				var dataHash = make(map[string]string)
				row := rows[i]
				for j, cell := range row {
					if j < len(keysList) {
						dataHash[keysList[j]] = cell
					}
				}
				if *annoGnomAD {
					dataHash = addGnomAD(tbx, dataHash)
				}
				geneSymbol := dataHash["Gene Symbol"]
				dataHash["突变频谱"] = geneDb[geneSymbol]

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

				if *annoACMG {
					acmg2015.AddEvidences(dataHash)
					dataHash["ACMG"] = acmg2015.PredACMG2015(dataHash, true)
				}
				anno.UpdateSnv(dataHash, gender)
				dataHashArray[i-1] = dataHash
				sem <- new(empty)
			}(i)
		}
		for i := 0; i < nrow-1; i++ {
			<-sem
		}
		for i := 0; i < nrow-1; i++ {
			dataHash := dataHashArray[i]
			for j, title := range titleList {
				axis, _ := excelize.CoordinatesToCellName(j+1, i+2)
				outputXlsx.SetCellValue(sheetName, axis, dataHash[title])
			}
		}
	}
	fmt.Printf("anno %d count\n", nrow)
	return nil
}

func updateSnv(dataHash map[string]string) map[string]string {
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
	dataHash["纯合，半合"] = dataHash["GnomAD HomoAlt Count"]
	if len(strings.Split(dataHash["MutationName"], ":")) > 1 {
		dataHash["MutationNameLite"] = dataHash["Transcript"] + ":" + strings.Split(dataHash["MutationName"], ":")[1]
	} else {
		dataHash["MutationNameLite"] = dataHash["MutationName"]
	}

	dataHash["历史样本检出个数"] = dataHash["sampleMut"] + "/" + dataHash["sampleAll"]

	for _, k := range [2]string{"GeneralizationEN", "GeneralizationCH"} {
		sep := "\n\n"
		keys := strings.Split(dataHash[k], sep)
		for i := range keys {
			keys[i] = indexReg.ReplaceAllLiteralString(keys[i], "")
		}
		dataHash[k] = strings.Join(keys, sep)
	}

	dataHash["自动化判断"] = long2short[dataHash["ACMG"]]
	return dataHash
}

func copySheet(inputXlsx, outputXlsx *excelize.File, sheetName string) {
	inputRows, err := inputXlsx.GetRows(sheetName)
	simple_util.CheckErr(err)
	outputXlsx.NewSheet(sheetName)
	for i, row := range inputRows {
		for j, cell := range row {
			axis, _ := excelize.CoordinatesToCellName(j+1, i+1)
			outputXlsx.SetCellValue(sheetName, axis, cell)
		}
	}
}

func sheet2mapArray(excel *excelize.File, sheetName string) []map[string]string {
	rows, err := excel.GetRows(sheetName)
	simple_util.CheckErr(err)
	var mapArray []map[string]string
	var title []string
	for i, row := range rows {
		if i == 0 {
			title = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				if j < len(title) {
					dataHash[title[j]] = cell
				}
			}
			mapArray = append(mapArray, dataHash)
		}
	}
	return mapArray
}

func sheet2mapHash(excel *excelize.File, sheetName, key string) map[string]map[string]string {
	rows, err := excel.GetRows(sheetName)
	simple_util.CheckErr(err)
	var mapHash = make(map[string]map[string]string)
	var title []string
	for i, row := range rows {
		if i == 0 {
			title = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				if j < len(title) {
					dataHash[title[j]] = cell
				}
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
	rows, err := inXlsx.GetRows(sheetName)
	simple_util.CheckErr(err)
	var title []string
	for i, row := range rows {
		if i == 0 {
			title = row
		} else {
			var dataHash = make(map[string]string)
			for j, cell := range row {
				if j < len(title) {
					dataHash[title[j]] = cell
				}
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
}
package main

import (
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/liserjrqlxue/annogo/GnomAD"
	"github.com/liserjrqlxue/simple-util"
	"github.com/tealeg/xlsx"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"time"
)

// os
var (
	ex, _  = os.Executable()
	exPath = filepath.Dir(ex)
	pSep   = string(os.PathSeparator)
	dbPath = exPath + pSep + "db" + pSep
)

// flag
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
		exPath+pSep+"崔淑歌 文献 ACMG推荐59个基因更新-20181030.xlsx",
		"database of ACMG",
	)
	acmgSheet = flag.String(
		"acmgSheet",
		"ACMG推荐59个基因",
		"sheet name of ACMG database in excel",
	)
	geneDbExcel = flag.String(
		"geneDb",
		dbPath+"基因库-更新版基因特征谱-加动态突变-20190110.xlsx",
		"database of 突变频谱",
	)
	geneDbSheet = flag.String(
		"geneDbSheet",
		"Sheet1",
		"sheet name of 突变频谱 database in excel",
	)
	titleTxt = flag.String(
		"title",
		exPath+pSep+"etc"+pSep+"title.txt",
		"output title list",
	)
	annoCnv = flag.Bool(
		"annoCnv",
		false,
		"anno exon_cnv sheet with disease of target gene",
	)
	annoGnomAD = flag.Bool(
		"annoGnomAD",
		false,
		"flag to update GnomAD info",
	)
	gnomAD = flag.String(
		"gnomAD",
		dbPath+"gnomad.exomes.r2.1.sites.vcf.gz",
		"gnomAD file path",
	)
	annoACMG = flag.Bool(
		"annoACMG",
		false,
		"flag to update ACMG info",
	)
	gender = flag.String(
		"gender",
		"",
		"gender of proband",
	)
)

var long2short = map[string]string{
	"Pathogenic":             "P",
	"Likely Pathogenic":      "LP",
	"Uncertain Significance": "VUS",
	"Likely Benign":          "LB",
	"Benign":                 "B",
	"P":                      "P",
	"LP":                     "LP",
	"VUS":                    "VUS",
	"LB":                     "LB",
	"B":                      "B",
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

// regexp
var (
	isHgmd     = regexp.MustCompile("DM")
	isClinvar  = regexp.MustCompile("Pathogenic|Likely_pathogenic")
	indexReg   = regexp.MustCompile(`\d+\.\s+`)
	newlineReg = regexp.MustCompile(`\n+`)
)

//var leftBracket = regexp.MustCompile("(")
var geneDb = make(map[string]string)

var acmgDb = make(map[string]map[string]string)

var geneDisease = make(map[string][]string)
var diseaseDb = make(map[string]map[string]string)
var diseaseKey = []string{
	"Phenotype MIM number",
	"Disease NameCH",
	"Alternative Disease NameEN",
	"GeneralizationEN",
	"GeneralizationCH",
	"SystemSort",
}

var myClient = &http.Client{Timeout: 10 * time.Second}
var host = "http://192.168.136.114:9898"
var exonCnvAdd = []string{
	"OMIM",
	"DiseaseNameEN",
	"DiseaseNameCH",
	"AliasEN",
	"Location",
	"Omim Gene",
	"Gene/Locus MIM number",
	"ModeInheritance",
	"GeneralizationEN",
	"GeneralizationCH",
	"SystemSort",
}

type annSheetArgs struct {
	InputSheet  *xlsx.Sheet
	outputExcel *xlsx.File
	sheetName   string
	titleList   []string
	annoInfo    map[string]interface{}
}

var tbx *GnomAD.Tbx

func main() {
	t0 := time.Now()
	flag.Parse()
	if *help || *outputExcel == "" {
		flag.Usage()
		os.Exit(0)
	}

	inputXlsx, err := xlsx.OpenFile(*inputExcel)
	simple_util.CheckErr(err)

	if *annoGnomAD {
		tbx, err = GnomAD.New(*gnomAD)
		simple_util.CheckErr(err)
		defer simple_util.DeferClose(tbx)
	}

	// 读取输出title list
	titleList := simple_util.File2Array(*titleTxt)

	// ACMG推荐基因数据库
	acmgDb = excel2MapMap(*acmgExcel, *acmgSheet, "Gene/Locus")

	// 基因-疾病数据库
	geneDiseaseDbXlsx, err := excelize.OpenFile(*geneDiseaseDbExcel)
	simple_util.CheckErr(err)
	geneDiseaseDb := sheet2mapArray(geneDiseaseDbXlsx, *geneDiseaseSheet)
	for _, db := range geneDiseaseDb {
		gene := db["Gene/Locus"]
		disease := db["Disease NameEN"]
		geneDisease[gene] = append(geneDisease[gene], disease)

	}

	//acmgDb2:=sheet2mapHash(acmgDbXlsx,*acmgSheet,"Gene/Locus")
	//fmt.Println(reflect.DeepEqual(acmgDb2,acmgDb))

	// 突变频谱数据库
	geneDbXlsx, err := excelize.OpenFile(*geneDbExcel)
	simple_util.CheckErr(err)
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
			if geneDb[dataHash["基因名"]] == "" {
				geneDb[dataHash["基因名"]] = dataHash["突变/致病多样性-补充/更正"]
			} else {
				geneDb[dataHash["基因名"]] = geneDb[dataHash["基因名"]] + ";" + dataHash["突变/致病多样性-补充/更正"]
			}
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
			t0 := time.Now()
			err = annoSheet3(*sheet, outputXlsx, sheetName, *gender, titleList)
			t1 := time.Now()
			fmt.Printf("The call took %v to run.\n", t1.Sub(t0))
			simple_util.CheckErr(err)

		} else if sheetName == "exon_cnv" {
			err = annoExonCnv(*sheet, outputXlsx, sheetName, *annoCnv)
			simple_util.CheckErr(err)
		} else {
			err = copySheet4(*sheet, outputXlsx, sheetName)
			simple_util.CheckErr(err)
		}
	}
	t1 := time.Now()
	fmt.Printf("The call took %v to run.\n", t1.Sub(t0))
	// 保存到 outputExcel
	err = outputXlsx.Save(*outputExcel)
	simple_util.CheckErr(err)
	fmt.Printf("The call took %v to run.\n", t1.Sub(t0))
}

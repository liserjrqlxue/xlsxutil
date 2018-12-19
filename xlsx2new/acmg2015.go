package main

func addACMG2015(inputData map[string]string) map[string]string {
	var info map[string]int

	inputData["ACMG"] = predACMG2015(info)
	return inputData
}

func predACMG2015(info map[string]int) string {

	PVS1 := info["PVS1"]

	PS1 := info["PS1"]
	PS2 := info["PS2"]
	PS3 := info["PS3"]
	PS4 := info["PS4"]

	PM1 := info["PM1"]
	PM2 := info["PM2"]
	PM3 := info["PM3"]
	PM4 := info["PM4"]
	PM5 := info["PM5"]
	PM6 := info["PM6"]

	PP1 := info["PP1"]
	PP2 := info["PP2"]
	PP3 := info["PP3"]
	PP4 := info["PP4"]
	PP5 := info["PP5"]

	BA1 := info["BA1"]

	BS1 := info["BS1"]
	BS2 := info["BS2"]
	BS3 := info["BS3"]
	BS4 := info["BS4"]

	BP1 := info["BP1"]
	BP2 := info["BP2"]
	BP3 := info["BP3"]
	BP4 := info["BP4"]
	BP5 := info["BP5"]
	BP6 := info["BP6"]
	BP7 := info["BP7"]

	// PVS
	//  PVS1 5 得分
	//  PVS1 6 不得分
	if PVS1 == 5 {
		PVS1 = 1
	}
	if PVS1 == 6 {
		PVS1 = 1
	}
	sumPVS := PVS1

	// PS
	//  PS1 2 暂时不得分
	//  PS1 3 得分
	//  PS1 4 不得分
	if PS1 == 2 || PS1 == 4 {
		PS1 = 0
	}
	if PS1 == 3 {
		PS1 = 1
	}
	//  PS4 5 得分
	if PS4 == 5 {
		PS4 = 1
	}
	sumPS := PS1 + PS2 + PS3 + PS4

	// PM
	//  不与PS4 同时得分
	if PS4 > 0 {
		PM2 = 0
	}
	//  PM3 2 升级到PS得分
	if PM3 == 2 {
		sumPS++
		PM3 = 0
	}
	//  PM5 2 暂时得分
	//  PM5 3 得分
	//  PM5 4 不得分
	//  PM5 5 得分
	if PM5 == 2 || PM5 == 3 || PM5 == 5 {
		PM5 = 1
	}
	if PM5 == 4 {
		PM5 = 0
	}
	sumPM := PM1 + PM2 + PM3 + PM4 + PM5 + PM6

	// PP
	//  PP1 2 升级到PM
	if PP1 == 2 {
		sumPM++
		PP1 = 0
	}
	//  PS1/PM5  PP5 不共存
	if PS1 > 0 || PM5 > 0 {
		PP5 = 0
	}
	//  ACMG 已取消该证据
	PP5 = 0
	sumPP := PP1 + PP2 + PP3 + PP4 + PP5
	// BA
	sumBA := BA1
	// BS
	sumBS := BS1 + BS2 + BS3 + BS4
	// BP
	sumBP := BP1 + BP2 + BP3 + BP4 + BP5 + BP6 + BP7

	var ACMG = make(map[string]bool)
	if sumPVS > 0 {
		if sumPS == 1 || sumPM > 1 || sumPP > 1 || (sumPM == 1 && sumPP == 1) {
			ACMG["P"] = true
		}
		if sumPM == 1 {
			ACMG["LP"] = true
		}
	}
	if sumPS > 1 {
		ACMG["P"] = true
	}
	if sumPS == 1 {
		if sumPM > 2 || (sumPM == 2 && sumPP >= 2) || (sumPM == 1 && sumPP >= 4) {
			ACMG["P"] = true
		}
		if sumPM == 1 || sumPM == 2 || sumPP > 1 {
			ACMG["LP"] = true
		}
	}
	if sumPM > 2 || (sumPM == 2 && sumPP > 1) || (sumPM == 1 && sumPP > 3) {
		ACMG["LP"] = true
	}
	if sumBA > 0 || sumBS > 1 {
		ACMG["B"] = true
	}
	if sumBP > 1 || (sumBP == 1 && sumBS == 1) {
		ACMG["LB"] = true
	}
	var PLP, BLB bool
	if ACMG["P"] || ACMG["LP"] {
		PLP = true
	}
	if ACMG["B"] || ACMG["LB"] {
		BLB = true
	}
	if PLP && BLB {
		return "VUS"
	} else if ACMG["P"] {
		return "P"
	} else if ACMG["LP"] {
		return "LP"
	} else if ACMG["B"] {
		return "B"
	} else if ACMG["LB"] {
		return "LB"
	} else {
		return "VUS"
	}
}

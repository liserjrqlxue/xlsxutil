package main

import (
	"fmt"
	"strconv"
)

func positionToAxis(row, col int) string {
	if row < 0 || col < 0 {
		return ""
	}
	rowString := strconv.Itoa(row + 1)
	colString := string('A' + col%26)
	//if col>26{
	col /= 26
	for col > 0 {
		col -= 1
		colString = string('A'+col%26) + colString
		col /= 26
	}
	//}

	return colString + rowString
}

func axisToPosition(axis string) (int, int, error) {
	col := 0
	for i, char := range axis {
		if char >= 'A' && char <= 'Z' {
			col *= 26
			col += int(char - 'A' + 1)
		} else if char >= 'a' && char <= 'z' {
			col *= 26
			col += int(char - 'a' + 1)
		} else {
			row, err := strconv.Atoi(axis[i:])
			return row - 1, col - 1, err
		}
	}
	return -1, -1, fmt.Errorf("invalid axis format %s", axis)
}

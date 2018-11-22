package main

import "testing"

func TestPositionToAxis(t *testing.T) {
	cells := []struct {
		row, col int
		axis     string
	}{
		{0, 0, "A1"},
		{1, 17, "R2"},
		{1, 3, "D2"},
		{41, 20, "U42"},
		{18, 26, "AA19"},
		{58, 36, "AK59"},
		{216, 407, "OR217"},
		{122, 702, "AAA123"},
		{5926, 3141, "DPV5927"},
		{3141, 5926, "HSY3142"},
		//{5, -5, ""},
	}

	for _, cell := range cells {
		axis := positionToAxis(cell.row, cell.col)
		if axis != cell.axis {
			t.Errorf("Expected %s, but got %s", cell.axis, axis)
		}
		row, col, err := axisToPosition(cell.axis)
		if err != nil {
			t.Errorf("Unexpected error: %s", err)
		}
		if row != cell.row || col != cell.col {
			t.Errorf("Expected %d %d, but got %d %d", cell.row, cell.col, row, col)
		}
	}
}

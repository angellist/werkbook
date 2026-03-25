package ooxml

import "testing"

func TestExpandSharedFormulas(t *testing.T) {
	sd := SheetData{
		Rows: []RowData{
			{Num: 4, Cells: []CellData{
				{Ref: "F4", Formula: `IF(D4=E4,"✅","❌")`, SharedIndex: -1},
			}},
			{Num: 7, Cells: []CellData{
				// Master: si=0, has formula text and ref
				{Ref: "F7", Formula: `IF(D7=E7,"✅","❌")`, FormulaType: "shared", FormulaRef: "F7:F10", SharedIndex: 0},
			}},
			{Num: 8, Cells: []CellData{
				// Child: si=0, no formula text
				{Ref: "F8", Formula: "", FormulaType: "shared", SharedIndex: 0, Type: "str", Value: "✅"},
			}},
			{Num: 9, Cells: []CellData{
				{Ref: "F9", Formula: "", FormulaType: "shared", SharedIndex: 0, Type: "str", Value: "✅"},
			}},
			{Num: 10, Cells: []CellData{
				{Ref: "F10", Formula: "", FormulaType: "shared", SharedIndex: 0, Type: "str", Value: "❌"},
			}},
		},
	}

	expandSharedFormulas(&sd)

	// F4 should be untouched (not a shared formula).
	if got := sd.Rows[0].Cells[0].Formula; got != `IF(D4=E4,"✅","❌")` {
		t.Errorf("F4 formula = %q, want unchanged", got)
	}

	// F7 (master) should keep its formula and become standalone.
	f7 := sd.Rows[1].Cells[0]
	if f7.Formula != `IF(D7=E7,"✅","❌")` {
		t.Errorf("F7 formula = %q, want unchanged master formula", f7.Formula)
	}
	if f7.FormulaType != "" {
		t.Errorf("F7 FormulaType = %q, want empty (standalone)", f7.FormulaType)
	}

	// F8 should have the formula shifted by +1 row.
	f8 := sd.Rows[2].Cells[0]
	if f8.Formula != `IF(D8=E8,"✅","❌")` {
		t.Errorf("F8 formula = %q, want IF(D8=E8,...)", f8.Formula)
	}
	if f8.FormulaType != "" {
		t.Errorf("F8 FormulaType = %q, want empty", f8.FormulaType)
	}

	// F9 should be shifted by +2 rows.
	f9 := sd.Rows[3].Cells[0]
	if f9.Formula != `IF(D9=E9,"✅","❌")` {
		t.Errorf("F9 formula = %q, want IF(D9=E9,...)", f9.Formula)
	}

	// F10 should be shifted by +3 rows.
	f10 := sd.Rows[4].Cells[0]
	if f10.Formula != `IF(D10=E10,"✅","❌")` {
		t.Errorf("F10 formula = %q, want IF(D10=E10,...)", f10.Formula)
	}
}

func TestExpandSharedFormulasMultipleGroups(t *testing.T) {
	sd := SheetData{
		Rows: []RowData{
			{Num: 1, Cells: []CellData{
				// Group 0 master in column A
				{Ref: "A1", Formula: "B1+C1", FormulaType: "shared", FormulaRef: "A1:A3", SharedIndex: 0},
				// Group 1 master in column D
				{Ref: "D1", Formula: "$A1*2", FormulaType: "shared", FormulaRef: "D1:D3", SharedIndex: 1},
			}},
			{Num: 2, Cells: []CellData{
				{Ref: "A2", Formula: "", FormulaType: "shared", SharedIndex: 0},
				{Ref: "D2", Formula: "", FormulaType: "shared", SharedIndex: 1},
			}},
			{Num: 3, Cells: []CellData{
				{Ref: "A3", Formula: "", FormulaType: "shared", SharedIndex: 0},
				{Ref: "D3", Formula: "", FormulaType: "shared", SharedIndex: 1},
			}},
		},
	}

	expandSharedFormulas(&sd)

	// Group 0: B1+C1 shifted down
	if got := sd.Rows[1].Cells[0].Formula; got != "B2+C2" {
		t.Errorf("A2 formula = %q, want B2+C2", got)
	}
	if got := sd.Rows[2].Cells[0].Formula; got != "B3+C3" {
		t.Errorf("A3 formula = %q, want B3+C3", got)
	}

	// Group 1: $A1*2 — $A is absolute col, row shifts
	if got := sd.Rows[1].Cells[1].Formula; got != "$A2*2" {
		t.Errorf("D2 formula = %q, want $A2*2", got)
	}
	if got := sd.Rows[2].Cells[1].Formula; got != "$A3*2" {
		t.Errorf("D3 formula = %q, want $A3*2", got)
	}
}

func TestShiftFormulaRefs(t *testing.T) {
	tests := []struct {
		name     string
		formula  string
		dCol     int
		dRow     int
		expected string
	}{
		{"simple shift down", "A1+B1", 0, 1, "A2+B2"},
		{"shift right and down", "A1+B1", 1, 1, "B2+C2"},
		{"absolute col preserved", "$A1+B1", 0, 1, "$A2+B2"},
		{"absolute row preserved", "A$1+B$1", 0, 1, "A$1+B$1"},
		{"fully absolute", "$A$1+$B$1", 1, 1, "$A$1+$B$1"},
		{"sheet ref", "Sheet1!A1+B1", 0, 1, "Sheet1!A2+B2"},
		{"string literals unchanged", `IF(A1=B1,"yes","no")`, 0, 1, `IF(A2=B2,"yes","no")`},
		{"no shift", "A1+B1", 0, 0, "A1+B1"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := shiftFormulaRefs(tt.formula, tt.dCol, tt.dRow)
			if got != tt.expected {
				t.Errorf("shiftFormulaRefs(%q, %d, %d) = %q, want %q",
					tt.formula, tt.dCol, tt.dRow, got, tt.expected)
			}
		})
	}
}

func TestCellRefToCoordinates(t *testing.T) {
	tests := []struct {
		ref     string
		col     int
		row     int
		wantErr bool
	}{
		{"A1", 1, 1, false},
		{"F7", 6, 7, false},
		{"Z26", 26, 26, false},
		{"AA1", 27, 1, false},
		{"", 0, 0, true},
		{"A", 0, 0, true},
		{"1", 0, 0, true},
	}

	for _, tt := range tests {
		col, row, err := cellRefToCoordinates(tt.ref)
		if (err != nil) != tt.wantErr {
			t.Errorf("cellRefToCoordinates(%q) error = %v, wantErr %v", tt.ref, err, tt.wantErr)
			continue
		}
		if !tt.wantErr && (col != tt.col || row != tt.row) {
			t.Errorf("cellRefToCoordinates(%q) = (%d,%d), want (%d,%d)", tt.ref, col, row, tt.col, tt.row)
		}
	}
}

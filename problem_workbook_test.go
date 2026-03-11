package werkbook_test

import (
	"testing"

	"github.com/jpoz/werkbook"
)

func TestProblemWorkbookRecalculateMatchesExcel(t *testing.T) {
	f, err := werkbook.Open("problem.xlsx")
	if err != nil {
		t.Fatalf("Open(problem.xlsx): %v", err)
	}

	f.Recalculate()

	summary := f.Sheet("Out - Summary")
	if summary == nil {
		t.Fatal("Out - Summary sheet not found")
	}

	tests := []struct {
		cell string
		want float64
	}{
		{cell: "B1", want: 8000},
		{cell: "B2", want: 8000},
		{cell: "B3", want: 0},
		{cell: "B4", want: 8000},
		{cell: "B5", want: 0},
		{cell: "B6", want: 3},
	}

	for _, tt := range tests {
		v, err := summary.GetValue(tt.cell)
		if err != nil {
			t.Fatalf("%s: %v", tt.cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != tt.want {
			t.Fatalf("%s = %#v, want %v", tt.cell, v, tt.want)
		}
	}
}

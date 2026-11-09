package werkbook

import (
	"bytes"
	"testing"

	"github.com/jpoz/werkbook/formula"
)

func TestSheetStoredExtentsMutations(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	check := func(s *Sheet, row, col int) {
		t.Helper()
		for i := 0; i < 2; i++ {
			if r, c := s.MaxRow(), s.MaxCol(); r != row || c != col {
				t.Fatalf("%s extents = (%d, %d), want (%d, %d)", s.Name(), r, c, row, col)
			}
		}
	}
	check(s, 0, 0)
	if err := s.SetRangeStyle("E4:E4", &Style{}); err != nil {
		t.Fatal(err)
	}
	check(s, 4, 5)
	if err := s.SetValue("C2", 1); err != nil {
		t.Fatal(err)
	}
	if err := s.SetFormula("H6", "1+1"); err != nil {
		t.Fatal(err)
	}
	check(s, 6, 8)
	if err := s.SetRowHeight(10, 20); err != nil {
		t.Fatal(err)
	}
	check(s, 10, 8)
	if err := s.SetValue("H6", nil); err != nil {
		t.Fatal(err)
	}
	check(s, 10, 8)
	if err := s.RemoveRowsFrom(6); err != nil {
		t.Fatal(err)
	}
	check(s, 4, 5)
	if err := s.RemoveRow(1); err != nil {
		t.Fatal(err)
	}
	check(s, 3, 5)
	copy, err := f.CopySheet("Sheet1", "Copy")
	if err != nil {
		t.Fatal(err)
	}
	check(copy, 3, 5)
	if err := copy.SetValue("Z20", 2); err != nil {
		t.Fatal(err)
	}
	check(copy, 20, 26)
	check(s, 3, 5)
	if err := copy.RemoveRow(20); err != nil {
		t.Fatal(err)
	}
	check(copy, 3, 5)
	if err := s.RemoveRowsFrom(1); err != nil {
		t.Fatal(err)
	}
	check(s, 0, 0)
}

func TestSheetStoredExtentsSpills(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.SetFormula("A1", "SEQUENCE(3,4)"); err != nil {
		t.Fatal(err)
	}
	check := func(row, col int, want float64) {
		t.Helper()
		if r, c := s.MaxRow(), s.MaxCol(); r != 1 || c != 1 {
			t.Fatalf("stored extents = (%d, %d), want (1, 1)", r, c)
		}
		fr := &fileResolver{file: f, currentSheet: s.Name()}
		values := fr.GetRangeValues(formula.RangeAddr{FromCol: 1, FromRow: 1, ToCol: MaxColumns, ToRow: MaxRows})
		if len(values) != row || len(values[0]) != col {
			t.Fatalf("spill range shape = %dx%d, want %dx%d", len(values), len(values[0]), row, col)
		}
		if v := values[row-1][col-1]; v.Type != formula.ValueNumber || v.Num != want {
			t.Fatalf("spill last value = %+v, want %g", v, want)
		}
	}
	f.Recalculate()
	check(3, 4, 12)
	if err := s.SetFormula("A1", "SEQUENCE(5,6)"); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()
	check(5, 6, 30)
	if err := s.SetFormula("A1", "SEQUENCE(2,2)"); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()
	check(2, 2, 4)
	if err := s.SetValue("B2", 99); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()
	if r, c := s.MaxRow(), s.MaxCol(); r != 2 || c != 2 {
		t.Fatalf("blocked spill extents = (%d, %d), want (2, 2)", r, c)
	}
	blocked, err := s.GetValue("A1")
	if err != nil || blocked.Type != TypeError || blocked.String != "#SPILL!" {
		t.Fatalf("blocked spill = %+v, %v", blocked, err)
	}
	if err := s.RemoveRow(2); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()
	check(2, 2, 4)
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	loaded, err := OpenReader(bytes.NewReader(buf.Bytes()), WithSpillCache())
	if err != nil {
		t.Fatal(err)
	}
	ls := loaded.Sheet("Sheet1")
	if r, c := ls.MaxRow(), ls.MaxCol(); r != 2 || c != 2 {
		t.Fatalf("imported shadow extents = (%d, %d), want (2, 2)", r, c)
	}
	loaded.ClearSpillShadowValues()
	if r, c := ls.MaxRow(), ls.MaxCol(); r != 2 || c != 2 {
		t.Fatalf("cleared shadow extents = (%d, %d), want (2, 2)", r, c)
	}
	if err := s.SetValue("A1", 7); err != nil {
		t.Fatal(err)
	}
	check(1, 1, 7)
}

func TestBoundedRangeStoredExtents(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	for _, ref := range []string{"B2", "D4", "Z4"} {
		if err := s.SetValue(ref, 1); err != nil {
			t.Fatal(err)
		}
	}
	if err := s.SetFormula("D4", "1/0"); err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		name                           string
		fromCol, fromRow, toCol, toRow int
		rows, cols                     int
	}{
		{"partially outside", 2, 2, 30, 8, 3, 29},
		{"outside rows", 2, 10, 4, 12, 1, 3},
		{"outside columns", 30, 2, 32, 4, 3, 3},
		{"empty row", 2, 3, 4, 3, 1, 3},
		{"full column", 2, 1, 2, MaxRows, 4, 1},
		{"full row", 1, 4, MaxColumns, 4, 1, 26},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			fr := &fileResolver{file: f, currentSheet: s.Name()}
			values := fr.GetRangeValues(formula.RangeAddr{FromCol: tt.fromCol, FromRow: tt.fromRow, ToCol: tt.toCol, ToRow: tt.toRow})
			if len(values) != tt.rows || len(values[0]) != tt.cols {
				t.Fatalf("shape = %dx%d, want %dx%d", len(values), len(values[0]), tt.rows, tt.cols)
			}
			for ri, row := range values {
				for ci, value := range row {
					r, c := ri+tt.fromRow, ci+tt.fromCol
					want := formula.EmptyVal()
					if r == 2 && c == 2 || r == 4 && c == 26 {
						want = formula.NumberVal(1)
					} else if r == 4 && c == 4 {
						want = formula.ErrorVal(formula.ErrValDIV0)
					}
					if value.Type != want.Type || value.Num != want.Num || value.Str != want.Str || value.Err != want.Err {
						t.Fatalf("cell (%d,%d) = %+v, want %+v", c, r, value, want)
					}
				}
			}
		})
	}
}

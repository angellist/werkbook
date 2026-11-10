package werkbook

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/jpoz/werkbook/formula"
)

var perfWorkbookSink *File
var perfRangeSink [][]formula.Value

// Measure the open path rather than the standalone expansion helper so that
// workbook-level name lookup reuse is included in the measurement.
func BenchmarkACEOpenDefinedNames(b *testing.B) {
	data := perfNamedWorkbook(b, 1, 200, 1000)
	b.ReportAllocs()
	for b.Loop() {
		f, err := OpenReader(bytes.NewReader(data))
		if err != nil {
			b.Fatal(err)
		}
		perfWorkbookSink = f
	}
}

func BenchmarkACEOpenRecalculate(b *testing.B) {
	data := perfNamedWorkbook(b, 4, 100, 1000)
	b.ReportAllocs()
	for b.Loop() {
		f, err := OpenReader(bytes.NewReader(data))
		if err != nil {
			b.Fatal(err)
		}
		f.Recalculate()
		perfWorkbookSink = f
	}
}

func perfNamedWorkbook(b *testing.B, sheets, formulas, names int) []byte {
	b.Helper()
	f := New()
	for i := 0; i < names; i++ {
		if err := f.SetDefinedName(DefinedName{Name: fmt.Sprintf("ACE_Name_%d", i), Value: "Sheet1!$A$1", LocalSheetID: -1}); err != nil {
			b.Fatal(err)
		}
	}
	for i := 0; i < sheets; i++ {
		s := f.Sheet("Sheet1")
		if i > 0 {
			var err error
			s, err = f.NewSheet(fmt.Sprintf("Statement%d", i))
			if err != nil {
				b.Fatal(err)
			}
		}
		if err := s.SetValue("A1", 42); err != nil {
			b.Fatal(err)
		}
		for row := 1; row <= formulas; row++ {
			if err := s.SetFormula(fmt.Sprintf("B%d", row), fmt.Sprintf("ACE_Name_%d+1", row%names)); err != nil {
				b.Fatal(err)
			}
		}
	}
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		b.Fatal(err)
	}
	return buf.Bytes()
}

func BenchmarkACERangeRead(b *testing.B) {
	for _, shape := range []struct {
		name       string
		rows, cols int
	}{
		{"wide", 100, 1024},
		{"tall", 10000, 8},
	} {
		b.Run(shape.name, func(b *testing.B) {
			f := New()
			s := f.Sheet("Sheet1")
			for row := 1; row <= shape.rows; row++ {
				for col := 1; col <= shape.cols; col++ {
					axis, err := CoordinatesToCellName(col, row)
					if err != nil {
						b.Fatal(err)
					}
					if err := s.SetValue(axis, 1); err != nil {
						b.Fatal(err)
					}
				}
			}
			fr := &fileResolver{file: f, currentSheet: s.Name()}
			addr := formula.RangeAddr{FromCol: 1, FromRow: 1, ToCol: 8, ToRow: 10}
			b.ReportAllocs()
			for b.Loop() {
				perfRangeSink = fr.GetRangeValues(addr)
			}
		})
	}
}

package werkbook

import (
	"fmt"
	"testing"

	"github.com/jpoz/werkbook/formula"
)

var wideRangeSink [][]formula.Value

func BenchmarkBoundedRangeWideSheet(b *testing.B) {
	for _, width := range []int{64, 1024, 8192} {
		f := New()
		s := f.Sheet("Sheet1")
		for row := 1; row <= 100; row++ {
			r := s.ensureRow(row)
			for col := 1; col <= width; col++ {
				r.ensureCell(col).value = Value{Type: TypeNumber, Number: 1}
			}
		}
		fr := &fileResolver{file: f, currentSheet: s.Name()}
		for _, shape := range []struct {
			name       string
			cols, rows int
		}{
			{"row", 10, 1},
			{"column", 1, 100},
		} {
			addr := formula.RangeAddr{FromCol: 1, FromRow: 1, ToCol: shape.cols, ToRow: shape.rows}
			wideRangeSink = fr.GetRangeValues(addr)
			b.Run(fmt.Sprintf("width=%d/%s", width, shape.name), func(b *testing.B) {
				b.ReportAllocs()
				for i := 0; i < b.N; i++ {
					wideRangeSink = fr.GetRangeValues(addr)
				}
			})
		}
	}
}

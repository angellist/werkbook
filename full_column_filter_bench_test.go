package werkbook_test

import (
	"fmt"
	"testing"

	"github.com/jpoz/werkbook"
)

var fullColumnFilterSiblingBenchmarkSink int

// BenchmarkResolveDefinedNameFullColumnSpillPostRecalc measures the named-range
// extraction path after recalculation has already completed.
// a tall source sheet, eight sibling FILTER spill anchors, workbook
// recalculation completed up front, then a tall named-range read over the spill
// output.
func BenchmarkResolveDefinedNameFullColumnSpillPostRecalc(b *testing.B) {
	for _, sourceRows := range []int{5000, 10000, 20000} {
		b.Run(fmt.Sprintf("rows=%d", sourceRows), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				f, expectedRows := buildFullColumnFilterSiblingWorkbook(b, sourceRows, 800)
				f.Recalculate()
				b.StartTimer()

				vals, err := f.ResolveDefinedName("Delinquents", -1)
				if err != nil {
					b.Fatal(err)
				}

				cols := 0
				if len(vals) > 0 {
					cols = len(vals[0])
				}
				if cols != 8 {
					b.Fatalf("ResolveDefinedName cols = %d, want 8", cols)
				}
				if len(vals) != expectedRows+1 {
					b.Fatalf("ResolveDefinedName rows = %d, want %d", len(vals), expectedRows+1)
				}
				fullColumnFilterSiblingBenchmarkSink = expectedRows + cols + len(vals)
			}
		})
	}
}

// BenchmarkRecalculateFullColumnFilterSiblings measures the recalculation cost
// for the same workbook shape: a large source sheet and several sibling FILTER
// spill anchors that all reuse the same large criteria range.
func BenchmarkRecalculateFullColumnFilterSiblings(b *testing.B) {
	for _, sourceRows := range []int{5000, 10000, 20000} {
		b.Run(fmt.Sprintf("rows=%d", sourceRows), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				f, expectedRows := buildFullColumnFilterSiblingWorkbook(b, sourceRows, 800)
				b.StartTimer()

				f.Recalculate()

				b.StopTimer()
				vals, err := f.ResolveDefinedName("Delinquents", -1)
				if err != nil {
					b.Fatal(err)
				}
				cols := 0
				if len(vals) > 0 {
					cols = len(vals[0])
				}
				if cols != 8 {
					b.Fatalf("ResolveDefinedName cols = %d, want 8", cols)
				}
				if len(vals) != expectedRows+1 {
					b.Fatalf("ResolveDefinedName rows = %d, want %d", len(vals), expectedRows+1)
				}
				fullColumnFilterSiblingBenchmarkSink = expectedRows + cols + len(vals)
				b.StartTimer()
			}
		})
	}
}

func buildFullColumnFilterSiblingWorkbook(tb testing.TB, sourceRows, matchEvery int) (*werkbook.File, int) {
	tb.Helper()

	f := werkbook.New(werkbook.FirstSheet("Source"))
	source := f.Sheet("Source")

	headers := []string{"Account", "Borrower", "City", "State", "Balance", "DaysLate", "Bucket", "Delinquent"}
	for i, header := range headers {
		benchMustSetValue(tb, source, benchCellRef(tb, i+1, 1), header)
	}

	expectedRows := 0
	for row := 2; row <= sourceRows+1; row++ {
		id := row - 1
		delinquent := 0.0
		if id%matchEvery == 0 {
			delinquent = 1
			expectedRows++
		}
		benchMustSetValue(tb, source, benchCellRef(tb, 1, row), float64(id))
		benchMustSetValue(tb, source, benchCellRef(tb, 2, row), fmt.Sprintf("Borrower %06d", id))
		benchMustSetValue(tb, source, benchCellRef(tb, 3, row), fmt.Sprintf("City %02d", id%97))
		benchMustSetValue(tb, source, benchCellRef(tb, 4, row), fmt.Sprintf("ST%02d", id%50))
		benchMustSetValue(tb, source, benchCellRef(tb, 5, row), float64(id)*10)
		benchMustSetValue(tb, source, benchCellRef(tb, 6, row), float64(id%120))
		benchMustSetValue(tb, source, benchCellRef(tb, 7, row), fmt.Sprintf("Bucket %d", id%5))
		benchMustSetValue(tb, source, benchCellRef(tb, 8, row), delinquent)
	}

	out := benchMustNewSheet(tb, f, "Out - Delinquent")
	for i, header := range headers {
		benchMustSetValue(tb, out, benchCellRef(tb, i+1, 1), header)
	}
	lastRow := sourceRows + 1
	for col := 1; col <= 8; col++ {
		formula := fmt.Sprintf(
			`FILTER(Source!%s2:%s%d,Source!H2:H%d<>0,"No rows")`,
			werkbook.ColumnNumberToName(col),
			werkbook.ColumnNumberToName(col),
			lastRow,
			lastRow,
		)
		benchMustSetFormula(tb, out, benchCellRef(tb, col, 2), formula)
	}

	if err := f.SetDefinedName(werkbook.DefinedName{
		Name:         "Delinquents",
		Value:        `'Out - Delinquent'!$A:$H`,
		LocalSheetID: -1,
	}); err != nil {
		tb.Fatalf("SetDefinedName(Delinquents): %v", err)
	}

	return f, expectedRows
}

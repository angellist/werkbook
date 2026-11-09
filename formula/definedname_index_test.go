package formula

import (
	"errors"
	"fmt"
	"testing"
)

func TestDefinedNameIndex(t *testing.T) {
	names := []DefinedNameInfo{{"Rate", "1", -1}, {"RATE", "2", 0}, {"rate", "3", -1}, {"Rate", "4", 1}}
	for _, tt := range []struct {
		sheet     int
		src, want string
	}{
		{0, "rate+Other!Rate", "2+4"},
		{1, "rate+'First Sheet'!Rate", "4+2"},
		{-1, "RATE", "1"},
		{0, `"Rate"+SUM(A1:A2)`, `"Rate"+SUM(A1:A2)`},
	} {
		index := NewDefinedNameIndex(names, tt.sheet, []string{"First Sheet", "Other"})
		got, err := index.ExpandBounded(tt.src, 0)
		if err != nil || got != tt.want {
			t.Fatalf("got %q, %v; want %q", got, err, tt.want)
		}
		if _, err := index.ExpandBounded(tt.src, 1); !errors.Is(err, ErrFormulaTooLarge) {
			t.Fatalf("got %v, want size error", err)
		}
	}
	index := NewDefinedNameIndex(nil, 0, nil)
	if got, err := index.ExpandBounded("", 1); err != nil || got != "" {
		t.Fatalf("empty expansion: %q, %v", got, err)
	}
	if _, err := index.ExpandBounded("123", 2); !errors.Is(err, ErrFormulaTooLarge) {
		t.Fatalf("empty index failed to enforce input limit: %v", err)
	}
}

func BenchmarkDefinedNameIndex(b *testing.B) {
	names := make([]DefinedNameInfo, 1000)
	for i := range names {
		names[i] = DefinedNameInfo{Name: fmt.Sprintf("DefinedName_%04d", i), Value: "Sheet1!$A$1", LocalSheetID: -1}
	}
	for _, indexed := range []bool{false, true} {
		b.Run(fmt.Sprintf("indexed=%v", indexed), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				var index *DefinedNameIndex
				if indexed {
					index = NewDefinedNameIndex(names, 0, []string{"Sheet1"})
				}
				for j := 0; j < 1000; j++ {
					var err error
					if indexed {
						_, err = index.ExpandBounded("DefinedName_0999+1", MaxExpandedFormulaBytes)
					} else {
						_, err = ExpandDefinedNamesBounded("DefinedName_0999+1", names, 0, []string{"Sheet1"}, MaxExpandedFormulaBytes)
					}
					if err != nil {
						b.Fatal(err)
					}
				}
			}
		})
	}
}

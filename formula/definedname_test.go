package formula

import (
	"errors"
	"fmt"
	"strings"
	"testing"
)

func TestDefinedNameExpanderReuse(t *testing.T) {
	names := []DefinedNameInfo{
		{Name: "Rate", Value: "2", LocalSheetID: 0},
		{Name: "RATE", Value: "1", LocalSheetID: -1},
		{Name: "Rate", Value: "99", LocalSheetID: -1},
		{Name: "rate", Value: "3", LocalSheetID: 0},
		{Name: "Rate", Value: "4", LocalSheetID: 1},
		{Name: "SUM", Value: "5", LocalSheetID: -1},
		{Name: "A1", Value: "6", LocalSheetID: -1},
		{Name: "Cube", Value: "LAMBDA(n,n*n*n)", LocalSheetID: -1},
		{Name: "Cycle", Value: "Cycle", LocalSheetID: -1},
		{Name: "Huge", Value: strings.Repeat("1", 100), LocalSheetID: -1},
	}
	e := NewDefinedNameExpander(names, []string{"North", "South Ops", "Summary"})
	tests := []struct {
		src   string
		sheet int
		want  string
	}{
		{"rAtE+North!RATE+'south ops'!rate", 0, "3+3+4"},
		{"Rate", 1, "4"},
		{"Rate", 2, "1"},
		{"Rate", -1, "1"},
		{`SUM(A1:A2)+SUM+A1+Sheet1!A1+"Rate"`, 0, `SUM(6:A2)+5+6+Sheet1!A1+"Rate"`},
		{"Cube(2)", 0, "LAMBDA(n,n*n*n)(2)"},
		{"Cycle", 0, "Cycle"},
		{"North!B2:Rate", 0, "North!B2:Rate"},
		{"", 0, ""},
	}
	for pass := 0; pass < 2; pass++ {
		for _, tt := range tests {
			got, err := e.ExpandBounded(tt.src, tt.sheet, MaxExpandedFormulaBytes)
			if err != nil || got != tt.want {
				t.Errorf("ExpandBounded(%q, %d) = %q, %v; want %q", tt.src, tt.sheet, got, err, tt.want)
			}
		}
	}
	for _, src := range []string{"Huge", strings.Repeat("x", 101)} {
		if _, err := e.ExpandBounded(src, 0, 99); !errors.Is(err, ErrFormulaTooLarge) {
			t.Errorf("ExpandBounded(%q): want ErrFormulaTooLarge, got %v", src, err)
		}
	}
	if got, err := e.ExpandBounded("Huge", 0, 0); err != nil || len(got) != 100 {
		t.Fatalf("unbounded expansion = %q, %v", got, err)
	}
	names[3].Value = "mutated"
	if got, err := e.ExpandBounded("Rate", 0, 100); err != nil || got != "3" {
		t.Fatalf("snapshot changed: %q, %v", got, err)
	}
	empty := NewDefinedNameExpander(nil, nil)
	if got, err := empty.ExpandBounded("A1+1", 0, 10); err != nil || got != "A1+1" {
		t.Fatalf("empty names = %q, %v", got, err)
	}
	if _, err := empty.ExpandBounded("A1+1", 0, 2); !errors.Is(err, ErrFormulaTooLarge) {
		t.Fatalf("empty names size limit: %v", err)
	}
}

func BenchmarkDefinedNameExpansion(b *testing.B) {
	for _, count := range []int{100, 1000, 10000} {
		names := make([]DefinedNameInfo, count)
		for i := range names {
			names[i] = DefinedNameInfo{Name: fmt.Sprintf("Name_%d", i), Value: "42", LocalSheetID: -1}
		}
		for _, reuse := range []bool{false, true} {
			b.Run(fmt.Sprintf("names=%d/reuse=%t", count, reuse), func(b *testing.B) {
				b.ReportAllocs()
				for i := 0; i < b.N; i++ {
					var e *DefinedNameExpander
					if reuse {
						e = NewDefinedNameExpander(names, nil)
					}
					for j := 0; j < 1000; j++ {
						if !reuse {
							e = NewDefinedNameExpander(names, nil)
						}
						if _, err := e.ExpandBounded("SUM(A1:A10)+Name_0", j%10, MaxExpandedFormulaBytes); err != nil {
							b.Fatal(err)
						}
					}
				}
			})
		}
	}
}

func TestExpandDefinedNames(t *testing.T) {
	names := []DefinedNameInfo{
		{Name: "OneRange", Value: "Sheet1!$A$10", LocalSheetID: -1},
		{Name: "valuation_date", Value: "Sheet1!$A$1", LocalSheetID: -1},
		{Name: "MyConst", Value: "42", LocalSheetID: -1},
	}

	tests := []struct {
		formula  string
		sheetIdx int
		want     string
	}{
		// Simple name replacement.
		{"B1+OneRange", 0, "B1+Sheet1!$A$10"},
		// Case-insensitive match.
		{"B1+onerange", 0, "B1+Sheet1!$A$10"},
		// Multiple names in one formula.
		{"OneRange+valuation_date", 0, "Sheet1!$A$10+Sheet1!$A$1"},
		// Name inside string literal should NOT be replaced.
		{`"OneRange"`, 0, `"OneRange"`},
		// Name followed by '(' is a function call, not replaced.
		{"OneRange(1,2)", 0, "OneRange(1,2)"},
		// Name followed by '!' is a sheet prefix, not replaced.
		{"OneRange!A1", 0, "OneRange!A1"},
		// Name preceded by '!' (sheet-qualified ref), not replaced.
		{"Sheet1!OneRange", 0, "Sheet1!OneRange"},
		// Constant value.
		{"A1+MyConst", 0, "A1+42"},
		// No names to expand.
		{"A1+B1", 0, "A1+B1"},
		// Name at end of formula.
		{"SUM(A1:A5)+OneRange", 0, "SUM(A1:A5)+Sheet1!$A$10"},
		// Name appearing as a substring inside a quoted sheet name must
		// not be expanded — otherwise 'my-OneRange'!A1 would corrupt into
		// 'my-Sheet1!$A$10'!A1.
		{"'my-OneRange'!A1", 0, "'my-OneRange'!A1"},
		{"SUM('my-OneRange'!A:A)", 0, "SUM('my-OneRange'!A:A)"},
		// Doubled '' inside a quoted sheet name represents a literal quote
		// and should not end the quoted region early.
		{"'Bob''s OneRange'!A1", 0, "'Bob''s OneRange'!A1"},
	}

	for _, tt := range tests {
		got := ExpandDefinedNames(tt.formula, names, tt.sheetIdx, nil)
		if got != tt.want {
			t.Errorf("ExpandDefinedNames(%q) = %q, want %q", tt.formula, got, tt.want)
		}
	}
}

func TestExpandDefinedNamesLocalScope(t *testing.T) {
	names := []DefinedNameInfo{
		{Name: "Rate", Value: "0.05", LocalSheetID: 0},               // local to sheet 0
		{Name: "Rate", Value: "0.10", LocalSheetID: 1},               // local to sheet 1
		{Name: "GlobalName", Value: "Sheet1!$B$2", LocalSheetID: -1}, // global
	}

	// On sheet 0, should get the sheet-0-local value.
	got := ExpandDefinedNames("A1*Rate", names, 0, nil)
	if want := "A1*0.05"; got != want {
		t.Errorf("sheet 0: got %q, want %q", got, want)
	}

	// On sheet 1, should get the sheet-1-local value.
	got = ExpandDefinedNames("A1*Rate", names, 1, nil)
	if want := "A1*0.10"; got != want {
		t.Errorf("sheet 1: got %q, want %q", got, want)
	}

	// On sheet 2, no local Rate is visible, so Rate is not expanded.
	got = ExpandDefinedNames("A1*Rate", names, 2, nil)
	if want := "A1*Rate"; got != want {
		t.Errorf("sheet 2: got %q, want %q", got, want)
	}

	// Global name should be visible on any sheet.
	got = ExpandDefinedNames("GlobalName+1", names, 2, nil)
	if want := "Sheet1!$B$2+1"; got != want {
		t.Errorf("global on sheet 2: got %q, want %q", got, want)
	}
}

func TestExpandDefinedNamesQualifiedLocalScope(t *testing.T) {
	names := []DefinedNameInfo{
		{Name: "Rate", Value: "North!$K$2", LocalSheetID: 1},
		{Name: "Rate", Value: "'South Ops'!$K$2", LocalSheetID: 2},
		{Name: "DeskLabel", Value: "'South Ops'!$K$4", LocalSheetID: 2},
	}
	sheetNames := []string{"Controls", "North", "South Ops", "Summary"}

	got := ExpandDefinedNames("'South Ops'!Rate+North!Rate", names, 3, sheetNames)
	if want := "'South Ops'!$K$2+North!$K$2"; got != want {
		t.Fatalf("qualified local names: got %q, want %q", got, want)
	}

	got = ExpandDefinedNames("IF(TRUE,'South Ops'!DeskLabel,\"x\")", names, 3, sheetNames)
	if want := "IF(TRUE,'South Ops'!$K$4,\"x\")"; got != want {
		t.Fatalf("quoted local string name: got %q, want %q", got, want)
	}

	got = ExpandDefinedNames("'South Ops'!B2:Rate", names, 3, sheetNames)
	if want := "'South Ops'!B2:Rate"; got != want {
		t.Fatalf("range endpoint should stay unchanged, got %q", got)
	}
}

func TestExpandDefinedNamesEmpty(t *testing.T) {
	// No names at all should return formula unchanged.
	got := ExpandDefinedNames("A1+B1", nil, 0, nil)
	if want := "A1+B1"; got != want {
		t.Errorf("nil names: got %q, want %q", got, want)
	}
}

func TestExpandDefinedNamesBoundedRejectsOversizeExpansion(t *testing.T) {
	names := []DefinedNameInfo{
		{
			Name:         "Huge",
			Value:        strings.Repeat("1+", MaxExpandedFormulaBytes/2),
			LocalSheetID: -1,
		},
	}

	_, err := ExpandDefinedNamesBounded("Huge+Huge", names, 0, nil, MaxExpandedFormulaBytes)
	if !errors.Is(err, ErrFormulaTooLarge) {
		t.Fatalf("expected ErrFormulaTooLarge, got %v", err)
	}
}

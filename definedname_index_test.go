package werkbook

import (
	"errors"
	"strings"
	"testing"

	"github.com/jpoz/werkbook/formula"
)

func TestDefinedNameIndexMutations(t *testing.T) {
	f := New()
	if _, err := f.NewSheet("Other"); err != nil {
		t.Fatal(err)
	}
	check := func(sheet, src, want string) {
		t.Helper()
		got, err := f.expandFormula(src, sheet, 1)
		if err != nil || got != want {
			t.Fatalf("expand(%q, %q) = %q, %v; want %q", sheet, src, got, err, want)
		}
	}
	check("Sheet1", "Rate", "Rate")
	f.AddDefinedName(DefinedName{Name: "Rate", Value: "10", LocalSheetID: -1})
	check("Sheet1", "rAtE", "10")
	index := f.definedNameIndexes[0]
	check("Sheet1", "Rate+Rate", "10+10")
	if f.definedNameIndexes[0] != index {
		t.Fatal("lookup was rebuilt without mutation")
	}
	f.AddDefinedName(DefinedName{Name: "RATE", Value: "20", LocalSheetID: 0})
	check("Sheet1", "Rate", "20")
	check("Other", "Rate+Sheet1!Rate", "10+20")
	if err := f.SetDefinedName(DefinedName{Name: "rate", Value: "30", LocalSheetID: 0}); err != nil {
		t.Fatal(err)
	}
	check("Other", "Sheet1!RATE", "30")
	if err := f.DeleteDefinedName("RATE", 0); err != nil {
		t.Fatal(err)
	}
	check("Sheet1", "Rate", "10")
	if err := f.DeleteDefinedName("Rate", -1); err != nil {
		t.Fatal(err)
	}
	f.AddDefinedName(DefinedName{Name: "RenamedRate", Value: "40", LocalSheetID: -1})
	check("Sheet1", "Rate+RenamedRate", "Rate+40")
	f.AddDefinedName(DefinedName{Name: "Local", Value: "Sheet1!$A$1", LocalSheetID: 0})
	check("Other", "Sheet1!Local", "Sheet1!$A$1")
	if err := f.SetSheetName("Sheet1", "New Sheet"); err != nil {
		t.Fatal(err)
	}
	check("Other", "'New Sheet'!Local", "'New Sheet'!$A$1")
	if _, err := f.CopySheet("New Sheet", "Copy"); err != nil {
		t.Fatal(err)
	}
	check("Copy", "Local+RenamedRate+'New Sheet'!Local", "Local+40+'New Sheet'!$A$1")
	f.AddDefinedName(DefinedName{Name: "CopyLocal", Value: "50", LocalSheetID: 2})
	check("Other", "Copy!CopyLocal", "50")
	if err := f.DeleteSheet("New Sheet"); err != nil {
		t.Fatal(err)
	}
	check("Copy", "CopyLocal+RenamedRate", "50+40")
	check("Other", "Copy!CopyLocal", "50")
	if err := f.DeleteDefinedName("absent", -1); err == nil {
		t.Fatal("expected missing-name error")
	}
	check("Copy", "CopyLocal", "50")
}

func TestDefinedNameIndexBoundedAfterMutation(t *testing.T) {
	f := New()
	f.AddDefinedName(DefinedName{Name: "Large", Value: "1", LocalSheetID: -1})
	if _, err := f.expandFormula("Large", "Sheet1", 1); err != nil {
		t.Fatal(err)
	}
	if err := f.SetDefinedName(DefinedName{Name: "Large", Value: strings.Repeat("x", formula.MaxExpandedFormulaBytes+1), LocalSheetID: -1}); err != nil {
		t.Fatal(err)
	}
	if _, err := f.expandFormula("Large", "Sheet1", 1); !errors.Is(err, formula.ErrFormulaTooLarge) {
		t.Fatalf("got %v, want expansion limit error", err)
	}
}

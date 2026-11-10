package werkbook

import "testing"

func TestDefinedNameExpansionCacheInvalidation(t *testing.T) {
	f := New()
	if _, err := f.NewSheet("South"); err != nil {
		t.Fatal(err)
	}
	check := func(src, sheet, want string) {
		t.Helper()
		got, err := f.expandFormula(src, sheet, 1)
		if err != nil || got != want {
			t.Fatalf("expandFormula(%q, %q) = %q, %v; want %q", src, sheet, got, err, want)
		}
	}
	check("Rate", "Sheet1", "Rate")
	f.AddDefinedName(DefinedName{Name: "Rate", Value: "1", LocalSheetID: -1})
	check("Rate", "Sheet1", "1")
	if err := f.SetDefinedName(DefinedName{Name: "rate", Value: "2", LocalSheetID: -1}); err != nil {
		t.Fatal(err)
	}
	check("Rate", "Sheet1", "2")
	f.AddDefinedName(DefinedName{Name: "RATE", Value: "South!$A$1", LocalSheetID: 1})
	check("Rate+South!Rate", "Sheet1", "2+South!$A$1")
	check("Rate", "South", "South!$A$1")
	if err := f.SetSheetName("South", "New South"); err != nil {
		t.Fatal(err)
	}
	check("Rate+'New South'!Rate", "Sheet1", "2+'New South'!$A$1")
	if err := f.DeleteDefinedName("rate", 1); err != nil {
		t.Fatal(err)
	}
	check("Rate", "New South", "2")
	if err := f.DeleteDefinedName("RATE", -1); err != nil {
		t.Fatal(err)
	}
	check("Rate", "Sheet1", "Rate")
	f.AddDefinedName(DefinedName{Name: "Rate", Value: "3", LocalSheetID: 1})
	check("Rate", "New South", "3")
	if err := f.DeleteSheet("Sheet1"); err != nil {
		t.Fatal(err)
	}
	check("Rate+'New South'!Rate", "New South", "3+3")
}

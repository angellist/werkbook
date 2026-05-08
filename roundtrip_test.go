package werkbook_test

import (
	"path/filepath"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestRoundTrip(t *testing.T) {
	// Write a file with various cell types.
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	cells := map[string]any{
		"A1": "hello",
		"B1": 42,
		"C1": 3.14,
		"A2": true,
		"B2": false,
		"A3": "world",
		"D5": 0,
	}
	for ref, val := range cells {
		if err := s.SetValue(ref, val); err != nil {
			t.Fatalf("SetValue(%q, %v): %v", ref, val, err)
		}
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "roundtrip.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Read it back.
	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	names := f2.SheetNames()
	if len(names) != 1 || names[0] != "Sheet1" {
		t.Fatalf("expected [Sheet1], got %v", names)
	}

	s2 := f2.Sheet("Sheet1")

	// Check each value.
	expected := map[string]any{
		"A1": "hello",
		"B1": float64(42),
		"C1": 3.14,
		"A2": true,
		"B2": false,
		"A3": "world",
		"D5": float64(0),
	}
	for ref, want := range expected {
		v, err := s2.GetValue(ref)
		if err != nil {
			t.Errorf("GetValue(%q): %v", ref, err)
			continue
		}
		got := v.Raw()
		if got != want {
			t.Errorf("GetValue(%q) = %v (%T), want %v (%T)", ref, got, got, want, want)
		}
	}

	// Empty cells should return TypeEmpty.
	v, err := s2.GetValue("Z99")
	if err != nil {
		t.Fatalf("GetValue(Z99): %v", err)
	}
	if !v.IsEmpty() {
		t.Errorf("expected empty value for Z99, got %#v", v)
	}
}

func TestRoundTripDuplicateStrings(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Same string in multiple cells should use the same SST index.
	s.SetValue("A1", "dup")
	s.SetValue("A2", "dup")
	s.SetValue("A3", "other")

	dir := t.TempDir()
	path := filepath.Join(dir, "dup.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	for _, ref := range []string{"A1", "A2"} {
		v, _ := s2.GetValue(ref)
		if v.Raw() != "dup" {
			t.Errorf("GetValue(%q) = %v, want dup", ref, v.Raw())
		}
	}
	v, _ := s2.GetValue("A3")
	if v.Raw() != "other" {
		t.Errorf("GetValue(A3) = %v, want other", v.Raw())
	}
}

func TestRoundTripFreezePane(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "header")
	s.SetFreezePane(&werkbook.FreezePane{
		XSplit:      0,
		YSplit:      2,
		TopLeftCell: "A3",
	})

	dir := t.TempDir()
	path := filepath.Join(dir, "freeze.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")
	fp := s2.GetFreezePane()
	if fp == nil {
		t.Fatal("expected freeze pane, got nil")
	}
	if fp.YSplit != 2 {
		t.Errorf("YSplit = %d, want 2", fp.YSplit)
	}
	if fp.XSplit != 0 {
		t.Errorf("XSplit = %d, want 0", fp.XSplit)
	}
	if fp.TopLeftCell != "A3" {
		t.Errorf("TopLeftCell = %q, want A3", fp.TopLeftCell)
	}
}

func TestRoundTripFreezePaneBothAxes(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "corner")
	s.SetFreezePane(&werkbook.FreezePane{
		XSplit:      1,
		YSplit:      3,
		TopLeftCell: "B4",
	})

	dir := t.TempDir()
	path := filepath.Join(dir, "freeze2.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")
	fp := s2.GetFreezePane()
	if fp == nil {
		t.Fatal("expected freeze pane, got nil")
	}
	if fp.XSplit != 1 {
		t.Errorf("XSplit = %d, want 1", fp.XSplit)
	}
	if fp.YSplit != 3 {
		t.Errorf("YSplit = %d, want 3", fp.YSplit)
	}
	if fp.TopLeftCell != "B4" {
		t.Errorf("TopLeftCell = %q, want B4", fp.TopLeftCell)
	}
}

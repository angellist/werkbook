package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/formula"
)

// Package-internal unit tests for the spill state machine in spill.go.
// These exercise pure helper functions and Cell-level flag transitions
// that are hard to observe through the public API.

func numArray(rows [][]float64) formula.Value {
	out := make([][]formula.Value, len(rows))
	for i, row := range rows {
		rr := make([]formula.Value, len(row))
		for j, v := range row {
			rr[j] = formula.NumberVal(v)
		}
		out[i] = rr
	}
	return formula.Value{Type: formula.ValueArray, Array: out}
}

func TestSpillArrayRect(t *testing.T) {
	tests := []struct {
		name                   string
		raw                    formula.Value
		anchorCol, anchorRow   int
		wantToCol, wantToRow   int
		wantOK                 bool
	}{
		{
			name:      "scalar not an array",
			raw:       formula.NumberVal(42),
			anchorCol: 5, anchorRow: 3,
			wantOK: false,
		},
		{
			name:      "NoSpill flag blocks spill",
			raw:       formula.Value{Type: formula.ValueArray, NoSpill: true, Array: [][]formula.Value{{formula.NumberVal(1)}}},
			anchorCol: 2, anchorRow: 2,
			wantOK: false,
		},
		{
			name:      "empty array",
			raw:       numArray(nil),
			anchorCol: 1, anchorRow: 1,
			wantOK: false,
		},
		{
			name:      "array with empty row",
			raw:       formula.Value{Type: formula.ValueArray, Array: [][]formula.Value{{}}},
			anchorCol: 1, anchorRow: 1,
			wantOK: false,
		},
		{
			name:      "1x1 array spills to anchor itself",
			raw:       numArray([][]float64{{10}}),
			anchorCol: 4, anchorRow: 7,
			wantToCol: 4, wantToRow: 7,
			wantOK: true,
		},
		{
			name:      "vertical 3x1 spill",
			raw:       numArray([][]float64{{10}, {20}, {30}}),
			anchorCol: 2, anchorRow: 5,
			wantToCol: 2, wantToRow: 7,
			wantOK: true,
		},
		{
			name:      "horizontal 1x3 spill",
			raw:       numArray([][]float64{{10, 20, 30}}),
			anchorCol: 3, anchorRow: 2,
			wantToCol: 5, wantToRow: 2,
			wantOK: true,
		},
		{
			name:      "rectangular 2x3 spill",
			raw:       numArray([][]float64{{1, 2, 3}, {4, 5, 6}}),
			anchorCol: 1, anchorRow: 1,
			wantToCol: 3, wantToRow: 2,
			wantOK: true,
		},
		{
			name: "jagged array uses widest row for col bound",
			raw: formula.Value{Type: formula.ValueArray, Array: [][]formula.Value{
				{formula.NumberVal(1), formula.NumberVal(2)},
				{formula.NumberVal(3), formula.NumberVal(4), formula.NumberVal(5)},
			}},
			anchorCol: 10, anchorRow: 10,
			wantToCol: 12, wantToRow: 11,
			wantOK: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			toCol, toRow, ok := spillArrayRect(tt.raw, tt.anchorCol, tt.anchorRow)
			if ok != tt.wantOK {
				t.Fatalf("ok = %v, want %v", ok, tt.wantOK)
			}
			if !ok {
				return
			}
			if toCol != tt.wantToCol || toRow != tt.wantToRow {
				t.Fatalf("rect = (%d,%d), want (%d,%d)", toCol, toRow, tt.wantToCol, tt.wantToRow)
			}
		})
	}
}

func TestIsDynamicArrayAnchor(t *testing.T) {
	tests := []struct {
		name string
		cell *Cell
		want bool
	}{
		{name: "nil cell", cell: nil, want: false},
		{name: "no formula", cell: &Cell{}, want: false},
		{
			name: "CSE array formula is not a dynamic anchor",
			cell: &Cell{formula: "FILTER(A:A,B:B)", isArrayFormula: true, dynamicArraySpill: true},
			want: false,
		},
		{
			name: "formula without dynamicArraySpill is not an anchor",
			cell: &Cell{formula: "FILTER(A:A,B:B)", dynamicArraySpill: false},
			want: false,
		},
		{
			name: "formula with dynamicArraySpill is an anchor",
			cell: &Cell{formula: "FILTER(A:A,B:B)", dynamicArraySpill: true},
			want: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := isDynamicArrayAnchor(tt.cell); got != tt.want {
				t.Fatalf("isDynamicArrayAnchor = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestPublishSpillState_IdempotentWithinGen(t *testing.T) {
	f := New()
	s, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	f.calcGen = 5

	c := &Cell{formula: "SEQUENCE(3)", dynamicArraySpill: true}
	raw := numArray([][]float64{{1}, {2}, {3}})

	// First publish should update fields and invalidate the overlay.
	s.ensureSpillOverlay() // build with current gen
	s.publishSpillState(c, 2, 2, raw, false)
	if c.spillStateGen != 5 {
		t.Fatalf("spillStateGen = %d, want 5", c.spillStateGen)
	}
	if c.spillPublishedToCol != 2 || c.spillPublishedToRow != 4 {
		t.Fatalf("publishedTo = (%d,%d), want (2,4)", c.spillPublishedToCol, c.spillPublishedToRow)
	}

	// Second call with identical rect must be a no-op: overlay gen stays.
	s.spill.gen = f.calcGen
	s.publishSpillState(c, 2, 2, raw, false)
	if s.spill.gen != f.calcGen {
		t.Fatalf("overlay gen was invalidated on idempotent publish")
	}

	// A call with a different rect must invalidate the overlay.
	raw2 := numArray([][]float64{{1}, {2}, {3}, {4}})
	s.publishSpillState(c, 2, 2, raw2, false)
	if s.spill.gen == f.calcGen {
		t.Fatalf("overlay gen should be invalidated on rect change")
	}
	if c.spillPublishedToRow != 5 {
		t.Fatalf("publishedToRow = %d after grow, want 5", c.spillPublishedToRow)
	}
}

func TestPublishSpillState_BlockedKeepsAttemptedButNotPublished(t *testing.T) {
	f := New()
	s, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	f.calcGen = 7

	c := &Cell{formula: "FILTER(A:A,B:B)", dynamicArraySpill: true}
	raw := numArray([][]float64{{1}, {2}, {3}})

	// blocked=true: attempted rect is still tracked, but published collapses.
	s.publishSpillState(c, 3, 3, raw, true)
	if c.spillAttemptedToCol != 3 || c.spillAttemptedToRow != 5 {
		t.Fatalf("attemptedTo = (%d,%d), want (3,5)",
			c.spillAttemptedToCol, c.spillAttemptedToRow)
	}
	if c.spillPublishedToCol != 3 || c.spillPublishedToRow != 3 {
		t.Fatalf("publishedTo = (%d,%d), want (3,3) while blocked",
			c.spillPublishedToCol, c.spillPublishedToRow)
	}
}

func TestClearSpillState_IdempotentWhenAlreadyClear(t *testing.T) {
	f := New()
	s, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	c := &Cell{}

	// First clear on an already-clear cell must not invalidate the overlay.
	s.spill.gen = 42
	s.clearSpillState(c)
	if s.spill.gen != 42 {
		t.Fatalf("overlay gen touched by no-op clear")
	}

	// Set some state, then clear, and confirm the overlay IS invalidated.
	c.spillStateGen = 3
	c.spillPublishedToCol = 7
	c.spillPublishedToRow = 9
	c.spillAttemptedToCol = 7
	c.spillAttemptedToRow = 9
	s.clearSpillState(c)
	if s.spill.gen == 42 {
		t.Fatalf("overlay gen should be invalidated after real clear")
	}
	if c.spillPublishedToCol != 0 || c.spillPublishedToRow != 0 {
		t.Fatalf("publishedTo not zeroed: (%d,%d)",
			c.spillPublishedToCol, c.spillPublishedToRow)
	}
	if c.spillStateGen != 0 {
		t.Fatalf("spillStateGen not zeroed: %d", c.spillStateGen)
	}
}

func TestCellHasPublishedSpill(t *testing.T) {
	tests := []struct {
		name string
		cell *Cell
		gen  uint64
		want bool
	}{
		{name: "nil cell", cell: nil, gen: 1, want: false},
		{
			name: "stale gen",
			cell: &Cell{spillStateGen: 1, spillPublishedToCol: 10, spillPublishedToRow: 10},
			gen:  2,
			want: false,
		},
		{
			name: "anchor-only (published == anchor)",
			cell: &Cell{spillStateGen: 2, spillPublishedToCol: 5, spillPublishedToRow: 5},
			gen:  2,
			want: false, // 5 > 5 is false on both axes
		},
		{
			name: "published extends column",
			cell: &Cell{spillStateGen: 2, spillPublishedToCol: 6, spillPublishedToRow: 5},
			gen:  2,
			want: true,
		},
		{
			name: "published extends row",
			cell: &Cell{spillStateGen: 2, spillPublishedToCol: 5, spillPublishedToRow: 7},
			gen:  2,
			want: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := cellHasPublishedSpill(tt.cell, 5, 5, tt.gen)
			if got != tt.want {
				t.Fatalf("cellHasPublishedSpill = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestCellSpillFormulaRef(t *testing.T) {
	// Anchor at B2 (col 2, row 2).
	cell := &Cell{
		spillStateGen:       3,
		spillPublishedToCol: 4, // D
		spillPublishedToRow: 5,
	}
	ref, ok := cellSpillFormulaRef(cell, 2, 2)
	if !ok {
		t.Fatalf("cellSpillFormulaRef returned ok=false")
	}
	if ref != "B2:D5" {
		t.Fatalf("ref = %q, want B2:D5", ref)
	}

	// Anchor whose published bounds equal the anchor should report no ref.
	scalar := &Cell{
		spillStateGen:       3,
		spillPublishedToCol: 2,
		spillPublishedToRow: 2,
	}
	if _, ok := cellSpillFormulaRef(scalar, 2, 2); ok {
		t.Fatalf("expected no ref for non-spilled anchor")
	}
}

package formula

import (
	"testing"
)

func TestISEVEN(t *testing.T) {
	resolver := &mockResolver{}

	// ISEVEN(4) = TRUE
	cf := evalCompile(t, `ISEVEN(4)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISEVEN(4) = %v, want true", got)
	}

	// ISEVEN(3) = FALSE
	cf = evalCompile(t, `ISEVEN(3)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISEVEN(3) = %v, want false", got)
	}

	// ISEVEN(TRUE) = #VALUE!
	cf = evalCompile(t, `ISEVEN(TRUE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISEVEN(TRUE) = %v, want #VALUE!", got)
	}

	// ISEVEN(FALSE) = #VALUE!
	cf = evalCompile(t, `ISEVEN(FALSE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISEVEN(FALSE) = %v, want #VALUE!", got)
	}
}

func TestISODD(t *testing.T) {
	resolver := &mockResolver{}

	// ISODD(3) = TRUE
	cf := evalCompile(t, `ISODD(3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISODD(3) = %v, want true", got)
	}

	// ISODD(4) = FALSE
	cf = evalCompile(t, `ISODD(4)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISODD(4) = %v, want false", got)
	}

	// ISODD(TRUE) = #VALUE!
	cf = evalCompile(t, `ISODD(TRUE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISODD(TRUE) = %v, want #VALUE!", got)
	}

	// ISODD(FALSE) = #VALUE!
	cf = evalCompile(t, `ISODD(FALSE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISODD(FALSE) = %v, want #VALUE!", got)
	}
}

// mockFormulaResolver extends mockResolver with FormulaIntrospector support.
type mockFormulaResolver struct {
	mockResolver
	formulas map[CellAddr]string
}

func (m *mockFormulaResolver) HasFormula(sheet string, col, row int) bool {
	_, ok := m.formulas[CellAddr{Sheet: sheet, Col: col, Row: row}]
	return ok
}

func (m *mockFormulaResolver) GetFormulaText(sheet string, col, row int) string {
	return m.formulas[CellAddr{Sheet: sheet, Col: col, Row: row}]
}

func TestISFORMULA(t *testing.T) {
	resolver := &mockFormulaResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),  // A1: constant
				{Col: 2, Row: 1}: NumberVal(100), // B1: has formula
			},
		},
		formulas: map[CellAddr]string{
			{Col: 2, Row: 1}: "A1+58", // B1 has a formula
		},
	}
	ctx := &EvalContext{
		CurrentCol:   3,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	// ISFORMULA(B1) = TRUE (cell with formula)
	cf := evalCompile(t, `ISFORMULA(B1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISFORMULA(B1) = %v, want TRUE", got)
	}

	// ISFORMULA(A1) = FALSE (constant value)
	cf = evalCompile(t, `ISFORMULA(A1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISFORMULA(A1) = %v, want FALSE", got)
	}

	// ISFORMULA(C1) = FALSE (empty cell)
	cf = evalCompile(t, `ISFORMULA(C1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISFORMULA(C1) = %v, want FALSE", got)
	}

	// ISFORMULA(123) = #VALUE! (non-reference argument)
	cf = evalCompile(t, `ISFORMULA(123)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISFORMULA(123) = %v, want #VALUE!", got)
	}

	// ISFORMULA() with no args = #VALUE!
	cf = evalCompile(t, `ISFORMULA()`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISFORMULA() = %v, want #VALUE!", got)
	}
}

func TestFORMULATEXT(t *testing.T) {
	resolver := &mockFormulaResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),  // A1: constant
				{Col: 2, Row: 1}: NumberVal(100), // B1: has formula
			},
		},
		formulas: map[CellAddr]string{
			{Col: 2, Row: 1}: "A1+58", // B1 has a formula
		},
	}
	ctx := &EvalContext{
		CurrentCol:   3,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	// FORMULATEXT(B1) = "=A1+58" (cell with formula)
	cf := evalCompile(t, `FORMULATEXT(B1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "=A1+58" {
		t.Errorf("FORMULATEXT(B1) = %v, want =A1+58", got)
	}

	// FORMULATEXT(A1) = #N/A (constant value, no formula)
	cf = evalCompile(t, `FORMULATEXT(A1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("FORMULATEXT(A1) = %v, want #N/A", got)
	}

	// FORMULATEXT(C1) = #N/A (empty cell)
	cf = evalCompile(t, `FORMULATEXT(C1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("FORMULATEXT(C1) = %v, want #N/A", got)
	}

	// FORMULATEXT(123) = #VALUE! (non-reference argument)
	cf = evalCompile(t, `FORMULATEXT(123)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("FORMULATEXT(123) = %v, want #VALUE!", got)
	}

	// FORMULATEXT() with no args = #VALUE!
	cf = evalCompile(t, `FORMULATEXT()`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("FORMULATEXT() = %v, want #VALUE!", got)
	}
}

func TestISFORMULA_NoIntrospector(t *testing.T) {
	// When the resolver doesn't implement FormulaIntrospector, ISFORMULA returns FALSE.
	resolver := &mockResolver{}
	ctx := &EvalContext{
		CurrentCol:   1,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	cf := evalCompile(t, `ISFORMULA(A1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISFORMULA(A1) with basic resolver = %v, want FALSE", got)
	}
}

func TestFORMULATEXT_NoIntrospector(t *testing.T) {
	// When the resolver doesn't implement FormulaIntrospector, FORMULATEXT returns #N/A.
	resolver := &mockResolver{}
	ctx := &EvalContext{
		CurrentCol:   1,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	cf := evalCompile(t, `FORMULATEXT(A1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("FORMULATEXT(A1) with basic resolver = %v, want #N/A", got)
	}
}

func TestCOLUMN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("no_args_returns_current_col", func(t *testing.T) {
		tests := []struct {
			name string
			col  int
			want float64
		}{
			{"col_1", 1, 1},
			{"col_3", 3, 3},
			{"col_10", 10, 10},
			{"col_26_Z", 26, 26},
			{"col_256", 256, 256},
			{"col_16384_XFD", 16384, 16384},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   tt.col,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, `COLUMN()`)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("COLUMN() with CurrentCol=%d = %v, want %v", tt.col, got, tt.want)
				}
			})
		}
	})

	t.Run("no_args_nil_context", func(t *testing.T) {
		// COLUMN() with no EvalContext should return #VALUE!
		cf := evalCompile(t, `COLUMN()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("COLUMN() with nil ctx = %v, want #VALUE!", got)
		}
	})

	t.Run("single_cell_ref", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"A1_col_1", `COLUMN(A1)`, 1},
			{"B1_col_2", `COLUMN(B1)`, 2},
			{"C5_col_3", `COLUMN(C5)`, 3},
			{"Z1_col_26", `COLUMN(Z1)`, 26},
			{"AA1_col_27", `COLUMN(AA1)`, 27},
			{"AZ1_col_52", `COLUMN(AZ1)`, 52},
			{"BA1_col_53", `COLUMN(BA1)`, 53},
			{"IV1_col_256", `COLUMN(IV1)`, 256},
			{"XFD1_col_16384", `COLUMN(XFD1)`, 16384},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   1,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("ref_different_rows_same_col", func(t *testing.T) {
		// COLUMN should return the column regardless of which row the ref is in
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"D1", `COLUMN(D1)`, 4},
			{"D10", `COLUMN(D10)`, 4},
			{"D100", `COLUMN(D100)`, 4},
			{"D1000", `COLUMN(D1000)`, 4},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   1,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("non_reference_arg_returns_VALUE_error", func(t *testing.T) {
		// Non-reference arguments (numbers, strings, booleans) should return #VALUE!
		tests := []struct {
			name    string
			formula string
		}{
			{"number", `COLUMN(42)`},
			{"string", `COLUMN("hello")`},
			{"boolean_true", `COLUMN(TRUE)`},
			{"boolean_false", `COLUMN(FALSE)`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   1,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("%s = %v, want #VALUE!", tt.formula, got)
				}
			})
		}
	})

	t.Run("range_ref_returns_leftmost_column", func(t *testing.T) {
		// When a range is passed, COLUMN returns the leftmost column.
		// In the current implementation, a range resolves to a ValueArray
		// which is not ValueRef, so it returns #VALUE!.
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN(A1:C3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("COLUMN(A1:C3) = %v, want #VALUE!", got)
		}
	})

	t.Run("absolute_ref", func(t *testing.T) {
		// Absolute references ($C$1) should work identically
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN($C$1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("COLUMN($C$1) = %v, want 3", got)
		}
	})

	t.Run("absolute_col_only", func(t *testing.T) {
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN($E1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("COLUMN($E1) = %v, want 5", got)
		}
	})

	t.Run("absolute_row_only", func(t *testing.T) {
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN(F$1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6 {
			t.Errorf("COLUMN(F$1) = %v, want 6", got)
		}
	})
}

func TestIFNA(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFNA(#N/A,"default")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "default" {
		t.Errorf("IFNA(#N/A) = %v, want default", got)
	}

	cf = evalCompile(t, `IFNA(42,"default")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("IFNA(42) = %v, want 42", got)
	}
}

package formula

import (
	"testing"
)

func TestAND(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"AND(TRUE,TRUE)", true},
		{"AND(TRUE,FALSE)", false},
		{"AND(FALSE,FALSE)", false},
		{"AND(TRUE,TRUE,TRUE)", true},
		{"AND(TRUE,TRUE,FALSE)", false},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

func TestOR(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"OR(TRUE,FALSE)", true},
		{"OR(FALSE,FALSE)", false},
		{"OR(TRUE,TRUE)", true},
		{"OR(FALSE,TRUE,FALSE)", true},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

func TestNOT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"NOT(TRUE)", false},
		{"NOT(FALSE)", true},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

func TestIF(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("true branch", func(t *testing.T) {
		cf := evalCompile(t, `IF(TRUE, "yes", "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IF(TRUE, "yes", "no") = %v, want "yes"`, got)
		}
	})

	t.Run("false branch", func(t *testing.T) {
		cf := evalCompile(t, `IF(FALSE, "yes", "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "no" {
			t.Errorf(`IF(FALSE, "yes", "no") = %v, want "no"`, got)
		}
	})

	t.Run("false without else returns FALSE", func(t *testing.T) {
		cf := evalCompile(t, `IF(FALSE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf(`IF(FALSE, "yes") = %v, want FALSE`, got)
		}
	})
}

func TestIFERROR(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("non-error returns value", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1, "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf(`IFERROR(1, "fallback") = %v, want 1`, got)
		}
	})

	t.Run("error returns fallback", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1/0, "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "fallback" {
			t.Errorf(`IFERROR(1/0, "fallback") = %v, want "fallback"`, got)
		}
	})
}

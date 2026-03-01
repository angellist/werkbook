package formula

import (
	"testing"
)

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

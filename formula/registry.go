package formula

import (
	"fmt"
	"strings"
)

// Func is the standard signature for all registered formula functions.
type Func func(args []Value, ctx *EvalContext) (Value, error)

var (
	registry = map[string]Func{}
	nameToID = map[string]int{}
	idToName []string
)

// Register adds a formula function to the global registry.
// It is safe to call from init(). Duplicate names overwrite silently,
// allowing external packages (e.g. werkbook-pro) to override or extend.
func Register(name string, fn Func) {
	upper := strings.ToUpper(name)
	if _, exists := nameToID[upper]; !exists {
		id := len(idToName)
		idToName = append(idToName, upper)
		nameToID[upper] = id
	}
	registry[upper] = fn
}

// LookupFunc returns the function ID for use by the compiler.
// Returns -1 if the function is not registered.
func LookupFunc(name string) int {
	id, ok := nameToID[strings.ToUpper(name)]
	if !ok {
		return -1
	}
	return id
}

// CallFunc dispatches a function call by ID at eval time.
func CallFunc(funcID int, args []Value, ctx *EvalContext) (Value, error) {
	if funcID < 0 || funcID >= len(idToName) {
		return Value{}, fmt.Errorf("unknown function ID %d", funcID)
	}
	name := idToName[funcID]
	fn := registry[name]
	if fn == nil {
		return Value{}, fmt.Errorf("unimplemented function: %s", name)
	}
	return fn(args, ctx)
}

// RegisteredFunctions returns the names of all registered functions.
func RegisteredFunctions() []string {
	out := make([]string, len(idToName))
	copy(out, idToName)
	return out
}

// NoCtx wraps a function that doesn't need EvalContext into a Func.
func NoCtx(fn func([]Value) (Value, error)) Func {
	return func(args []Value, _ *EvalContext) (Value, error) {
		return fn(args)
	}
}

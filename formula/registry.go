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

// arrayForcingFuncs lists functions that evaluate their arguments in array
// context, suppressing implicit intersection. In legacy (non-dynamic-array)
// these functions treat expressions like range*range element-wise
// even when the formula is not entered as CSE (Ctrl+Shift+Enter).
var arrayForcingFuncs = map[string]bool{
	"SUMPRODUCT": true,
	"MMULT":      true,
	"TREND":      true,
	"GROWTH":     true,
	"LINEST":     true,
	"LOGEST":     true,
	"FREQUENCY":  true,
	"TRANSPOSE":  true,
	// Functions that accept range arguments and must not have them
	// implicitly intersected to a single cell.
	"SUMIF":      true,
	"SUMIFS":     true,
	"COUNTIF":    true,
	"COUNTIFS":   true,
	"AVERAGEIF":  true,
	"AVERAGEIFS": true,
	"MAXIFS":     true,
	"MINIFS":     true,
	"MATCH":      true,
	"INDEX":      true,
	"LOOKUP":     true,
	"VLOOKUP":    true,
	"HLOOKUP":    true,
}

// FuncArgEvalMode describes how the compiler should evaluate an argument for a
// particular function in a legacy non-array formula context.
type FuncArgEvalMode int

const (
	FuncArgEvalDefault FuncArgEvalMode = iota
	FuncArgEvalDirectRange
	FuncArgEvalArray
)

// directRangeAllArgFuncs preserve plain direct range references like A:A or
// 1:1 for every argument position, but still allow expressions like A:A*B:B to
// follow legacy implicit-intersection behavior unless the function is fully
// array-forcing.
var directRangeAllArgFuncs = map[string]bool{
	"AVERAGE":   true,
	"AVERAGEA":  true,
	"AVEDEV":    true,
	"COLUMNS":   true,
	"COUNT":     true,
	"COUNTA":    true,
	"DEVSQ":     true,
	"GEOMEAN":   true,
	"HARMEAN":   true,
	"KURT":      true,
	"MAX":       true,
	"MAXA":      true,
	"MEDIAN":    true,
	"MIN":       true,
	"MINA":      true,
	"MODE":      true,
	"MODE.MULT": true,
	"MODE.SNGL": true,
	"PRODUCT":   true,
	"ROWS":      true,
	"SKEW":      true,
	"SKEW.P":    true,
	"STDEV":     true,
	"STDEV.P":   true,
	"STDEV.S":   true,
	"STDEVA":    true,
	"STDEVP":    true,
	"STDEVPA":   true,
	"SUM":       true,
	"SUMSQ":     true,
	"VAR":       true,
	"VAR.P":     true,
	"VAR.S":     true,
	"VARA":      true,
	"VARP":      true,
	"VARPA":     true,
}

// directRangeArgFuncs preserve direct range references only for the argument
// positions that are range-like for a given function. This avoids accidentally
// suppressing implicit intersection for scalar parameters such as the first
// STANDARDIZE argument.
var directRangeArgFuncs = map[string]map[int]bool{
	"COUNTBLANK":      {0: true},
	"CORREL":          {0: true, 1: true},
	"COVAR":           {0: true, 1: true},
	"COVARIANCE.P":    {0: true, 1: true},
	"COVARIANCE.S":    {0: true, 1: true},
	"FORECAST":        {1: true, 2: true},
	"FORECAST.LINEAR": {1: true, 2: true},
	"INTERCEPT":       {0: true, 1: true},
	"LARGE":           {0: true},
	"PEARSON":         {0: true, 1: true},
	"PERCENTILE":      {0: true},
	"PERCENTILE.EXC":  {0: true},
	"PERCENTILE.INC":  {0: true},
	"PERCENTRANK":     {0: true},
	"PERCENTRANK.EXC": {0: true},
	"PERCENTRANK.INC": {0: true},
	"QUARTILE":        {0: true},
	"QUARTILE.EXC":    {0: true},
	"QUARTILE.INC":    {0: true},
	"RANK":            {1: true},
	"RANK.AVG":        {1: true},
	"RANK.EQ":         {1: true},
	"RSQ":             {0: true, 1: true},
	"SLOPE":           {0: true, 1: true},
	"SMALL":           {0: true},
	"STEYX":           {0: true, 1: true},
	"TRIMMEAN":        {0: true},
	"XLOOKUP":         {1: true, 2: true},
	"XMATCH":          {1: true},
}

// directRangeArgStartFuncs preserve direct ranges from the given argument
// index onward. This is used for functions that accept a trailing list of
// references after one or more scalar control arguments.
var directRangeArgStartFuncs = map[string]int{
	"AGGREGATE": 2,
	"SUBTOTAL":  1,
}

// arrayArgFuncs evaluate the listed argument positions in array context,
// suppressing implicit intersection for the whole argument expression. This is
// required for functions like FILTER whose include argument is commonly a
// boolean range expression rather than a plain range reference.
var arrayArgFuncs = map[string]map[int]bool{
	"FILTER": {0: true, 1: true},
}

// arrayFirstArgFuncs evaluate the first argument in array context because it is
// semantically an array input.
var arrayFirstArgFuncs = map[string]bool{
	"ARRAYTOTEXT": true,
	"CHOOSECOLS":  true,
	"CHOOSEROWS":  true,
	"DROP":        true,
	"EXPAND":      true,
	"SORT":        true,
	"TAKE":        true,
	"TOCOL":       true,
	"TOROW":       true,
	"UNIQUE":      true,
	"WRAPCOLS":    true,
	"WRAPROWS":    true,
}

// arrayAllArgFuncs evaluate every argument in array context.
var arrayAllArgFuncs = map[string]bool{
	"HSTACK": true,
	"VSTACK": true,
}

// IsArrayFunc reports whether the named function forces array evaluation of
// its arguments. The compiler uses this to emit OpEnterArrayCtx / OpLeaveArrayCtx
// around the function's argument expressions.
func IsArrayFunc(name string) bool {
	return arrayForcingFuncs[strings.ToUpper(name)]
}

// ArgEvalModeForFuncArg reports how the compiler should evaluate the given
// argument position for a function in legacy non-array formula contexts.
func ArgEvalModeForFuncArg(name string, argIndex int) FuncArgEvalMode {
	upper := strings.ToUpper(name)
	if arrayAllArgFuncs[upper] {
		return FuncArgEvalArray
	}
	if arrayFirstArgFuncs[upper] && argIndex == 0 {
		return FuncArgEvalArray
	}
	if positions, ok := arrayArgFuncs[upper]; ok && positions[argIndex] {
		return FuncArgEvalArray
	}
	if upper == "SORTBY" && (argIndex == 0 || argIndex%2 == 1) {
		return FuncArgEvalArray
	}
	if directRangeAllArgFuncs[upper] {
		return FuncArgEvalDirectRange
	}
	if start, ok := directRangeArgStartFuncs[upper]; ok && argIndex >= start {
		return FuncArgEvalDirectRange
	}
	if positions, ok := directRangeArgFuncs[upper]; ok && positions[argIndex] {
		return FuncArgEvalDirectRange
	}
	return FuncArgEvalDefault
}

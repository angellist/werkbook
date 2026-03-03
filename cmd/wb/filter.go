package main

import (
	"fmt"
	"strconv"
	"strings"
)

// filterCondition represents a single --where condition.
type filterCondition struct {
	Column string // header name or column letter (A, B, ...)
	Op     string // =, !=, <, >, <=, >=
	Value  string
}

// parseWhere parses a string like "column=value" or "column>=42" into a filterCondition.
func parseWhere(s string) (filterCondition, error) {
	// Try operators in order of longest first to avoid matching "<" before "<=".
	for _, op := range []string{"<=", ">=", "!=", "=", "<", ">"} {
		idx := strings.Index(s, op)
		if idx > 0 {
			col := s[:idx]
			val := s[idx+len(op):]
			return filterCondition{Column: col, Op: op, Value: val}, nil
		}
	}
	return filterCondition{}, fmt.Errorf("invalid --where expression %q: expected column<op>value (operators: =, !=, <, >, <=, >=)", s)
}

// resolveColumnIndex maps a filter column reference to a 0-based index into the row slice.
// If headers are provided, matches by header name (case-insensitive).
// Otherwise, matches by column letter (A, B, ...).
func resolveColumnIndex(col string, headers []string, col1 int) (int, error) {
	if len(headers) > 0 {
		lower := strings.ToLower(col)
		for i, h := range headers {
			if strings.ToLower(h) == lower {
				return i, nil
			}
		}
		return -1, fmt.Errorf("column %q not found in headers", col)
	}
	// Match by column letter.
	upper := strings.ToUpper(col)
	for i := 0; ; i++ {
		letter := columnNumberToLetter(col1 + i)
		if letter == upper {
			return i, nil
		}
		if col1+i > 16384 {
			break
		}
	}
	return -1, fmt.Errorf("column %q not found", col)
}

// columnNumberToLetter converts a 1-based column number to a letter (1->A, 2->B, ...).
func columnNumberToLetter(col int) string {
	var result []byte
	for col > 0 {
		col--
		result = append([]byte{byte('A' + col%26)}, result...)
		col /= 26
	}
	return string(result)
}

// matchesFilters checks whether a row (as a slice of string values) passes all filter conditions.
func matchesFilters(row []string, filters []resolvedFilter) bool {
	for _, f := range filters {
		if f.colIdx < 0 || f.colIdx >= len(row) {
			return false
		}
		if !compareValues(row[f.colIdx], f.cond.Op, f.cond.Value) {
			return false
		}
	}
	return true
}

type resolvedFilter struct {
	cond   filterCondition
	colIdx int
}

// compareValues compares cellVal against filterVal using op.
// If both parse as floats, numeric comparison is used.
// Otherwise, string comparison (case-insensitive for = and !=).
func compareValues(cellVal, op, filterVal string) bool {
	cellNum, cellErr := strconv.ParseFloat(cellVal, 64)
	filterNum, filterErr := strconv.ParseFloat(filterVal, 64)

	if cellErr == nil && filterErr == nil {
		return compareNumeric(cellNum, op, filterNum)
	}
	return compareString(cellVal, op, filterVal)
}

func compareNumeric(a float64, op string, b float64) bool {
	switch op {
	case "=":
		return a == b
	case "!=":
		return a != b
	case "<":
		return a < b
	case ">":
		return a > b
	case "<=":
		return a <= b
	case ">=":
		return a >= b
	}
	return false
}

func compareString(a, op, b string) bool {
	switch op {
	case "=":
		return strings.EqualFold(a, b)
	case "!=":
		return !strings.EqualFold(a, b)
	case "<":
		return a < b
	case ">":
		return a > b
	case "<=":
		return a <= b
	case ">=":
		return a >= b
	}
	return false
}

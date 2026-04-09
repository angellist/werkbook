package werkbook

import (
	"fmt"
	"strings"

	"github.com/jpoz/werkbook/formula"
)

// DefinedName represents a defined name, including workbook-scoped
// and sheet-scoped named ranges.
type DefinedName struct {
	Name         string
	Value        string
	LocalSheetID int // -1 for workbook scope; otherwise 0-based sheet index
}

// DefinedNames returns the workbook's defined names in file order.
func (f *File) DefinedNames() []DefinedName {
	out := make([]DefinedName, len(f.definedNames))
	for i, dn := range f.definedNames {
		out[i] = DefinedName{
			Name:         dn.Name,
			Value:        dn.Value,
			LocalSheetID: dn.LocalSheetID,
		}
	}
	return out
}

// AddDefinedName adds a new defined name to the workbook. Use LocalSheetID -1
// for workbook scope or a 0-based sheet index for sheet scope.
func (f *File) AddDefinedName(dn DefinedName) {
	if err := f.validateDefinedName(dn); err != nil {
		return
	}
	f.definedNames = append(f.definedNames, formula.DefinedNameInfo{
		Name:         dn.Name,
		Value:        dn.Value,
		LocalSheetID: dn.LocalSheetID,
	})
	f.rebuildFormulaState()
}

// SetDefinedName inserts or replaces a defined name with the same name and scope.
func (f *File) SetDefinedName(dn DefinedName) error {
	if err := f.validateDefinedName(dn); err != nil {
		return err
	}
	lower := strings.ToLower(dn.Name)
	for i, existing := range f.definedNames {
		if strings.ToLower(existing.Name) == lower && existing.LocalSheetID == dn.LocalSheetID {
			f.definedNames[i] = formula.DefinedNameInfo{
				Name:         dn.Name,
				Value:        dn.Value,
				LocalSheetID: dn.LocalSheetID,
			}
			f.rebuildFormulaState()
			return nil
		}
	}
	f.definedNames = append(f.definedNames, formula.DefinedNameInfo{
		Name:         dn.Name,
		Value:        dn.Value,
		LocalSheetID: dn.LocalSheetID,
	})
	f.rebuildFormulaState()
	return nil
}

// DeleteDefinedName removes the first defined name matching name and scope.
// Returns an error if no matching name is found.
func (f *File) DeleteDefinedName(name string, localSheetID int) error {
	lower := strings.ToLower(name)
	for i, dn := range f.definedNames {
		if strings.ToLower(dn.Name) == lower && dn.LocalSheetID == localSheetID {
			f.definedNames = append(f.definedNames[:i], f.definedNames[i+1:]...)
			f.rebuildFormulaState()
			return nil
		}
	}
	return fmt.Errorf("defined name %q not found", name)
}

func (f *File) validateDefinedName(dn DefinedName) error {
	if dn.LocalSheetID < -1 || dn.LocalSheetID >= len(f.sheets) {
		return fmt.Errorf("defined name %q has invalid local sheet index %d", dn.Name, dn.LocalSheetID)
	}
	return nil
}

// ResolveDefinedName looks up a defined name by its name and returns the
// resolved cell values as a 2D grid. For a single-cell reference the result
// is a 1x1 grid. The lookup is case-insensitive. If sheetIndex >= 0, a
// sheet-scoped name for that sheet takes precedence over a global name.
// Pass -1 for sheetIndex to match only workbook-scoped names.
func (f *File) ResolveDefinedName(name string, sheetIndex int) ([][]Value, error) {
	ref, err := f.lookupDefinedName(name, sheetIndex)
	if err != nil {
		return nil, err
	}

	sheetName, cellRef, err := parseDefinedNameRef(ref)
	if err != nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: %w", name, err)
	}

	s := f.Sheet(sheetName)
	if s == nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: sheet %q not found", name, sheetName)
	}

	area, err := parseDefinedNameArea(sheetName, cellRef)
	if err != nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: %w", name, err)
	}
	if area.isRange {
		return f.resolveDefinedNameRange(area.rangeAddr), nil
	}

	// Single cell.
	singleRef, err := CoordinatesToCellName(area.cellCol, area.cellRow)
	if err != nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: %w", name, err)
	}
	v, err := s.GetValue(singleRef)
	if err != nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: %w", name, err)
	}
	return [][]Value{{v}}, nil
}

type definedNameArea struct {
	rangeAddr formula.RangeAddr
	cellCol   int
	cellRow   int
	isRange   bool
}

func parseDefinedNameArea(sheetName, cellRef string) (definedNameArea, error) {
	parts := strings.SplitN(strings.TrimSpace(cellRef), ":", 2)
	if len(parts) == 2 {
		fromCol, fromRow, err := parseDefinedNameCoord(parts[0])
		if err != nil {
			return definedNameArea{}, err
		}
		toCol, toRow, err := parseDefinedNameCoord(parts[1])
		if err != nil {
			return definedNameArea{}, err
		}
		addr, err := buildDefinedNameRangeAddr(sheetName, fromCol, fromRow, toCol, toRow)
		if err != nil {
			return definedNameArea{}, err
		}
		return definedNameArea{rangeAddr: addr, isRange: true}, nil
	}

	col, row, err := parseDefinedNameCoord(parts[0])
	if err != nil {
		return definedNameArea{}, err
	}
	switch {
	case col > 0 && row > 0:
		return definedNameArea{cellCol: col, cellRow: row}, nil
	case col > 0:
		return definedNameArea{
			rangeAddr: formula.RangeAddr{
				Sheet:   sheetName,
				FromCol: col,
				FromRow: 1,
				ToCol:   col,
				ToRow:   MaxRows,
			},
			isRange: true,
		}, nil
	case row > 0:
		return definedNameArea{
			rangeAddr: formula.RangeAddr{
				Sheet:   sheetName,
				FromCol: 1,
				FromRow: row,
				ToCol:   MaxColumns,
				ToRow:   row,
			},
			isRange: true,
		}, nil
	default:
		return definedNameArea{}, fmt.Errorf("invalid reference %q", cellRef)
	}
}

func buildDefinedNameRangeAddr(sheetName string, fromCol, fromRow, toCol, toRow int) (formula.RangeAddr, error) {
	switch {
	case fromCol > 0 && fromRow > 0 && toCol > 0 && toRow > 0:
		if fromCol > toCol {
			fromCol, toCol = toCol, fromCol
		}
		if fromRow > toRow {
			fromRow, toRow = toRow, fromRow
		}
	case fromCol > 0 && fromRow == 0 && toCol > 0 && toRow == 0:
		if fromCol > toCol {
			fromCol, toCol = toCol, fromCol
		}
		fromRow = 1
		toRow = MaxRows
	case fromCol == 0 && fromRow > 0 && toCol == 0 && toRow > 0:
		if fromRow > toRow {
			fromRow, toRow = toRow, fromRow
		}
		fromCol = 1
		toCol = MaxColumns
	default:
		return formula.RangeAddr{}, fmt.Errorf("mixed reference types in %q", fmt.Sprintf("%s:%s", formatDefinedNameCoord(fromCol, fromRow), formatDefinedNameCoord(toCol, toRow)))
	}
	return formula.RangeAddr{
		Sheet:   sheetName,
		FromCol: fromCol,
		FromRow: fromRow,
		ToCol:   toCol,
		ToRow:   toRow,
	}, nil
}

func parseDefinedNameCoord(ref string) (col, row int, err error) {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return 0, 0, fmt.Errorf("empty reference")
	}

	i := 0
	for i < len(ref) && isAlpha(ref[i]) {
		i++
	}
	if i == 0 {
		row, err = parseRow(ref)
		if err != nil {
			return 0, 0, fmt.Errorf("invalid reference %q: %w", ref, err)
		}
		return 0, row, nil
	}

	col, err = ColumnNameToNumber(ref[:i])
	if err != nil {
		return 0, 0, fmt.Errorf("invalid reference %q: %w", ref, err)
	}
	if i == len(ref) {
		return col, 0, nil
	}
	row, err = parseRow(ref[i:])
	if err != nil {
		return 0, 0, fmt.Errorf("invalid reference %q: %w", ref, err)
	}
	return col, row, nil
}

func formatDefinedNameCoord(col, row int) string {
	switch {
	case col > 0 && row > 0:
		ref, err := CoordinatesToCellName(col, row)
		if err == nil {
			return ref
		}
	case col > 0:
		return ColumnNumberToName(col)
	case row > 0:
		return fmt.Sprintf("%d", row)
	}
	return ""
}

func (f *File) resolveDefinedNameRange(addr formula.RangeAddr) [][]Value {
	resolver := &fileResolver{file: f, currentSheet: addr.Sheet}
	rawRows := resolver.GetRangeValues(addr)
	if !isDefinedNameOpenEndedRange(addr) {
		expectedRows := addr.ToRow - addr.FromRow + 1
		expectedCols := addr.ToCol - addr.FromCol + 1
		for len(rawRows) < expectedRows {
			blankRow := make([]formula.Value, expectedCols)
			for i := range blankRow {
				blankRow[i] = formula.EmptyVal()
			}
			rawRows = append(rawRows, blankRow)
		}
	}
	out := make([][]Value, len(rawRows))
	for i, rawRow := range rawRows {
		out[i] = make([]Value, len(rawRow))
		for j, raw := range rawRow {
			out[i][j] = formulaGridValueToCellValue(raw)
		}
	}
	return out
}

func isDefinedNameOpenEndedRange(addr formula.RangeAddr) bool {
	return (addr.FromRow == 1 && addr.ToRow >= MaxRows) ||
		(addr.FromCol == 1 && addr.ToCol >= MaxColumns)
}

func formulaGridValueToCellValue(v formula.Value) Value {
	switch v.Type {
	case formula.ValueNumber:
		return Value{Type: TypeNumber, Number: v.Num}
	case formula.ValueString:
		return Value{Type: TypeString, String: v.Str}
	case formula.ValueBool:
		return Value{Type: TypeBool, Bool: v.Bool}
	case formula.ValueError:
		return Value{Type: TypeError, String: v.Err.String()}
	default:
		return Value{Type: TypeEmpty}
	}
}

// lookupDefinedName finds the best-matching defined name, preferring a
// sheet-scoped name for sheetIndex over a global name.
func (f *File) lookupDefinedName(name string, sheetIndex int) (string, error) {
	lower := strings.ToLower(name)
	var globalVal string
	globalFound := false
	for _, dn := range f.definedNames {
		if strings.ToLower(dn.Name) != lower {
			continue
		}
		if dn.LocalSheetID == sheetIndex {
			return dn.Value, nil
		}
		if dn.LocalSheetID == -1 && !globalFound {
			globalVal = dn.Value
			globalFound = true
		}
	}
	if globalFound {
		return globalVal, nil
	}
	return "", fmt.Errorf("defined name %q not found", name)
}

// parseDefinedNameRef parses a defined name value like "Sheet1!$A$1:$C$10"
// into a sheet name and a cell/range reference with $ signs stripped.
func parseDefinedNameRef(ref string) (sheetName, cellRef string, err error) {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return "", "", fmt.Errorf("empty reference")
	}

	idx := strings.LastIndex(ref, "!")
	if idx < 0 {
		return "", "", fmt.Errorf("reference %q has no sheet qualifier", ref)
	}

	sheetName = ref[:idx]
	cellRef = ref[idx+1:]

	// Strip surrounding quotes from sheet name (e.g. 'My Sheet'!A1).
	if len(sheetName) >= 2 && sheetName[0] == '\'' && sheetName[len(sheetName)-1] == '\'' {
		sheetName = strings.ReplaceAll(sheetName[1:len(sheetName)-1], "''", "'")
	}

	// Strip $ signs from the cell reference.
	cellRef = strings.ReplaceAll(cellRef, "$", "")

	if sheetName == "" || cellRef == "" {
		return "", "", fmt.Errorf("invalid reference %q", ref)
	}
	return sheetName, cellRef, nil
}

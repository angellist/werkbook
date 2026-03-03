package main

import (
	"fmt"
	"os"

	werkbook "github.com/werkbook/werkbook"
)

type readData struct {
	File    string    `json:"file"`
	Sheet   string    `json:"sheet"`
	Range   string    `json:"range"`
	Headers []string  `json:"headers,omitempty"`
	Rows    []rowData `json:"rows"`
}

type rowData struct {
	Row   int                `json:"row"`
	Cells map[string]cellData `json:"cells"`
}

type cellData struct {
	Value      any    `json:"value"`
	Type       string `json:"type"`
	Formatted  string `json:"formatted,omitempty"`
	Formula    string `json:"formula,omitempty"`
	HasFormula bool   `json:"has_formula,omitempty"`
	Style      any    `json:"style,omitempty"`
}

func cmdRead(args []string, globals globalFlags) int {
	cmd := "read"

	if hasHelpFlag(args) {
		fmt.Fprintln(os.Stderr, `Usage: werkbook read [flags] <file>

Read cell data from a workbook. Returns stored/cached values.
Use 'calc' to force formula recalculation.

Flags:
  --sheet <name>        Read from the named sheet (default: first sheet)
  --range <A1:D10>      Read a specific range (default: full used range)
  --include-formulas    Include formula strings in output
  --include-styles      Include style objects in output
  --headers             Treat first row as headers

Date cells are automatically detected and returned with type "date"
and a "formatted" field containing an ISO 8601 date string.

Examples:
  werkbook read data.xlsx
  werkbook read --range A1:C10 data.xlsx
  werkbook read --headers --include-formulas data.xlsx
  werkbook read --format csv --headers data.xlsx`)
		return ExitSuccess
	}

	var sheetFlag, rangeFlag string
	var includeFormulas, includeStyles, headersFlag bool

	i := 0
	var filePath string
	for i < len(args) {
		switch args[i] {
		case "--sheet":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--sheet requires a value"), globals)
				return ExitUsage
			}
			sheetFlag = args[i+1]
			i += 2
		case "--range":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--range requires a value"), globals)
				return ExitUsage
			}
			rangeFlag = args[i+1]
			i += 2
		case "--include-formulas":
			includeFormulas = true
			i++
		case "--include-styles":
			includeStyles = true
			i++
		case "--headers":
			headersFlag = true
			i++
		default:
			if filePath == "" && len(args[i]) > 0 && args[i][0] != '-' {
				filePath = args[i]
				i++
			} else {
				writeError(cmd, errUsage("unknown flag: "+args[i]), globals)
				return ExitUsage
			}
		}
	}

	if filePath == "" {
		writeError(cmd, errUsage("file path required"), globals)
		return ExitUsage
	}

	f, err := werkbook.Open(filePath)
	if err != nil {
		if os.IsNotExist(err) {
			writeError(cmd, errFileNotFound(filePath, err), globals)
		} else {
			writeError(cmd, errFileOpen(filePath, err), globals)
		}
		return ExitFileIO
	}

	// Resolve sheet.
	sheetName := sheetFlag
	if sheetName == "" {
		names := f.SheetNames()
		if len(names) == 0 {
			writeError(cmd, errInternal(fmt.Errorf("workbook has no sheets")), globals)
			return ExitInternal
		}
		sheetName = names[0]
	}

	s := f.Sheet(sheetName)
	if s == nil {
		writeError(cmd, errSheetNotFound(sheetName), globals)
		return ExitValidate
	}

	// Determine range.
	var col1, row1, col2, row2 int
	if rangeFlag != "" {
		col1, row1, col2, row2, err = werkbook.RangeToCoordinates(rangeFlag)
		if err != nil {
			writeError(cmd, errInvalidRange(rangeFlag, err), globals)
			return ExitValidate
		}
	} else {
		maxRow := s.MaxRow()
		maxCol := s.MaxCol()
		if maxRow == 0 || maxCol == 0 {
			// Empty sheet.
			data := readData{File: filePath, Sheet: sheetName, Rows: []rowData{}}
			writeSuccess(cmd, data, globals)
			return ExitSuccess
		}
		col1, row1, col2, row2 = 1, 1, maxCol, maxRow
	}

	rangeStr := rangeFlag
	if rangeStr == "" {
		start, _ := werkbook.CoordinatesToCellName(col1, row1)
		end, _ := werkbook.CoordinatesToCellName(col2, row2)
		rangeStr = start + ":" + end
	}

	// Read headers from row1 if requested.
	var headers []string
	if headersFlag {
		for c := col1; c <= col2; c++ {
			ref, _ := werkbook.CoordinatesToCellName(c, row1)
			v, _ := s.GetValue(ref)
			headers = append(headers, valueToString(v))
		}
	}

	startRow := row1
	if headersFlag {
		startRow = row1 + 1
	}

	// Handle non-JSON formats without building the JSON row data.
	if globals.format == FormatMarkdown || globals.format == FormatCSV {
		var tableRows [][]string
		for r := startRow; r <= row2; r++ {
			var row []string
			for c := col1; c <= col2; c++ {
				ref, _ := werkbook.CoordinatesToCellName(c, r)
				v, _ := s.GetValue(ref)
				if v.Type == werkbook.TypeNumber && isDateCell(s, ref) {
					row = append(row, werkbook.ExcelSerialToTime(v.Number).Format("2006-01-02"))
				} else {
					row = append(row, valueToString(v))
				}
			}
			tableRows = append(tableRows, row)
		}
		output := formatTable(globals.format, headers, tableRows)
		fmt.Print(output)
		return ExitSuccess
	}

	// Build row data for JSON output.
	var rows []rowData
	for r := startRow; r <= row2; r++ {
		cells := make(map[string]cellData)
		for c := col1; c <= col2; c++ {
			ref, _ := werkbook.CoordinatesToCellName(c, r)
			v, _ := s.GetValue(ref)
			if v.IsEmpty() {
				continue
			}

			cd := cellData{
				Value: v.Raw(),
				Type:  valueTypeName(v),
			}

			// Detect date cells.
			if v.Type == werkbook.TypeNumber {
				if isDateCell(s, ref) {
					cd.Type = "date"
					cd.Formatted = werkbook.ExcelSerialToTime(v.Number).Format("2006-01-02")
				}
			}

			// Always signal if the cell has a formula.
			formula, _ := s.GetFormula(ref)
			if formula != "" {
				cd.HasFormula = true
				if includeFormulas {
					cd.Formula = formula
				}
			}

			if includeStyles {
				style, _ := s.GetStyle(ref)
				if style != nil {
					cd.Style = styleToJSON(style)
				}
			}

			cells[ref] = cd
		}

		if len(cells) > 0 {
			rows = append(rows, rowData{Row: r, Cells: cells})
		}
	}

	if rows == nil {
		rows = []rowData{}
	}

	data := readData{
		File:    filePath,
		Sheet:   sheetName,
		Range:   rangeStr,
		Headers: headers,
		Rows:    rows,
	}

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

func valueToString(v werkbook.Value) string {
	switch v.Type {
	case werkbook.TypeNumber:
		if v.Number == float64(int64(v.Number)) {
			return fmt.Sprintf("%d", int64(v.Number))
		}
		return fmt.Sprintf("%g", v.Number)
	case werkbook.TypeString:
		return v.String
	case werkbook.TypeBool:
		if v.Bool {
			return "TRUE"
		}
		return "FALSE"
	case werkbook.TypeError:
		return v.String
	default:
		return ""
	}
}

func valueTypeName(v werkbook.Value) string {
	switch v.Type {
	case werkbook.TypeNumber:
		return "number"
	case werkbook.TypeString:
		return "string"
	case werkbook.TypeBool:
		return "bool"
	case werkbook.TypeError:
		return "error"
	default:
		return "empty"
	}
}

// isDateCell checks whether the cell's style indicates a date number format.
func isDateCell(s *werkbook.Sheet, ref string) bool {
	style, _ := s.GetStyle(ref)
	if style == nil {
		return false
	}
	return werkbook.IsDateFormat(style.NumFmt, style.NumFmtID)
}

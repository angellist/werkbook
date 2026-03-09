package formula

import (
	"math"
	"testing"
)

// makeDBResolver builds a mockResolver with a database in A1:E11 style layout
// and optional criteria cells. The database argument is a [][]Value where the
// first row contains headers. Criteria are placed starting at the given cell
// address offset.
func makeDBResolver(db [][]Value, dbStartCol, dbStartRow int, crit [][]Value, critStartCol, critStartRow int) *mockResolver {
	cells := make(map[CellAddr]Value)
	for r, row := range db {
		for c, v := range row {
			cells[CellAddr{Col: dbStartCol + c, Row: dbStartRow + r}] = v
		}
	}
	for r, row := range crit {
		for c, v := range row {
			cells[CellAddr{Col: critStartCol + c, Row: critStartRow + r}] = v
		}
	}
	return &mockResolver{cells: cells}
}

// dbRange returns a range string like "A1:D5" for a database grid.
func dbRange(startCol, startRow, numCols, numRows int) string {
	return cellName(startCol, startRow) + ":" + cellName(startCol+numCols-1, startRow+numRows-1)
}

func cellName(col, row int) string {
	c := string(rune('A' + col - 1))
	return c + dbItoa(row)
}

func dbItoa(n int) string {
	s := ""
	for n > 0 {
		s = string(rune('0'+n%10)) + s
		n /= 10
	}
	if s == "" {
		return "0"
	}
	return s
}

func TestDSUM(t *testing.T) {
	// Standard test database (similar to Excel documentation):
	// | Tree    | Height | Age | Yield | Profit |
	// | Apple   | >10    | ... | ...   | ...    |
	// | Pear    | ...    | ... | ...   | ...    |
	// etc.
	//
	// Database in A5:E11 (rows 5-11, cols A-E)
	// Criteria in various locations

	// Database: columns = Tree, Height, Age, Yield, Profit
	db := [][]Value{
		{StringVal("Tree"), StringVal("Height"), StringVal("Age"), StringVal("Yield"), StringVal("Profit")},
		{StringVal("Apple"), NumberVal(18), NumberVal(20), NumberVal(14), NumberVal(105)},
		{StringVal("Pear"), NumberVal(12), NumberVal(12), NumberVal(10), NumberVal(96)},
		{StringVal("Cherry"), NumberVal(13), NumberVal(14), NumberVal(9), NumberVal(105)},
		{StringVal("Apple"), NumberVal(14), NumberVal(15), NumberVal(10), NumberVal(75)},
		{StringVal("Pear"), NumberVal(9), NumberVal(8), NumberVal(8), NumberVal(76.8)},
		{StringVal("Apple"), NumberVal(8), NumberVal(9), NumberVal(6), NumberVal(45)},
	}

	tests := []struct {
		name     string
		db       [][]Value
		crit     [][]Value
		field    string // field as string in the formula
		wantType ValueType
		wantNum  float64
		wantErr  ErrorValue
	}{
		{
			name: "basic DSUM with string field - Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  225, // 105 + 75 + 45
		},
		{
			name: "DSUM with numeric field index",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5", // 5th column = Profit
			wantType: ValueNumber,
			wantNum:  225,
		},
		{
			name: "DSUM with Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  172.8, // 96 + 76.8
		},
		{
			name: "multiple criteria AND - Apple with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  180, // 105 (h=18) + 75 (h=14)
		},
		{
			name: "multiple criteria rows OR - Apple OR Pear",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  397.8, // 225 + 172.8
		},
		{
			name: "numeric comparison >10 on Height, sum Yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  43, // 14+10+9+10 (heights 18,12,13,14)
		},
		{
			name: "numeric comparison <10 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  121.8, // 76.8 (h=9) + 45 (h=8)
		},
		{
			name: "numeric comparison >=14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  180, // 105 (h=18) + 75 (h=14)
		},
		{
			name: "numeric comparison <=12 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<=12")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  217.8, // 96 (h=12) + 76.8 (h=9) + 45 (h=8)
		},
		{
			name: "numeric comparison <>13 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("<>13")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  397.8, // all except Cherry (105) = 105+96+75+76.8+45
		},
		{
			name: "exact numeric match =14 on Height",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal("=14")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // only Apple h=14
		},
		{
			name: "text case-insensitive match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  225,
		},
		{
			name: "blank criteria matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")}, // blank = match all
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8, // sum of all profits
		},
		{
			name: "no criteria rows matches all",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				// no condition rows
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  502.8,
		},
		{
			name: "no matching records returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Orange")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name:     "field name not found returns #VALUE!",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"NonExistent"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index out of range returns #VALUE!",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "99",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "field index 0 returns #VALUE!",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    "0",
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "wildcard * in criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("A*")}, // matches Apple
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  225,
		},
		{
			name: "wildcard ? in criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pea?")}, // matches Pear
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  172.8,
		},
		{
			name: "mixed types in field column - only sum numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), EmptyVal()},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")}, // match all
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  30, // 10 + 20 (text, bool, empty ignored)
		},
		{
			name: "multiple AND criteria with text and numeric",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Age")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  180, // Apple with age>10: age 20 (profit 105) + age 15 (profit 75)
		},
		{
			name: "complex OR criteria - Apple with height>10 OR Pear with height<10",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
				{StringVal("Pear"), StringVal("<10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  256.8, // Apple h>10: 105+75=180, Pear h<10: 76.8 → 256.8
		},
		{
			name: "empty database returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				// no data rows
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "exact match with = prefix",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("Apple"), NumberVal(10)},
				{StringVal("Apple Pie"), NumberVal(20)},
				{StringVal("apple"), NumberVal(30)},
			},
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("=Apple")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  40, // "Apple" and "apple" match (case-insensitive), not "Apple Pie"
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// Place database at A1
			dbRows := len(tt.db)
			dbCols := len(tt.db[0])
			// Place criteria at col G (7), row 1
			critRows := len(tt.crit)
			critCols := len(tt.crit[0])

			resolver := makeDBResolver(tt.db, 1, 1, tt.crit, 7, 1)

			formula := "DSUM(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tt.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"

			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", formula, err)
			}

			if got.Type != tt.wantType {
				t.Fatalf("DSUM type = %v, want %v (value: %+v)", got.Type, tt.wantType, got)
			}

			switch tt.wantType {
			case ValueNumber:
				if diff := got.Num - tt.wantNum; diff > 1e-9 || diff < -1e-9 {
					t.Errorf("DSUM = %g, want %g", got.Num, tt.wantNum)
				}
			case ValueError:
				if got.Err != tt.wantErr {
					t.Errorf("DSUM error = %v, want %v", got.Err, tt.wantErr)
				}
			}
		})
	}
}

func TestDSUM_WrongArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few arguments
	cf := evalCompile(t, "DSUM(A1:B2,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// With only 2 args the function should return an error.
	// Note: the parser may not pass exactly 2 args depending on how
	// the formula is compiled. We test the function directly too.
	_ = got

	// Direct function call with wrong arg counts
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDSum(tt.args)
			if err != nil {
				t.Fatalf("fnDSum error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDSum(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

func TestDSUM_ErrorPropagation(t *testing.T) {
	// If the database contains an error in the summed field, propagate it.
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")}, // match all
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	formula := "DSUM(A1:B4,\"Value\",G1:G2)"
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("DSUM with error cell = %+v, want #DIV/0!", got)
	}
}

func TestDSUM_FieldCaseInsensitive(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("VALUE")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)
	// Field name uses different case
	formula := `DSUM(A1:B3,"value",G1:G2)`
	cf := evalCompile(t, formula)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("DSUM case-insensitive field = %+v, want 30", got)
	}
}

// ---------------------------------------------------------------------------
// Helper: run a D-function table-driven test
// ---------------------------------------------------------------------------

// dbTestCase is shared by all D-function table tests.
type dbTestCase struct {
	name     string
	db       [][]Value
	crit     [][]Value
	field    string
	wantType ValueType
	wantNum  float64
	wantStr  string
	wantBool bool
	wantErr  ErrorValue
}

// runDBTests runs table-driven tests for a D-function.
func runDBTests(t *testing.T, funcName string, tests []dbTestCase) {
	t.Helper()
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dbRows := len(tt.db)
			dbCols := len(tt.db[0])
			critRows := len(tt.crit)
			critCols := len(tt.crit[0])

			resolver := makeDBResolver(tt.db, 1, 1, tt.crit, 7, 1)

			formula := funcName + "(" +
				dbRange(1, 1, dbCols, dbRows) + "," +
				tt.field + "," +
				dbRange(7, 1, critCols, critRows) + ")"

			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", formula, err)
			}

			if got.Type != tt.wantType {
				t.Fatalf("%s type = %v, want %v (value: %+v)", funcName, got.Type, tt.wantType, got)
			}

			switch tt.wantType {
			case ValueNumber:
				if diff := got.Num - tt.wantNum; diff > 1e-9 || diff < -1e-9 {
					t.Errorf("%s = %g, want %g", funcName, got.Num, tt.wantNum)
				}
			case ValueError:
				if got.Err != tt.wantErr {
					t.Errorf("%s error = %v, want %v", funcName, got.Err, tt.wantErr)
				}
			case ValueString:
				if got.Str != tt.wantStr {
					t.Errorf("%s = %q, want %q", funcName, got.Str, tt.wantStr)
				}
			case ValueBool:
				if got.Bool != tt.wantBool {
					t.Errorf("%s = %v, want %v", funcName, got.Bool, tt.wantBool)
				}
			}
		})
	}
}

// standardDB is the shared test database used across D-function tests.
func standardDB() [][]Value {
	return [][]Value{
		{StringVal("Tree"), StringVal("Height"), StringVal("Age"), StringVal("Yield"), StringVal("Profit")},
		{StringVal("Apple"), NumberVal(18), NumberVal(20), NumberVal(14), NumberVal(105)},
		{StringVal("Pear"), NumberVal(12), NumberVal(12), NumberVal(10), NumberVal(96)},
		{StringVal("Cherry"), NumberVal(13), NumberVal(14), NumberVal(9), NumberVal(105)},
		{StringVal("Apple"), NumberVal(14), NumberVal(15), NumberVal(10), NumberVal(75)},
		{StringVal("Pear"), NumberVal(9), NumberVal(8), NumberVal(8), NumberVal(76.8)},
		{StringVal("Apple"), NumberVal(8), NumberVal(9), NumberVal(6), NumberVal(45)},
	}
}

// ---------------------------------------------------------------------------
// DAVERAGE
// ---------------------------------------------------------------------------

func TestDAVERAGE(t *testing.T) {
	db := standardDB()

	mixedDB := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), StringVal("text")},
		{StringVal("C"), NumberVal(20)},
		{StringVal("D"), BoolVal(true)},
		{StringVal("E"), EmptyVal()},
	}

	tests := []dbTestCase{
		{
			name: "average Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  75, // (105+75+45)/3
		},
		{
			name: "average Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  86.4, // (96+76.8)/2
		},
		{
			name: "average all yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  57.0 / 6.0, // (14+10+9+10+8+6)/6 = 9.5
		},
		{
			name: "average with numeric criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  95.25, // (105+96+105+75)/4
		},
		{
			name: "average single matching record",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name:     "average no matching records returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "average with mixed types - only numeric",
			db:   mixedDB,
			crit: [][]Value{
				{StringVal("Name")},
				{StringVal("")},
			},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  15, // (10+20)/2
		},
		{
			name: "average with numeric field index",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    "5",
			wantType: ValueNumber,
			wantNum:  75,
		},
		{
			name: "average with AND criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal(">10")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  90, // (105+75)/2
		},
		{
			name: "average with OR criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  82.5, // (105+75+45+105)/4
		},
		{
			name:     "average empty database returns DIV/0",
			db:       [][]Value{{StringVal("Name"), StringVal("Value")}},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "average all text values returns DIV/0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), StringVal("x")},
				{StringVal("B"), StringVal("y")},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "average height of all trees",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  74.0 / 6.0, // (18+12+13+14+9+8)/6
		},
		{
			name:     "average field not found",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"Missing"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name: "average wildcard criteria",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("P*")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  86.4, // (96+76.8)/2
		},
	}

	runDBTests(t, "DAVERAGE", tests)
}

func TestDAVERAGE_WrongArgCount(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{"zero args", nil},
		{"one arg", []Value{NumberVal(1)}},
		{"two args", []Value{NumberVal(1), NumberVal(2)}},
		{"four args", []Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := fnDAverage(tt.args)
			if err != nil {
				t.Fatalf("fnDAverage error: %v", err)
			}
			if result.Type != ValueError || result.Err != ErrValVALUE {
				t.Errorf("fnDAverage(%d args) = %+v, want #VALUE!", len(tt.args), result)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// DCOUNT
// ---------------------------------------------------------------------------

func TestDCOUNT(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "count Apple profit values",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		{
			name: "count all profit values",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		{
			name: "count mixed types - only numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), EmptyVal()},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  2, // only 10, 20
		},
		{
			name:     "count no matches returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "count with Height>10",
			db:   db,
			crit: [][]Value{
				{StringVal("Height")},
				{StringVal(">10")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  4,
		},
	}

	runDBTests(t, "DCOUNT", tests)
}

func TestDCOUNT_WrongArgCount(t *testing.T) {
	result, err := fnDCount(nil)
	if err != nil {
		t.Fatalf("fnDCount error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDCount(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DCOUNTA
// ---------------------------------------------------------------------------

func TestDCOUNTA(t *testing.T) {
	tests := []dbTestCase{
		{
			name: "counta mixed types - count non-empty",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(20)},
				{StringVal("D"), BoolVal(true)},
				{StringVal("E"), EmptyVal()},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  4, // 10, "text", 20, true (not empty)
		},
		{
			name: "counta all numbers",
			db:   standardDB(),
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  3,
		},
		{
			name: "counta tree column (strings)",
			db:   standardDB(),
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Tree"`,
			wantType: ValueNumber,
			wantNum:  6,
		},
		{
			name: "counta no matches returns 0",
			db:   standardDB(),
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Orange")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "counta all empty returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), EmptyVal()},
				{StringVal("B"), EmptyVal()},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
	}

	runDBTests(t, "DCOUNTA", tests)
}

func TestDCOUNTA_WrongArgCount(t *testing.T) {
	result, err := fnDCountA(nil)
	if err != nil {
		t.Fatalf("fnDCountA error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDCountA(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DGET
// ---------------------------------------------------------------------------

func TestDGET(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "dget single match returns value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name:     "dget no matches returns VALUE error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "dget multiple matches returns NUM error",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Apple")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValNUM,
		},
		{
			name: "dget returns string value",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Cherry"), StringVal("=13")},
			},
			field:    `"Tree"`,
			wantType: ValueString,
			wantStr:  "Cherry",
		},
		{
			name: "dget with specific AND criteria yields one match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree"), StringVal("Height")},
				{StringVal("Apple"), StringVal("=18")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
	}

	runDBTests(t, "DGET", tests)
}

func TestDGET_WrongArgCount(t *testing.T) {
	result, err := fnDGet(nil)
	if err != nil {
		t.Fatalf("fnDGet error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDGet(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DMAX
// ---------------------------------------------------------------------------

func TestDMAX(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "max Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  105,
		},
		{
			name: "max all height",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  18,
		},
		{
			name:     "max no matches returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "max Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name: "max with negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-5)},
				{StringVal("C"), NumberVal(-20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -5,
		},
	}

	runDBTests(t, "DMAX", tests)
}

func TestDMAX_WrongArgCount(t *testing.T) {
	result, err := fnDMax(nil)
	if err != nil {
		t.Fatalf("fnDMax error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDMax(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DMIN
// ---------------------------------------------------------------------------

func TestDMIN(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "min Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  45,
		},
		{
			name: "min all height",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("")},
			},
			field:    `"Height"`,
			wantType: ValueNumber,
			wantNum:  8,
		},
		{
			name:     "min no matches returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "min Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  8,
		},
		{
			name: "min with negative numbers",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(-10)},
				{StringVal("B"), NumberVal(-5)},
				{StringVal("C"), NumberVal(-20)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  -20,
		},
	}

	runDBTests(t, "DMIN", tests)
}

func TestDMIN_WrongArgCount(t *testing.T) {
	result, err := fnDMin(nil)
	if err != nil {
		t.Fatalf("fnDMin error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDMin(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DPRODUCT
// ---------------------------------------------------------------------------

func TestDPRODUCT(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "product Apple yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  840, // 14*10*6
		},
		{
			name: "product Pear yield",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  80, // 10*8
		},
		{
			name:     "product no matches returns 0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name: "product single match",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Yield"`,
			wantType: ValueNumber,
			wantNum:  9,
		},
		{
			name: "product ignores text",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(3)},
				{StringVal("B"), StringVal("text")},
				{StringVal("C"), NumberVal(5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  15, // 3*5
		},
	}

	runDBTests(t, "DPRODUCT", tests)
}

func TestDPRODUCT_WrongArgCount(t *testing.T) {
	result, err := fnDProduct(nil)
	if err != nil {
		t.Fatalf("fnDProduct error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDProduct(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DSTDEV
// ---------------------------------------------------------------------------

func TestDSTDEV(t *testing.T) {
	db := standardDB()

	// Apple profits: 105, 75, 45. mean=75, var=((30^2+0+30^2)/2)=900, stdev=30
	tests := []dbTestCase{
		{
			name: "stdev Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  30, // sample stdev
		},
		{
			name: "stdev Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Pear profits: 96, 76.8. mean=86.4, var=((9.6^2+9.6^2)/1)=184.32, stdev=sqrt(184.32)
			wantNum: math.Sqrt(184.32),
		},
		{
			name:     "stdev single value returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name:     "stdev no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "stdev equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(5)},
				{StringVal("B"), NumberVal(5)},
				{StringVal("C"), NumberVal(5)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
	}

	runDBTests(t, "DSTDEV", tests)
}

func TestDSTDEV_WrongArgCount(t *testing.T) {
	result, err := fnDStdev(nil)
	if err != nil {
		t.Fatalf("fnDStdev error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDStdev(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DSTDEVP
// ---------------------------------------------------------------------------

func TestDSTDEVP(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "stdevp Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Apple profits: 105,75,45. mean=75, ss=1800, var=600, stdevp=sqrt(600)
			wantNum: math.Sqrt(600),
		},
		{
			name: "stdevp single value returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name:     "stdevp no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "stdevp Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			// Pear: 96,76.8. mean=86.4, ss=184.32, var=92.16, stdevp=sqrt(92.16)
			wantNum: math.Sqrt(92.16),
		},
		{
			name: "stdevp equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(7)},
				{StringVal("B"), NumberVal(7)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
	}

	runDBTests(t, "DSTDEVP", tests)
}

func TestDSTDEVP_WrongArgCount(t *testing.T) {
	result, err := fnDStdevP(nil)
	if err != nil {
		t.Fatalf("fnDStdevP error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDStdevP(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DVAR
// ---------------------------------------------------------------------------

func TestDVAR(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "var Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  900, // sample var: (30^2+0+30^2)/2 = 900
		},
		{
			name:     "var single value returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Cherry")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name:     "var no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "var Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  184.32, // (9.6^2+9.6^2)/1
		},
		{
			name: "var equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(10)},
				{StringVal("B"), NumberVal(10)},
				{StringVal("C"), NumberVal(10)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
	}

	runDBTests(t, "DVAR", tests)
}

func TestDVAR_WrongArgCount(t *testing.T) {
	result, err := fnDVar(nil)
	if err != nil {
		t.Fatalf("fnDVar error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDVar(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// DVARP
// ---------------------------------------------------------------------------

func TestDVARP(t *testing.T) {
	db := standardDB()

	tests := []dbTestCase{
		{
			name: "varp Apple profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Apple")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  600, // population var: (30^2+0+30^2)/3 = 600
		},
		{
			name: "varp single value returns 0",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Cherry")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
		{
			name:     "varp no matches returns DIV/0",
			db:       db,
			crit:     [][]Value{{StringVal("Tree")}, {StringVal("Orange")}},
			field:    `"Profit"`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name: "varp Pear profit",
			db:   db,
			crit: [][]Value{
				{StringVal("Tree")},
				{StringVal("Pear")},
			},
			field:    `"Profit"`,
			wantType: ValueNumber,
			wantNum:  92.16, // (9.6^2+9.6^2)/2
		},
		{
			name: "varp equal values returns 0",
			db: [][]Value{
				{StringVal("Name"), StringVal("Value")},
				{StringVal("A"), NumberVal(3)},
				{StringVal("B"), NumberVal(3)},
			},
			crit:     [][]Value{{StringVal("Name")}, {StringVal("")}},
			field:    `"Value"`,
			wantType: ValueNumber,
			wantNum:  0,
		},
	}

	runDBTests(t, "DVARP", tests)
}

func TestDVARP_WrongArgCount(t *testing.T) {
	result, err := fnDVarP(nil)
	if err != nil {
		t.Fatalf("fnDVarP error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("fnDVarP(nil) = %+v, want #VALUE!", result)
	}
}

// ---------------------------------------------------------------------------
// Error propagation tests for new D-functions
// ---------------------------------------------------------------------------

func TestDFunctions_ErrorPropagation(t *testing.T) {
	db := [][]Value{
		{StringVal("Name"), StringVal("Value")},
		{StringVal("A"), NumberVal(10)},
		{StringVal("B"), ErrorVal(ErrValDIV0)},
		{StringVal("C"), NumberVal(20)},
	}
	crit := [][]Value{
		{StringVal("Name")},
		{StringVal("")},
	}

	resolver := makeDBResolver(db, 1, 1, crit, 7, 1)

	funcs := []string{"DAVERAGE", "DCOUNT", "DCOUNTA", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DVAR", "DVARP"}
	for _, fn := range funcs {
		t.Run(fn, func(t *testing.T) {
			formula := fn + `(A1:B4,"Value",G1:G2)`
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", formula, err)
			}
			if got.Type != ValueError || got.Err != ErrValDIV0 {
				t.Errorf("%s with error cell = %+v, want #DIV/0!", fn, got)
			}
		})
	}
}

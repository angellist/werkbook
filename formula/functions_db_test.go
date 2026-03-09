package formula

import (
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

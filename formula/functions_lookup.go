package formula

func init() {
	Register("HLOOKUP", NoCtx(fnHLOOKUP))
	Register("INDEX", NoCtx(fnINDEX))
	Register("MATCH", NoCtx(fnMATCH))
	Register("VLOOKUP", NoCtx(fnVLOOKUP))
}

func fnVLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	table := args[1]
	if table.Type == ValueError {
		return table, nil
	}
	if table.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}
	colIdx, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	ci := int(colIdx)
	if ci < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	rangeLookup := true
	if len(args) == 4 {
		rangeLookup = IsTruthy(args[3])
	}

	if rangeLookup {
		lastMatch := -1
		for i, row := range table.Array {
			if len(row) == 0 {
				continue
			}
			cmp := CompareValues(row[0], lookup)
			if cmp == 0 {
				lastMatch = i
				break
			}
			if cmp > 0 {
				break
			}
			lastMatch = i
		}
		if lastMatch < 0 {
			return ErrorVal(ErrValNA), nil
		}
		if ci > len(table.Array[lastMatch]) {
			return ErrorVal(ErrValREF), nil
		}
		return table.Array[lastMatch][ci-1], nil
	}

	for _, row := range table.Array {
		if len(row) == 0 {
			continue
		}
		if CompareValues(row[0], lookup) == 0 {
			if ci > len(row) {
				return ErrorVal(ErrValREF), nil
			}
			return row[ci-1], nil
		}
	}
	return ErrorVal(ErrValNA), nil
}

func fnHLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	table := args[1]
	if table.Type != ValueArray || len(table.Array) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	rowIdx, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	ri := int(rowIdx)
	if ri < 1 || ri > len(table.Array) {
		return ErrorVal(ErrValREF), nil
	}

	rangeLookup := true
	if len(args) == 4 {
		rangeLookup = IsTruthy(args[3])
	}

	firstRow := table.Array[0]

	if rangeLookup {
		lastMatch := -1
		for i, cell := range firstRow {
			cmp := CompareValues(cell, lookup)
			if cmp == 0 {
				lastMatch = i
				break
			}
			if cmp > 0 {
				break
			}
			lastMatch = i
		}
		if lastMatch < 0 {
			return ErrorVal(ErrValNA), nil
		}
		if lastMatch >= len(table.Array[ri-1]) {
			return ErrorVal(ErrValREF), nil
		}
		return table.Array[ri-1][lastMatch], nil
	}

	for i, cell := range firstRow {
		if CompareValues(cell, lookup) == 0 {
			if i >= len(table.Array[ri-1]) {
				return ErrorVal(ErrValREF), nil
			}
			return table.Array[ri-1][i], nil
		}
	}
	return ErrorVal(ErrValNA), nil
}

func fnINDEX(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	arr := args[0]
	if arr.Type != ValueArray {
		return arr, nil
	}
	rowNum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	ri := int(rowNum) - 1

	colNum := 0
	if len(args) == 3 {
		cn, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		colNum = int(cn) - 1
	}

	if ri < 0 || ri >= len(arr.Array) {
		return ErrorVal(ErrValREF), nil
	}
	if colNum < 0 || colNum >= len(arr.Array[ri]) {
		return ErrorVal(ErrValREF), nil
	}
	return arr.Array[ri][colNum], nil
}

func fnMATCH(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	arr := args[1]
	matchType := 1
	if len(args) == 3 {
		mt, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		matchType = int(mt)
	}

	var values []Value
	if arr.Type == ValueArray {
		for _, row := range arr.Array {
			values = append(values, row...)
		}
	} else {
		values = []Value{arr}
	}

	switch matchType {
	case 0:
		for i, v := range values {
			if CompareValues(v, lookup) == 0 {
				return NumberVal(float64(i + 1)), nil
			}
		}
		return ErrorVal(ErrValNA), nil

	case 1:
		last := -1
		for i, v := range values {
			cmp := CompareValues(v, lookup)
			if cmp <= 0 {
				last = i
			}
			if cmp > 0 {
				break
			}
		}
		if last < 0 {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(float64(last + 1)), nil

	case -1:
		last := -1
		for i, v := range values {
			cmp := CompareValues(v, lookup)
			if cmp >= 0 {
				last = i
			}
			if cmp < 0 {
				break
			}
		}
		if last < 0 {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(float64(last + 1)), nil
	}

	return ErrorVal(ErrValVALUE), nil
}

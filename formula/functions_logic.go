package formula

func init() {
	Register("AND", noCtx(fnAND))
	Register("IF", noCtx(fnIF))
	Register("IFERROR", noCtx(fnIFERROR))
	Register("NOT", noCtx(fnNOT))
	Register("OR", noCtx(fnOR))
}

func fnIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	// Array formula: when the condition is an array, apply IF element-wise.
	if args[0].Type == ValueArray {
		cond := args[0]
		rows := make([][]Value, len(cond.Array))
		for i, row := range cond.Array {
			out := make([]Value, len(row))
			for j, cell := range row {
				if isTruthy(cell) {
					out[j] = arrayElement(args[1], i, j)
				} else if len(args) == 3 {
					out[j] = arrayElement(args[2], i, j)
				} else {
					out[j] = BoolVal(false)
				}
			}
			rows[i] = out
		}
		return Value{Type: ValueArray, Array: rows}, nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	if isTruthy(args[0]) {
		return args[1], nil
	}
	if len(args) == 3 {
		return args[2], nil
	}
	return BoolVal(false), nil
}

func fnIFERROR(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[1], nil
	}
	return args[0], nil
}

func fnAND(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if !isTruthy(arg) {
			return BoolVal(false), nil
		}
	}
	return BoolVal(true), nil
}

func fnOR(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if isTruthy(arg) {
			return BoolVal(true), nil
		}
	}
	return BoolVal(false), nil
}

func fnNOT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return BoolVal(!isTruthy(args[0])), nil
}

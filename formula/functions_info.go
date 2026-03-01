package formula

func init() {
	Register("IFNA", NoCtx(fnIFNA))
}

func fnIFNA(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError && args[0].Err == ErrValNA {
		return args[1], nil
	}
	return args[0], nil
}

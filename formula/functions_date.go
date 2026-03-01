package formula

import (
	"math"
	"time"
)

func init() {
	Register("DATE", noCtx(fnDATE))
	Register("DAY", noCtx(fnDAY))
	Register("MONTH", noCtx(fnMONTH))
	Register("NOW", noCtx(fnNOW))
	Register("TODAY", noCtx(fnTODAY))
	Register("YEAR", noCtx(fnYEAR))
}

// Serial date helpers — duplicated from werkbook/date.go to avoid circular imports.
var excelEpoch = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)

// maxExcelSerial is the largest valid Excel serial date (Dec 31, 9999).
const maxExcelSerial = 2958465

func timeToExcelSerial(t time.Time) float64 {
	duration := t.Sub(excelEpoch)
	days := duration.Hours() / 24
	if days >= 60 {
		days++
	}
	return days
}

func excelSerialToTime(serial float64) time.Time {
	if serial > 60 {
		serial--
	}
	days := int(serial)
	frac := serial - float64(days)
	t := excelEpoch.AddDate(0, 0, days)
	t = t.Add(time.Duration(frac * 24 * float64(time.Hour)))
	return t
}

func fnDATE(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	year, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	month, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	day, e := coerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Excel checks the raw float BEFORE truncation: negative values → #NUM!
	// e.g. int(-0.5) truncates to 0 in Go, but Excel sees -0.5 < 0 → #NUM!
	if year < 0 || year >= 10000 {
		return ErrorVal(ErrValNUM), nil
	}

	y := int(year)

	// Excel adds 1900 to years in the range 0–1899.
	// e.g. DATE(108,1,2) → year 2008.
	if y >= 0 && y <= 1899 {
		y += 1900
	}

	// After normalization the year must still be in range.
	if y < 0 || y >= 10000 {
		return ErrorVal(ErrValNUM), nil
	}

	// Excel uses INT (floor) semantics to truncate month and day to integers,
	// not TRUNC (toward zero). E.g. INT(-0.5) = -1, not 0.
	m := int(math.Floor(month))
	d := int(math.Floor(day))

	// Guard against extreme month/day values that would overflow time.Duration
	// (max ≈ 292 years in nanoseconds). Excel's valid range is 1/1/1900–12/31/9999,
	// so values that shift the year far outside always produce #NUM!.
	if m < -120000 || m > 120000 || d < -4000000 || d > 4000000 {
		return ErrorVal(ErrValNUM), nil
	}

	t := time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.UTC)
	serial := timeToExcelSerial(t)
	if serial < 0 || serial > maxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(serial), nil
}

func fnDAY(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 || n > maxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Day())), nil
}

func fnMONTH(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 || n > maxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Month())), nil
}

func fnNOW(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return NumberVal(timeToExcelSerial(time.Now())), nil
}

func fnTODAY(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	now := time.Now()
	today := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.UTC)
	return NumberVal(math.Floor(timeToExcelSerial(today))), nil
}

func fnYEAR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 || n > maxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Year())), nil
}

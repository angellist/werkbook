package werkbook

import "strings"

// sheetRefNeedsQuoting reports whether a sheet name must be single-quoted in
// formula text. Matches the formula package's needsQuoting logic.
func sheetRefNeedsQuoting(name string) bool {
	for _, c := range name {
		if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_') {
			return true
		}
	}
	return false
}

// escapeSheetName doubles any apostrophes in name for use inside a
// single-quoted sheet reference (e.g. Fund's Data → Fund''s Data).
func escapeSheetName(name string) string {
	return strings.ReplaceAll(name, "'", "''")
}

// formatSheetRef returns the properly formatted sheet!-prefix for use in
// formula text. The name is quoted only when necessary, and internal
// apostrophes are doubled per Excel conventions.
func formatSheetRef(name string) string {
	if sheetRefNeedsQuoting(name) {
		return "'" + escapeSheetName(name) + "'!"
	}
	return name + "!"
}

// rewriteSheetRefsInFormula rewrites every occurrence of old sheet name
// references in a formula string to use the new name. It handles:
//   - Quoted refs: 'Old Name'!A1 → 'New Name'!A1
//   - Unquoted refs: OldName!A1 → NewName!A1
//   - Doubled apostrophes in quoted names: 'Fund''s Data'!A1
//
// Double-quoted string literals ("...") are skipped to avoid corrupting
// text content inside formulas.
func rewriteSheetRefsInFormula(src, oldName, newName string) string {
	if src == "" || oldName == newName {
		return src
	}

	// Pre-compute the escaped form of the old name for matching inside
	// single-quoted regions, and the formatted replacement prefix.
	oldEscaped := escapeSheetName(oldName)
	newPrefix := formatSheetRef(newName)

	var b strings.Builder
	b.Grow(len(src))
	i := 0

	for i < len(src) {
		ch := src[i]

		// Skip double-quoted string literals verbatim.
		if ch == '"' {
			j := i + 1
			for j < len(src) {
				if src[j] == '"' {
					j++
					if j < len(src) && src[j] == '"' {
						j++ // doubled quote escape
						continue
					}
					break
				}
				j++
			}
			b.WriteString(src[i:j])
			i = j
			continue
		}

		// Quoted sheet reference: 'Sheet Name'!
		if ch == '\'' {
			j := i + 1
			var name strings.Builder
			matched := false
			for j < len(src) {
				if src[j] == '\'' {
					if j+1 < len(src) && src[j+1] == '\'' {
						// Doubled apostrophe — part of sheet name.
						name.WriteString("''")
						j += 2
						continue
					}
					// Closing quote. Check for '!' following.
					if j+1 < len(src) && src[j+1] == '!' {
						escapedName := name.String()
						if escapedName == oldEscaped {
							b.WriteString(newPrefix)
							j += 2 // skip past '!
							matched = true
						}
					}
					break
				}
				name.WriteByte(src[j])
				j++
			}
			if matched {
				i = j
				continue
			}
			// Not a match — copy as-is and advance past closing quote.
			if j < len(src) && src[j] == '\'' {
				b.WriteString(src[i : j+1])
				i = j + 1
			} else {
				// Unterminated quote — copy remainder.
				b.WriteString(src[i:])
				return b.String()
			}
			continue
		}

		// Unquoted sheet reference: Identifier! (letters, digits, underscore, dot).
		if isUnquotedSheetStart(ch) {
			j := i + 1
			for j < len(src) && isUnquotedSheetCont(src[j]) {
				j++
			}
			word := src[i:j]
			if j < len(src) && src[j] == '!' && strings.EqualFold(word, oldName) && !sheetRefNeedsQuoting(oldName) {
				b.WriteString(newPrefix)
				i = j + 1 // skip past !
				continue
			}
			b.WriteString(word)
			i = j
			continue
		}

		b.WriteByte(ch)
		i++
	}

	return b.String()
}

func isUnquotedSheetStart(ch byte) bool {
	return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch == '_'
}

func isUnquotedSheetCont(ch byte) bool {
	return isUnquotedSheetStart(ch) || (ch >= '0' && ch <= '9') || ch == '.'
}

// rewriteSheetNameRefs updates all raw formula text and defined-name Value
// strings that reference oldName to use newName instead. Called by
// SetSheetName before rebuildFormulaState so that the recompiled formulas
// and dep graph reflect the renamed sheet.
func (f *File) rewriteSheetNameRefs(oldName, newName string) {
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c.formula == "" {
					continue
				}
				c.formula = rewriteSheetRefsInFormula(c.formula, oldName, newName)
			}
		}
	}
	for i := range f.definedNames {
		f.definedNames[i].Value = rewriteSheetRefsInFormula(f.definedNames[i].Value, oldName, newName)
	}
}

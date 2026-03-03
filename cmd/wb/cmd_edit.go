package main

import (
	"errors"
	"fmt"
	"io"
	"os"

	werkbook "github.com/jpoz/werkbook"
)

type editData struct {
	File       string     `json:"file"`
	DryRun     bool       `json:"dry_run,omitempty"`
	Applied    int        `json:"applied"`
	Operations []opResult `json:"operations"`
}

func cmdEdit(args []string, globals globalFlags) int {
	cmd := "edit"

	if hasHelpFlag(args) {
		fmt.Fprintln(os.Stderr, `Usage: wb edit [flags] <file>

Apply JSON patch operations to an existing workbook.

Flags:
  --patch <json>     Patch JSON array (or pass via stdin)
  --sheet <name>     Default sheet for operations (default: first sheet)
  --output <path>    Save to a different file (default: overwrite input)
  --dry-run          Report changes without saving

Patch format:
  [
    {"cell": "A1", "value": "hello"},
    {"cell": "B1", "value": 42},
    {"cell": "C1", "formula": "SUM(A1:B1)"},
    {"cell": "D1", "style": {"font": {"bold": true}}},
    {"cell": "A1", "clear": true},
    {"cell": "A1:C3", "clear": true},
    {"add_sheet": "NewSheet"},
    {"delete_sheet": "OldSheet"},
    {"cell": "A", "column_width": 25.0},
    {"row": 1, "row_height": 30.0}
  ]

Examples:
  wb edit --patch '[{"cell":"A1","value":"updated"}]' data.xlsx
  echo '[{"cell":"B1","formula":"SUM(A1:A10)"}]' | wb edit data.xlsx
  wb edit --dry-run --patch '[{"cell":"A1","clear":true}]' data.xlsx

Note: Setting cell values does not auto-expand formula ranges. If you add
data beyond a range like SUM(B2:B3), update the formula separately.`)
		return ExitSuccess
	}

	var sheetFlag, patchFlag, outputFlag string
	var dryRun bool

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
		case "--patch":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--patch requires a value"), globals)
				return ExitUsage
			}
			patchFlag = args[i+1]
			i += 2
		case "--output":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--output requires a value"), globals)
				return ExitUsage
			}
			outputFlag = args[i+1]
			i += 2
		case "--dry-run":
			dryRun = true
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

	// Read patch from flag or stdin.
	var patchBytes []byte
	if patchFlag != "" {
		patchBytes = []byte(patchFlag)
	} else {
		var err error
		patchBytes, err = io.ReadAll(os.Stdin)
		if err != nil {
			writeError(cmd, errInternal(err), globals)
			return ExitInternal
		}
	}

	if len(patchBytes) == 0 {
		writeError(cmd, errInvalidPatch(nil), globals)
		return ExitValidate
	}

	ops, err := parsePatchOps(patchBytes)
	if err != nil {
		writeError(cmd, errInvalidPatch(err), globals)
		return ExitValidate
	}

	f, err := werkbook.Open(filePath)
	if err != nil {
		if os.IsNotExist(err) {
			writeError(cmd, errFileNotFound(filePath, err), globals)
		} else if errors.Is(err, werkbook.ErrEncryptedFile) {
			writeError(cmd, errEncryptedFile(filePath), globals)
		} else {
			writeError(cmd, errFileOpen(filePath, err), globals)
		}
		return ExitFileIO
	}

	// Determine default sheet.
	defaultSheet := sheetFlag
	if defaultSheet == "" {
		names := f.SheetNames()
		if len(names) > 0 {
			defaultSheet = names[0]
		}
	}

	results, applied := applyPatches(f, ops, defaultSheet)

	if !dryRun {
		savePath := outputFlag
		if savePath == "" {
			savePath = filePath
		}
		if err := f.SaveAs(savePath); err != nil {
			writeError(cmd, errFileSave(savePath, err), globals)
			return ExitFileIO
		}
	}

	data := editData{
		File:       filePath,
		DryRun:     dryRun,
		Applied:    applied,
		Operations: results,
	}

	if applied < len(ops) {
		resp := &Response{
			OK:      false,
			Command: cmd,
			Data:    data,
			Error: &ErrorInfo{
				Code:    ErrCodePartialFailure,
				Message: fmt.Sprintf("%d of %d operations failed", len(ops)-applied, len(ops)),
				Hint:    "Check the 'operations' array for per-operation errors.",
			},
		}
		out, err := marshalJSON(resp)
		if err != nil {
			fmt.Fprintf(os.Stderr, `{"ok":false,"error":{"code":"INTERNAL","message":%q}}`+"\n", err.Error())
		} else {
			fmt.Fprintln(os.Stderr, string(out))
		}
		return ExitPartial
	}

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

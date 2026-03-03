package main

import (
	"fmt"
	"os"
	"sort"

	"github.com/jpoz/werkbook/formula"
)

type formulaListData struct {
	Functions []string `json:"functions"`
	Count     int      `json:"count"`
}

func cmdFormula(args []string, globals globalFlags) int {
	cmd := "formula"

	if hasHelpFlag(args) {
		fmt.Fprintln(os.Stderr, `Usage: wb formula <subcommand>

Formula-related subcommands.

Subcommands:
  list    List all registered formula functions

Examples:
  wb formula list`)
		return ExitSuccess
	}

	if len(args) == 0 {
		writeError(cmd, errUsage("subcommand required: 'formula list'"), globals)
		return ExitUsage
	}

	switch args[0] {
	case "list":
		return cmdFormulaList(globals)
	default:
		writeError(cmd, errUsage("unknown subcommand: "+args[0]+". Available: list"), globals)
		return ExitUsage
	}
}

func cmdFormulaList(globals globalFlags) int {
	funcs := formula.RegisteredFunctions()
	sort.Strings(funcs)

	data := formulaListData{
		Functions: funcs,
		Count:     len(funcs),
	}
	writeSuccess("formula", data, globals)
	return ExitSuccess
}

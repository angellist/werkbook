package main

var version = "dev"

type versionData struct {
	Version string `json:"version"`
}

func cmdVersion(globals globalFlags) int {
	data := versionData{Version: version}
	writeSuccess("version", data, globals)
	return ExitSuccess
}

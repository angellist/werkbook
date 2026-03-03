package main

import "encoding/json"

// Response is the JSON envelope wrapping all command output.
type Response struct {
	OK      bool        `json:"ok"`
	Command string      `json:"command"`
	Data    any         `json:"data"`
	Error   *ErrorInfo  `json:"error,omitempty"`
}

// ErrorInfo describes a structured error with remediation hints.
type ErrorInfo struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Hint    string `json:"hint,omitempty"`
}

func successResponse(command string, data any) *Response {
	return &Response{OK: true, Command: command, Data: data}
}

func errorResponse(command string, ei *ErrorInfo) *Response {
	return &Response{OK: false, Command: command, Data: nil, Error: ei}
}

func marshalJSON(v any) ([]byte, error) {
	return json.MarshalIndent(v, "", "  ")
}

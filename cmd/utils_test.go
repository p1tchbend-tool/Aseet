package cmd

import "testing"

func TestIsExcelFile(t *testing.T) {
	tests := []struct {
		ext      string
		expected bool
	}{
		{".xlsx", true},
		{".xlsm", true},
		{".xlam", true},
		{".xltm", true},
		{".xltx", true},
		{".xls", false},
		{".txt", false},
		{"", false},
	}

	for _, tt := range tests {
		t.Run(tt.ext, func(t *testing.T) {
			if got := isExcelFile(tt.ext); got != tt.expected {
				t.Errorf("isExcelFile(%q) = %v, want %v", tt.ext, got, tt.expected)
			}
		})
	}
}

func TestEscapeCSVField(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{"Normal", "hello", "hello"},
		{"With comma", "hello,world", "\"hello,world\""},
		{"With quotes", "hello\"world", "\"hello\"\"world\""},
		{"With newline", "hello\nworld", "\"hello\\nworld\""},
		{"Mixed", "a,b\n\"c\"", "\"a,b\\n\"\"c\"\"\""},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := escapeCSVField(tt.input); got != tt.expected {
				t.Errorf("escapeCSVField(%q) = %q, want %q", tt.input, got, tt.expected)
			}
		})
	}
}

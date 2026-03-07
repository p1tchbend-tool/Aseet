package cmd

import (
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

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

func TestGetSheetData_TestData(t *testing.T) {
	f, err := excelize.OpenFile("testdata/testbook1.xlsx")
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	sheetName := "Sheet1"

	t.Run("Without Formula", func(t *testing.T) {
		got, err := getSheetData(f, sheetName, false)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}

		expected := [][]string{
			{"Hello", "World"},
		}

		if !reflect.DeepEqual(got, expected) {
			t.Errorf("getSheetData() = %v, want %v", got, expected)
		}
	})
}

func TestGetSheetContents_TestData(t *testing.T) {
	f, err := excelize.OpenFile("testdata/testbook1.xlsx")
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	sheetName := "Sheet1"

	t.Run("Without Formula", func(t *testing.T) {
		got, err := getSheetContents(f, sheetName, false)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}

		expected := "Hello,World\n"

		if got != expected {
			t.Errorf("getSheetContents() = %q, want %q", got, expected)
		}
	})
}

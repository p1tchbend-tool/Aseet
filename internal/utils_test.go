package internal

import (
	"os"
	"path/filepath"
	"reflect"
	"sort"
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
			if got := IsExcelFile(tt.ext); got != tt.expected {
				t.Errorf("IsExcelFile(%q) = %v, want %v", tt.ext, got, tt.expected)
			}
		})
	}
}

func TestGetSheetData_TestData(t *testing.T) {
	// 一時ディレクトリを作成
	tempDir := t.TempDir()
	filePath := filepath.Join(tempDir, "testbook.xlsx")

	// 動的にExcelファイルを作成
	fNew := excelize.NewFile()
	defer fNew.Close()

	sheetName := "Sheet1"
	// デフォルトのシート名が "Sheet1" でない場合を考慮して作成/取得
	fNew.NewSheet(sheetName)

	// テスト用データの書き込み (A1: "Hello", B1: "World")
	_ = fNew.SetCellValue(sheetName, "A1", "Hello")
	_ = fNew.SetCellValue(sheetName, "B1", "World")

	// ファイルに保存
	if err := fNew.SaveAs(filePath); err != nil {
		t.Fatalf("failed to create temporary excel file: %v", err)
	}

	// 作成したファイルを読み込んでテストを実行
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	t.Run("Without Formula", func(t *testing.T) {
		got, err := GetSheetData(f, sheetName, false)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}

		expected := [][]string{
			{"Hello", "World"},
		}

		if !reflect.DeepEqual(got, expected) {
			t.Errorf("GetSheetData() = %v, want %v", got, expected)
		}
	})
}

func TestCountNonEmpty(t *testing.T) {
	tests := []struct {
		name     string
		row      []string
		expected int
	}{
		{"All empty", []string{"", "", ""}, 0},
		{"Mixed", []string{"a", "", "b"}, 2},
		{"All filled", []string{"a", "b", "c"}, 3},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := CountNonEmpty(tt.row); got != tt.expected {
				t.Errorf("CountNonEmpty() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestTranspose(t *testing.T) {
	input := [][]string{
		{"1", "2", "3"},
		{"4", "5"},
	}
	expected := [][]string{
		{"1", "4"},
		{"2", "5"},
		{"3", ""},
	}

	got := Transpose(input)
	if !reflect.DeepEqual(got, expected) {
		t.Errorf("Transpose() = %v, want %v", got, expected)
	}
}

func TestCalcMatchCost(t *testing.T) {
	tests := []struct {
		name     string
		row1     []string
		row2     []string
		expected int
	}{
		{"Identical", []string{"a", "b"}, []string{"a", "b"}, 0},
		{"Different", []string{"a", "b"}, []string{"a", "c"}, 1},
		{"Different length", []string{"a"}, []string{"a", "b"}, 1},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := CalcMatchCost(tt.row1, tt.row2); got != tt.expected {
				t.Errorf("CalcMatchCost() = %v, want %v", got, tt.expected)
			}
		})
	}
}

func TestAlign(t *testing.T) {
	a := [][]string{{"A"}, {"B"}, {"C"}}
	b := [][]string{{"A"}, {"C"}}

	// Bが削除されたケースのアライメントパスを期待
	// [0,0] -> AとAがマッチ
	// [1,-1] -> Bが削除
	// [2,1] -> CとCがマッチ
	expected := [][2]int{
		{0, 0},
		{1, -1},
		{2, 1},
	}

	got := Align(a, b)
	if !reflect.DeepEqual(got, expected) {
		t.Errorf("Align() = %v, want %v", got, expected)
	}
}

func TestFindExcelFiles(t *testing.T) {
	// テスト用の一時ディレクトリを作成（テスト終了時に自動削除）
	tempDir := t.TempDir()

	// テスト用のディレクトリ構造とファイルを作成
	// tempDir/
	// ├── a.xlsx      (対象)
	// ├── b.txt       (対象外)
	// └── sub/
	//     ├── c.xlsm  (対象)
	//     └── d.csv   (対象外)

	subDir := filepath.Join(tempDir, "sub")
	if err := os.Mkdir(subDir, 0755); err != nil {
		t.Fatalf("failed to create sub directory: %v", err)
	}

	filesToCreate := []string{
		filepath.Join(tempDir, "a.xlsx"),
		filepath.Join(tempDir, "b.txt"),
		filepath.Join(subDir, "c.xlsm"),
		filepath.Join(subDir, "d.csv"),
	}

	for _, f := range filesToCreate {
		if err := os.WriteFile(f, []byte("dummy"), 0644); err != nil {
			t.Fatalf("failed to create dummy file %s: %v", f, err)
		}
	}

	t.Run("Non-recursive", func(t *testing.T) {
		got := FindExcelFiles(tempDir, false)
		expected := []string{
			filepath.Join(tempDir, "a.xlsx"),
		}

		// 順序に依存しないようにソートして比較
		sort.Strings(got)
		sort.Strings(expected)

		if !reflect.DeepEqual(got, expected) {
			t.Errorf("FindExcelFiles(recursive=false) = %v, want %v", got, expected)
		}
	})

	t.Run("Recursive", func(t *testing.T) {
		got := FindExcelFiles(tempDir, true)
		expected := []string{
			filepath.Join(tempDir, "a.xlsx"),
			filepath.Join(subDir, "c.xlsm"),
		}

		// 順序に依存しないようにソートして比較
		sort.Strings(got)
		sort.Strings(expected)

		if !reflect.DeepEqual(got, expected) {
			t.Errorf("FindExcelFiles(recursive=true) = %v, want %v", got, expected)
		}
	})
}

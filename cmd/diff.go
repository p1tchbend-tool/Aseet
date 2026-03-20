package cmd

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/pmezard/go-difflib/difflib"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// 差分表示用のカラーコード
const (
	colorDel    = "[#d55e00]"
	colorAdd    = "[#56b4e9]"
	colorChange = "[#f0e442]"
	colorReset  = "[-]"
)

var diffFormula bool
var diffOpen bool
var diffSheetName string

// 変更されたシートとそのセル座標を保持する構造体
type modifiedSheet struct {
	name  string
	cells []string
}

var diffCmd = &cobra.Command{
	Use:   "diff [file1/dir1] [file2/dir2]",
	Short: "Compare sheet names and cell contents of two Excel files or directories",
	Long:  `Compare sheet names of two Excel files and output the differences in unified diff format. For sheets with the same name, compare the cell contents cell by cell. If directories are provided, compares Excel files with matching relative paths.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		path1 := args[0]
		path2 := args[1]

		info1, err1 := os.Stat(path1)
		info2, err2 := os.Stat(path2)

		if err1 != nil || err2 != nil {
			fmt.Println("Error accessing paths.")
			os.Exit(1)
		}

		if info1.IsDir() && info2.IsDir() {
			// ディレクトリ同士の比較
			files1 := findExcelFiles(path1, true)
			files2 := findExcelFiles(path2, true)

			fileMap1 := make(map[string]string)
			fileMap2 := make(map[string]string)
			var allRelPaths []string
			relPathMap := make(map[string]bool)

			for _, f := range files1 {
				rel, _ := filepath.Rel(path1, f)
				fileMap1[rel] = f
				if !relPathMap[rel] {
					allRelPaths = append(allRelPaths, rel)
					relPathMap[rel] = true
				}
			}
			for _, f := range files2 {
				rel, _ := filepath.Rel(path2, f)
				fileMap2[rel] = f
				if !relPathMap[rel] {
					allRelPaths = append(allRelPaths, rel)
					relPathMap[rel] = true
				}
			}

			sort.Strings(allRelPaths)

			var books []bookResult
			for _, rel := range allRelPaths {
				f1, ok1 := fileMap1[rel]
				f2, ok2 := fileMap2[rel]

				var results []sheetResult
				if ok1 && ok2 {
					results = compareExcelFiles(f1, f2)
				} else if ok1 {
					results = []sheetResult{{title: "Info", content: fmt.Sprintf("%s only exists in %s", rel, path1)}}
				} else if ok2 {
					results = []sheetResult{{title: "Info", content: fmt.Sprintf("%s only exists in %s", rel, path2)}}
				}

				if len(results) > 0 {
					books = append(books, bookResult{
						fileName: rel,
						sheets:   results,
					})
				}
			}

			if len(books) == 0 {
				fmt.Println("No differences found.")
				return
			}

			if err := displayDirTui(books); err != nil {
				fmt.Printf("Error running TUI: %v\n", err)
				os.Exit(1)
			}

		} else if !info1.IsDir() && !info2.IsDir() {
			// ファイル同士の比較
			if diffOpen {
				handleDiffOpen(path1, path2)
			}

			results := compareExcelFiles(path1, path2)
			if len(results) == 0 {
				fmt.Println("No differences found.")
				return
			}

			if err := displayFileTui(results); err != nil {
				fmt.Printf("Error running TUI: %v\n", err)
				os.Exit(1)
			}

		} else {
			// ディレクトリとファイルの比較
			fmt.Println("Cannot compare a file with a directory.")
			os.Exit(1)
		}
	},
}

func handleDiffOpen(file1, file2 string) {
	cacheDir, err := os.UserCacheDir()
	if err != nil {
		fmt.Printf("Error getting cache directory: %v\n", err)
		os.Exit(1)
	}
	aseetDir := filepath.Join(cacheDir, "aseet")
	if err := os.MkdirAll(aseetDir, 0755); err != nil {
		fmt.Printf("Error creating cache directory: %v\n", err)
		os.Exit(1)
	}

	localPath := filepath.Join(aseetDir, "[OLD]"+filepath.Base(file1))
	remotePath := filepath.Join(aseetDir, "[NEW]"+filepath.Base(file2))

	if err := copyFile(file1, localPath); err != nil {
		fmt.Printf("Error copying file1: %v\n", err)
		os.Exit(1)
	}
	if err := copyFile(file2, remotePath); err != nil {
		fmt.Printf("Error copying file2: %v\n", err)
		os.Exit(1)
	}

	openFile(localPath)
	openFile(remotePath)
}

func compareExcelFiles(file1, file2 string) []sheetResult {
	f1, err := excelize.OpenFile(file1)
	if err != nil {
		return []sheetResult{{title: "Error", content: fmt.Sprintf("Error opening file %s", file1)}}
	}
	defer f1.Close()

	f2, err := excelize.OpenFile(file2)
	if err != nil {
		return []sheetResult{{title: "Error", content: fmt.Sprintf("Error opening file %s", file2)}}
	}
	defer f2.Close()

	var results []sheetResult

	sheets1 := f1.GetSheetList()
	sheets2 := f2.GetSheetList()

	sort.Strings(sheets1)
	sort.Strings(sheets2)

	var sheetListDiff string

	if diffSheetName == "" {
		text1 := strings.Join(sheets1, "\n") + "\n"
		text2 := strings.Join(sheets2, "\n") + "\n"

		diff := difflib.UnifiedDiff{
			A:        difflib.SplitLines(text1),
			B:        difflib.SplitLines(text2),
			FromFile: file1,
			ToFile:   file2,
			Context:  3,
		}

		text, err := difflib.GetUnifiedDiffString(diff)
		if err == nil && text != "" {
			var diffLines []string
			lines := strings.Split(text, "\n")
			for _, line := range lines {
				if line == "" || strings.HasPrefix(line, "---") || strings.HasPrefix(line, "+++") || strings.HasPrefix(line, "@@") {
					continue
				}
				if strings.HasPrefix(line, "-") {
					diffLines = append(diffLines, colorDel+line+colorReset)
				} else if strings.HasPrefix(line, "+") {
					diffLines = append(diffLines, colorAdd+line+colorReset)
				} else {
					diffLines = append(diffLines, line)
				}
			}
			sheetListDiff = strings.Join(diffLines, "\n")
		}
	}

	sheetMap1 := make(map[string]bool)
	sheetMap2 := make(map[string]bool)
	var allSheets []string

	for _, s := range sheets1 {
		sheetMap1[s] = true
	}
	for _, s := range sheets2 {
		sheetMap2[s] = true
	}

	if diffSheetName != "" {
		if !sheetMap1[diffSheetName] && !sheetMap2[diffSheetName] {
			return []sheetResult{{title: "Error", content: fmt.Sprintf("Sheet %s does not exist in either file.", diffSheetName)}}
		}
		allSheets = []string{diffSheetName}
	} else {
		for _, s := range sheets1 {
			allSheets = append(allSheets, s)
		}
		for _, s := range sheets2 {
			if !sheetMap1[s] {
				allSheets = append(allSheets, s)
			}
		}
	}

	var modifiedSheets []modifiedSheet

	for _, sheet := range allSheets {
		in1 := sheetMap1[sheet]
		in2 := sheetMap2[sheet]

		if in1 && in2 {
			rows1, err := getSheetData(f1, sheet, diffFormula)
			if err != nil {
				continue
			}
			rows2, err := getSheetData(f2, sheet, diffFormula)
			if err != nil {
				continue
			}

			rowAlign := align(rows1, rows2)
			colAlign := align(transpose(rows1), transpose(rows2))

			hasSheetDiff := false
			var sheetOutput []string
			var changedCells []string

			for _, rPair := range rowAlign {
				r1, r2 := rPair[0], rPair[1]
				var diffCells []string

				for _, cPair := range colAlign {
					c1, c2 := cPair[0], cPair[1]
					val1, val2 := "", ""

					if r1 != -1 && c1 != -1 && r1 < len(rows1) && c1 < len(rows1[r1]) {
						val1 = rows1[r1][c1]
					}
					if r2 != -1 && c2 != -1 && r2 < len(rows2) && c2 < len(rows2[r2]) {
						val2 = rows2[r2][c2]
					}

					if val1 == val2 {
						diffCells = append(diffCells, escapeCSVField(val1))
					} else {
						hasSheetDiff = true

						rIdx := r2
						if rIdx == -1 {
							rIdx = r1
						}
						cIdx := c2
						if cIdx == -1 {
							cIdx = c1
						}
						if rIdx != -1 && cIdx != -1 {
							cellName, _ := excelize.CoordinatesToCellName(cIdx+1, rIdx+1)
							if cellName != "" {
								changedCells = append(changedCells, cellName)
							}
						}

						var cellDiff string
						if val1 != "" && val2 != "" {
							cellDiff = fmt.Sprintf("%s-%s%s %s+%s%s", colorDel, escapeCSVField(val1), colorReset, colorAdd, escapeCSVField(val2), colorReset)
						} else if val1 != "" {
							cellDiff = fmt.Sprintf("%s-%s%s", colorDel, escapeCSVField(val1), colorReset)
						} else if val2 != "" {
							cellDiff = fmt.Sprintf("%s+%s%s", colorAdd, escapeCSVField(val2), colorReset)
						}
						diffCells = append(diffCells, cellDiff)
					}
				}
				sheetOutput = append(sheetOutput, strings.Join(diffCells, ","))
			}

			if hasSheetDiff {
				modifiedSheets = append(modifiedSheets, modifiedSheet{
					name:  sheet,
					cells: changedCells,
				})
				results = append(results, sheetResult{
					title:   colorChange + sheet + colorReset,
					content: strings.Join(sheetOutput, "\n"),
				})
			}
		} else if in1 {
			content, _ := getSheetContents(f1, sheet, diffFormula)
			results = append(results, sheetResult{
				title:   fmt.Sprintf("%s%s : %s%s", colorDel, filepath.Base(file1), sheet, colorReset),
				content: content,
			})
		} else if in2 {
			content, _ := getSheetContents(f2, sheet, diffFormula)
			results = append(results, sheetResult{
				title:   fmt.Sprintf("%s%s : %s%s", colorAdd, filepath.Base(file2), sheet, colorReset),
				content: content,
			})
		}
	}

	var summaryBuilder strings.Builder
	if sheetListDiff != "" {
		summaryBuilder.WriteString("\n[Sheet Name Differences]\n\n")
		summaryBuilder.WriteString(sheetListDiff)
		summaryBuilder.WriteString("\n")
	}
	if len(modifiedSheets) > 0 {
		summaryBuilder.WriteString("\n[Modified Sheets (Cell Differences)]\n\n")
		sort.Slice(modifiedSheets, func(i, j int) bool {
			return modifiedSheets[i].name < modifiedSheets[j].name
		})
		for _, ms := range modifiedSheets {
			summaryBuilder.WriteString(fmt.Sprintf("%s%s: %s%s\n", colorChange, ms.name, strings.Join(ms.cells, ", "), colorReset))
		}
	}

	summaryText := summaryBuilder.String()
	if summaryText != "" {
		results = append([]sheetResult{{
			title:   "Summary",
			content: summaryText,
		}}, results...)
	}

	return results
}

func init() {
	rootCmd.AddCommand(diffCmd)
	diffCmd.Flags().BoolVarP(&diffFormula, "formula", "f", false, "If the cell value is a formula, compare the formula instead of the value.")
	diffCmd.Flags().StringVarP(&diffSheetName, "name", "n", "", "Compare only the specified sheet.")
	diffCmd.Flags().BoolVarP(&diffOpen, "open", "o", false, "Copy the two files to the cache directory with [LOCAL] and [REMOTE] prefixes and open them.")
}

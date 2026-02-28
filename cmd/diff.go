package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/pmezard/go-difflib/difflib"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

const (
	colorLightOrange = "\033[38;5;215m"
	colorLightBlue   = "\033[38;5;117m"
	colorCyan        = "\033[36m"
	colorReset       = "\033[0m"
)

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Compare sheet names and cell contents of two Excel files",
	Long:  `Compare sheet names of two Excel files and output the differences in unified diff format. For sheets with the same name, compare the cell contents cell by cell.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		file1 := args[0]
		file2 := args[1]

		f1, err := excelize.OpenFile(file1)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", file1, err)
			os.Exit(1)
		}
		defer f1.Close()

		f2, err := excelize.OpenFile(file2)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error opening file %s: %v\n", file2, err)
			os.Exit(1)
		}
		defer f2.Close()

		sheets1 := f1.GetSheetList()
		sheets2 := f2.GetSheetList()

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
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error generating diff: %v\n", err)
			os.Exit(1)
		}

		if text != "" {
			// Colorize unified diff output
			lines := strings.Split(text, "\n")
			for i, line := range lines {
				if strings.HasPrefix(line, "---") || strings.HasPrefix(line, "+++") {
					// Keep original for file headers
				} else if strings.HasPrefix(line, "-") {
					lines[i] = colorLightOrange + line + colorReset
				} else if strings.HasPrefix(line, "+") {
					lines[i] = colorLightBlue + line + colorReset
				} else if strings.HasPrefix(line, "@@") {
					lines[i] = colorCyan + line + colorReset
				}
			}
			fmt.Print(strings.Join(lines, "\n"))
		}

		// Compare cell contents for common sheets
		sheetMap1 := make(map[string]bool)
		for _, s := range sheets1 {
			sheetMap1[s] = true
		}

		for _, sheet := range sheets2 {
			if sheetMap1[sheet] {
				rows1, err := f1.GetRows(sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file1, err)
					continue
				}

				rows2, err := f2.GetRows(sheet)
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error reading sheet %s from %s: %v\n", sheet, file2, err)
					continue
				}

				rowAlign := align(rows1, rows2)
				colAlign := align(transpose(rows1), transpose(rows2))

				hasSheetDiff := false
				var sheetOutput []string

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
							diffCells = append(diffCells, val1)
						} else {
							hasSheetDiff = true
							var cellDiff string
							if val1 != "" && val2 != "" {
								cellDiff = fmt.Sprintf("%s-%s%s %s+%s%s", colorLightOrange, val1, colorReset, colorLightBlue, val2, colorReset)
							} else if val1 != "" {
								cellDiff = fmt.Sprintf("%s-%s%s", colorLightOrange, val1, colorReset)
							} else if val2 != "" {
								cellDiff = fmt.Sprintf("%s+%s%s", colorLightBlue, val2, colorReset)
							}
							diffCells = append(diffCells, cellDiff)
						}
					}
					sheetOutput = append(sheetOutput, strings.Join(diffCells, ","))
				}

				if hasSheetDiff {
					fmt.Printf("\n[diff %s]\n", sheet)
					for _, line := range sheetOutput {
						fmt.Println(line)
					}
				}
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(diffCmd)
}

func countNonEmpty(row []string) int {
	c := 0
	for _, v := range row {
		if v != "" {
			c++
		}
	}
	return c
}

func transpose(matrix [][]string) [][]string {
	maxCol := 0
	for _, row := range matrix {
		if len(row) > maxCol {
			maxCol = len(row)
		}
	}
	res := make([][]string, maxCol)
	for i := 0; i < maxCol; i++ {
		res[i] = make([]string, len(matrix))
		for j, row := range matrix {
			if i < len(row) {
				res[i][j] = row[i]
			}
		}
	}
	return res
}

func align(a, b [][]string) [][2]int {
	n, m := len(a), len(b)
	dp := make([][]int, n+1)
	for i := range dp {
		dp[i] = make([]int, m+1)
	}
	for i := 1; i <= n; i++ {
		dp[i][0] = dp[i-1][0] + countNonEmpty(a[i-1])
	}
	for j := 1; j <= m; j++ {
		dp[0][j] = dp[0][j-1] + countNonEmpty(b[j-1])
	}

	for i := 1; i <= n; i++ {
		for j := 1; j <= m; j++ {
			costDel := dp[i-1][j] + countNonEmpty(a[i-1])
			costIns := dp[i][j-1] + countNonEmpty(b[j-1])

			matchCost := 0
			maxL := len(a[i-1])
			if len(b[j-1]) > maxL {
				maxL = len(b[j-1])
			}
			for k := 0; k < maxL; k++ {
				v1, v2 := "", ""
				if k < len(a[i-1]) {
					v1 = a[i-1][k]
				}
				if k < len(b[j-1]) {
					v2 = b[j-1][k]
				}
				if v1 != v2 {
					matchCost++
				}
			}
			costMatch := dp[i-1][j-1] + matchCost

			minCost := costDel
			if costIns < minCost {
				minCost = costIns
			}
			if costMatch < minCost {
				minCost = costMatch
			}
			dp[i][j] = minCost
		}
	}

	var path [][2]int
	i, j := n, m
	for i > 0 || j > 0 {
		if i > 0 && j > 0 {
			matchCost := 0
			maxL := len(a[i-1])
			if len(b[j-1]) > maxL {
				maxL = len(b[j-1])
			}
			for k := 0; k < maxL; k++ {
				v1, v2 := "", ""
				if k < len(a[i-1]) {
					v1 = a[i-1][k]
				}
				if k < len(b[j-1]) {
					v2 = b[j-1][k]
				}
				if v1 != v2 {
					matchCost++
				}
			}
			if dp[i][j] == dp[i-1][j-1]+matchCost {
				path = append([][2]int{{i - 1, j - 1}}, path...)
				i--
				j--
				continue
			}
		}
		if i > 0 && dp[i][j] == dp[i-1][j]+countNonEmpty(a[i-1]) {
			path = append([][2]int{{i - 1, -1}}, path...)
			i--
		} else {
			path = append([][2]int{{-1, j - 1}}, path...)
			j--
		}
	}
	return path
}

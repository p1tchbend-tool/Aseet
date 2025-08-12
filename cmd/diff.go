package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// getSheetNames opens an excel file and returns a list of its sheet names.
func getSheetNames(filePath string) ([]string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Fprintf(os.Stderr, "Error closing file %s: %v\n", filePath, err)
		}
	}()
	return f.GetSheetList(), nil
}

var diffCmd = &cobra.Command{
	Use:   "diff [file1] [file2]",
	Short: "Show the difference in sheet names between two excel files",
	Long:  `Show the difference in sheet names between two excel files.`,
	Args:  cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		file1Path := args[0]
		file2Path := args[1]

		sheets1, err := getSheetNames(file1Path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error processing file %s: %v\n", file1Path, err)
			os.Exit(1)
		}

		sheets2, err := getSheetNames(file2Path)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error processing file %s: %v\n", file2Path, err)
			os.Exit(1)
		}

		map1 := make(map[string]bool)
		for _, s := range sheets1 {
			map1[s] = true
		}

		map2 := make(map[string]bool)
		for _, s := range sheets2 {
			map2[s] = true
		}

		onlyIn1 := []string{}
		for _, s := range sheets1 {
			if !map2[s] {
				onlyIn1 = append(onlyIn1, s)
			}
		}

		onlyIn2 := []string{}
		for _, s := range sheets2 {
			if !map1[s] {
				onlyIn2 = append(onlyIn2, s)
			}
		}

		hasDiff := false
		if len(onlyIn1) > 0 {
			hasDiff = true
			fmt.Printf("Sheets only in %s:\n", file1Path)
			for _, s := range onlyIn1 {
				fmt.Printf("- %s\n", s)
			}
			fmt.Println()
		}

		if len(onlyIn2) > 0 {
			hasDiff = true
			fmt.Printf("Sheets only in %s:\n", file2Path)
			for _, s := range onlyIn2 {
				fmt.Printf("- %s\n", s)
			}
			fmt.Println()
		}

		if !hasDiff {
			fmt.Println("The sheet names are identical in both files.")
		}
	},
}

func init() {
	rootCmd.AddCommand(diffCmd)
}

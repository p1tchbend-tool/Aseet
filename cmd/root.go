package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "aseet",
	Short: "A CLI tool for various operations on Excel files",
	Long:  `aseet is a CLI tool that allows you to perform various operations on Excel files (OOXML format). It supports commands like cat, clear, diff, grep, and sd to view, compare, search, and replace contents within Excel files.`,
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

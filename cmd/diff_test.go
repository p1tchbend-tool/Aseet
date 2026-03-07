package cmd

import (
	"reflect"
	"testing"
)

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
			if got := countNonEmpty(tt.row); got != tt.expected {
				t.Errorf("countNonEmpty() = %v, want %v", got, tt.expected)
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

	got := transpose(input)
	if !reflect.DeepEqual(got, expected) {
		t.Errorf("transpose() = %v, want %v", got, expected)
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
			if got := calcMatchCost(tt.row1, tt.row2); got != tt.expected {
				t.Errorf("calcMatchCost() = %v, want %v", got, tt.expected)
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

	got := align(a, b)
	if !reflect.DeepEqual(got, expected) {
		t.Errorf("align() = %v, want %v", got, expected)
	}
}

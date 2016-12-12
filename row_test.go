package excl

import "testing"

func BenchmarkNewRow(b *testing.B) {
	row := &Row{rowID: 10}
	row.CreateCells(1, b.N)
}

package node

import "github.com/tealeg/xlsx"

type col struct {
	title string
	width float64
}

func writeSitePrefix(r *xlsx.Row, nbCol int) {
	for i := 0; i < nbCol; i++ {
		r.AddCell()
	}
}

func addStyleOnRow(r *xlsx.Row, st *xlsx.Style, nbCols int) {
	for i := 0; i < nbCols; i++ {
		r.Cells[i].SetStyle(st)
	}
}

func addHeaderRow(xs *xlsx.Sheet, cols []col) {
	r := xs.AddRow()
	for i, ci := range cols {
		r.AddCell().SetString(ci.title)
		xs.Col(i).Width = ci.width
	}
}

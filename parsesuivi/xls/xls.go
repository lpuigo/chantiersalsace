package xls

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

func RcToAxis(row, col int) string {
	res, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		res = "A1"
	}
	return res
	colName, _ := excelize.ColumnNumberToName(col)
	return colName + strconv.Itoa(row+1)
}

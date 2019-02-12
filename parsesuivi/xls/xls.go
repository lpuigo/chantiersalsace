package xls

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

func RcToAxis(row, col int) string {
	return excelize.ToAlphaString(col) + strconv.Itoa(row+1)
}

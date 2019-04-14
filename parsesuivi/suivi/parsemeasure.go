package suivi

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"strings"
	"time"
)

type MeasurementParser struct {
	row      int
	activity string
}

func NewMeasurementParser() *MeasurementParser {
	return &MeasurementParser{
		row:      0,
		activity: sheetMeasure,
	}
}

func (rm *MeasurementParser) Activity() string {
	return rm.activity
}

const (
	colMeasName     int = 0
	colMeasNbFiber  int = 1
	colMeasNbSplice int = 2
	colMeasStatus   int = 6
	colMeasDate     int = 10
)

func (rm *MeasurementParser) ParseBlock(sh *xlsx.Sheet, catalog *bpu.Catalog, row int) (items []*bpu.Item, err ParsingError, nextRow int) {
	var e error

	boxName := sh.Cell(row, colMeasName).Value
	boxNbFiber := sh.Cell(row, colMeasNbFiber).Value
	boxNbSplice := sh.Cell(row, colMeasNbSplice).Value

	// Check for end of data
	if boxNbSplice == "" && boxName == "" {
		return
	}

	// check for box declaration
	if !(boxNbSplice != "" && boxNbFiber != "" && boxName != "") {
		err.Add(fmt.Errorf("invalid Measurement definition in line %s:%d", rm.activity, row+1))
		return
	}

	// parse measurement declaration line
	items, e = rm.newItemFromXLSRow(sh, row, catalog)
	if e != nil {
		err.Add(e)
	}

	// parse remaining block detail lines
	nextRow = row + 1
	for {
		boxName := sh.Cell(nextRow, colMeasName).Value
		boxNbFiber := sh.Cell(nextRow, colMeasNbFiber).Value
		boxNbSplice := sh.Cell(nextRow, colMeasNbSplice).Value
		// check for end of data
		if !(boxName == "" && boxNbFiber == "" && boxNbSplice != "") {
			break
		}
		nextRow++
	}
	return
}

func (rm *MeasurementParser) newItemFromXLSRow(sh *xlsx.Sheet, row int, catalog *bpu.Catalog) (items []*bpu.Item, err error) {
	boxName := sh.Cell(row, colMeasName).Value
	boxNbFiber := sh.Cell(row, colMeasNbFiber).Value
	boxNbSplice := sh.Cell(row, colMeasNbSplice).Value
	info := fmt.Sprintf("Mesure %s fibres - %s epissures", boxNbFiber, boxNbSplice)

	isDone := sh.Cell(row, colMeasStatus).Value
	var done, todo bool
	switch strings.ToLower(isDone) {
	case "ok":
		done = true
		todo = true
	case "na", "annule", "supprime", "suprime":
		todo = false
	case "", "nok", "ko":
		todo = true
	default:
		err = fmt.Errorf(
			"unknown Status '%s' in cell %s!%s",
			isDone,
			rm.activity,
			xls.RcToAxis(row, colMeasStatus),
		)
		return
	}

	var idate time.Time
	if done {
		date, e := sh.Cell(row, colMeasDate).GetTime(false)
		if e != nil {
			err = fmt.Errorf(
				"could not parse date from '%s' in cell %s!%s",
				sh.Cell(row, colMeasDate).Value,
				rm.activity,
				xls.RcToAxis(row, colMeasDate),
			)
			return
		}
		idate = GetMonday(date)
	}

	// get relevant chapters
	catChapters := catalog.GetCategoryChapters(rm.activity)
	mainChapter, err := catChapters.GetChapterForSize("Mesure", 1)
	if err != nil {
		return
	}
	qty1 := 1

	items = append(items,
		bpu.NewItem(rm.activity, boxName, info, idate, mainChapter, qty1, todo, done),
	)
	return
}

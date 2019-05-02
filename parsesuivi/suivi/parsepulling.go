package suivi

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"strconv"
	"strings"
	"time"
)

type PullingParser struct {
	row      int
	activity string
}

func NewPullingParser() *PullingParser {
	return &PullingParser{
		row:      0,
		activity: sheetTirage,
	}
}

func (pp *PullingParser) Activity() string {
	return pp.activity
}

const (
	colTirageCableType         int = 0
	colTirageTronconName       int = 1
	colTirageTotLength         int = 6
	colTirageLoveLength        int = 7
	colTirageUndergroundLength int = 8
	colTirageAerialLength      int = 9
	colTirageFacadeLength      int = 10
	colTirageStatus            int = 11
	colTirageDate              int = 15

	catPullUnderground string = "Tirage Souterain"
	catPullAerial      string = "Tirage Aérien"
	catPullFacade      string = "Tirage Façade"
)

func (pp *PullingParser) ParseBlock(sh *xlsx.Sheet, catalog *bpu.Catalog, row int) (items []*bpu.Item, err ParsingError, nextRow int) {
	var e error

	cableName := sh.Cell(row, colTirageCableType).Value
	tronconName := sh.Cell(row, colTirageTronconName).Value

	// Check for end of data
	if cableName == "" && tronconName == "" {
		return
	}

	// check for box declaration
	if !(cableName != "" && tronconName != "") {
		err.Add(fmt.Errorf("invalid cable pulling definition in line %s:%d", pp.activity, row+1), true)
		return
	}

	// parse cable pulling declaration line
	items, e = pp.newItemFromXLSRow(sh, row, catalog)
	if e != nil {
		err.Add(e, true)
	}

	// parse remaining block detail lines
	nextRow = row + 1
	for {
		cableName = sh.Cell(nextRow, colTirageCableType).Value
		tronconName = sh.Cell(nextRow, colTirageTronconName).Value
		// check for end of data
		if !(cableName == "" && tronconName != "") {
			break
		}
		nextRow++
	}
	return
}

func (pp *PullingParser) newItemFromXLSRow(sh *xlsx.Sheet, row int, catalog *bpu.Catalog) (items []*bpu.Item, err error) {
	cableType := sh.Cell(row, colTirageCableType).Value
	tronconName := sh.Cell(row, colTirageTronconName).Value
	cableSize, e := getCableSize(cableType)
	if e != nil {
		err = fmt.Errorf(
			"%s in cell %s!%s",
			e.Error(),
			pp.activity,
			xls.RcToAxis(row, colTirageCableType),
		)
		return
	}

	loveLength, e := sh.Cell(row, colTirageLoveLength).Int()
	if e != nil {
		err = fmt.Errorf(
			"could not parse Love Length from '%s' in cell %s!%s",
			sh.Cell(row, colTirageLoveLength).Value,
			pp.activity,
			xls.RcToAxis(row, colTirageLoveLength),
		)
		return
	}

	undergroundLength, e := sh.Cell(row, colTirageUndergroundLength).Int()
	if e != nil {
		err = fmt.Errorf(
			"could not parse Underground Length from '%s' in cell %s!%s",
			sh.Cell(row, colTirageUndergroundLength).Value,
			pp.activity,
			xls.RcToAxis(row, colTirageUndergroundLength),
		)
		return
	}

	aerialLength, e := sh.Cell(row, colTirageAerialLength).Int()
	if e != nil {
		err = fmt.Errorf(
			"could not parse Aerial Length from '%s' in cell %s!%s",
			sh.Cell(row, colTirageAerialLength).Value,
			pp.activity,
			xls.RcToAxis(row, colTirageAerialLength),
		)
		return
	}

	facadeLength, e := sh.Cell(row, colTirageFacadeLength).Int()
	if e != nil {
		err = fmt.Errorf(
			"could not parse Facade Length from '%s' in cell %s!%s",
			sh.Cell(row, colTirageFacadeLength).Value,
			pp.activity,
			xls.RcToAxis(row, colTirageFacadeLength),
		)
		return
	}

	todo, done, err := parseStatus(sh, row, colTirageStatus)
	if err != nil {
		return
	}

	var idate time.Time
	if done {
		date, e := sh.Cell(row, colTirageDate).GetTime(false)
		if e != nil {
			err = fmt.Errorf(
				"could not parse date from '%s' in cell %s!%s",
				sh.Cell(row, colTirageDate).Value,
				pp.activity,
				xls.RcToAxis(row, colTirageDate),
			)
			return
		}
		idate = GetMonday(date)
	}

	// get relevant chapters
	catChapters := catalog.GetCategoryChapters(pp.activity)

	// Item for underground cable pulling
	if loveLength+undergroundLength > 0 {
		chapter, e := catChapters.GetChapterForSize(catPullUnderground, cableSize)
		if e != nil {
			err = fmt.Errorf(
				"could not define bpu chapter: '%s' in line %s:%d",
				e.Error(),
				pp.activity,
				row+1,
			)
			return
		}
		info := fmt.Sprintf("Tirage %s (%dml)", cableType, loveLength+undergroundLength)
		items = append(items,
			bpu.NewItem(pp.activity, tronconName, info, idate, chapter, loveLength+undergroundLength, loveLength+undergroundLength, todo, done),
		)
	}

	// Item for Aerial cable pulling
	if aerialLength+facadeLength > 0 {
		chapter, e := catChapters.GetChapterForSize(catPullAerial, cableSize)
		if e != nil {
			err = fmt.Errorf(
				"could not define bpu chapter: '%s' in line %s:%d",
				e.Error(),
				pp.activity,
				row+1,
			)
			return
		}
		info := fmt.Sprintf("Tirage %s (%dml)", cableType, aerialLength+facadeLength)
		items = append(items,
			bpu.NewItem(pp.activity, tronconName, info, idate, chapter, aerialLength+facadeLength, aerialLength+facadeLength, todo, done),
		)
	}

	// Item for Facade cable pulling
	if facadeLength > 0 {
		chapter, e := catChapters.GetChapterForSize(catPullFacade, cableSize)
		if e != nil {
			err = fmt.Errorf(
				"could not define bpu chapter: '%s' in line %s:%d",
				e.Error(),
				pp.activity,
				row+1,
			)
			return
		}
		info := fmt.Sprintf("Tirage %s (%dml)", cableType, facadeLength)
		items = append(items,
			bpu.NewItem(pp.activity, tronconName, info, idate, chapter, facadeLength, facadeLength, todo, done),
		)
	}

	return
}

func getCableSize(cableType string) (int, error) {
	parts := strings.Split(cableType, "_")
	if len(parts) < 2 {
		return 0, fmt.Errorf("misformatted cable type '%': can not detect _nnFO_ chunk", cableType)
	}
	size, e := strconv.ParseInt(strings.TrimSuffix(parts[1], "FO"), 10, 64)
	if e != nil {
		return 0, fmt.Errorf("misformatted cable type: can not get number of fiber in '%'", parts[1])
	}
	return int(size), nil
}

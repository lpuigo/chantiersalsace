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

type RaccoParser struct {
	row      int
	activity string
}

func NewRaccoParser() *RaccoParser {
	return &RaccoParser{
		row:      0,
		activity: sheetRacco,
	}
}

func (rp *RaccoParser) Activity() string {
	return rp.activity
}

const (
	colRaccoName    int = 0
	colRaccoBoxName int = 2
	colRaccoBoxType int = 3
	colRaccoSize    int = 6
	colRaccoOpe     int = 7
	colRaccoFiber   int = 8
	colRaccoSplice  int = 9
	colRaccoStatus  int = 10
	colRaccoDate    int = 14
)

type box struct {
	nbFiber  int
	nbSplice int
}

func (rp *RaccoParser) ParseBlock(sh *xlsx.Sheet, catalog *bpu.Catalog, row int) (items []*bpu.Item, err ParsingError, nextRow int) {
	var nBox *box
	var e error

	boxName := sh.Cell(row, colRaccoName).Value
	boxOpe := sh.Cell(row, colRaccoOpe).Value
	// Check for end of data
	if boxOpe == "" && boxName == "" {
		return
	}
	// check for box declaration
	if boxOpe != "TOTAL" {
		err.Add(fmt.Errorf("invalid Box definition in line %s:%d", rp.activity, row+1), true)
		return
	}

	// parse box declaration line
	items, nBox, e = rp.newItemFromXLSRow(sh, row, catalog)
	if e != nil {
		err.Add(e, true)
	}

	// parse remaining block detail lines
	var nbFiber, nbSplice int
	nextRow = row + 1
	for {
		boxOpe := sh.Cell(nextRow, colRaccoOpe).Value
		// check for end of data
		if boxOpe == "" || boxOpe == "TOTAL" {
			break
		}

		nbF, nbS, perr := rp.getNbFiberAndSplice(sh, nextRow)
		if perr.HasError() {
			err.Append(perr)
		}
		nbFiber += nbF
		nbSplice += nbS
		nextRow++
	}

	// check Box declaration consistency
	if nBox == nil {
		return
	}
	if nBox.nbFiber != nbFiber {
		err.Add(fmt.Errorf("wrong Nb Fiber for box in cell %s!%s", sheetRacco, xls.RcToAxis(row, colRaccoFiber)), true)
	}
	if nBox.nbSplice != nbSplice {
		err.Add(fmt.Errorf("wrong Nb Splice for box in cell %s!%s", sheetRacco, xls.RcToAxis(row, colRaccoSplice)), true)
	}
	return
}

func (rp *RaccoParser) getNbFiberAndSplice(sh *xlsx.Sheet, row int) (nbFiber, nbSplice int, err ParsingError) {
	nbFiber, e := sh.Cell(row, colRaccoFiber).Int()
	if e != nil {
		err.Add(fmt.Errorf("could not parse Nb Fiber in cell %s!%s", rp.activity, xls.RcToAxis(row, colRaccoFiber)), true)
	}

	nbSplice, e = sh.Cell(row, colRaccoSplice).Int()
	if e != nil {
		err.Add(fmt.Errorf("could not parse Nb Splice in cell %s!%s", rp.activity, xls.RcToAxis(row, colRaccoSplice)), true)
	}
	return
}

func (rp *RaccoParser) newItemFromXLSRow(sh *xlsx.Sheet, row int, catalog *bpu.Catalog) (items []*bpu.Item, nBox *box, err error) {
	var mainChapter, optChapter *bpu.Article
	var qty1, qty2 int
	var e error

	name := sh.Cell(row, colRaccoName).Value
	boxType := sh.Cell(row, colRaccoBoxName).Value
	boxCatergory := sh.Cell(row, colRaccoBoxType).Value

	if !catalog.IsBoxDefined(boxCatergory, boxType) {
		err = fmt.Errorf(
			"unknown Box Type '%s' for Category '%s' in line %s:%d",
			boxType,
			boxCatergory,
			rp.activity,
			row+1,
		)
		return
	}

	size := sh.Cell(row, colRaccoSize).Value
	if !strings.HasSuffix(size, "FO") {
		err = fmt.Errorf(
			"unexpected Box Size format '%s' in cell %s!%s",
			size,
			rp.activity,
			xls.RcToAxis(row, colRaccoSize),
		)
		return
	}
	isize, e := strconv.ParseInt(strings.TrimSuffix(size, "FO"), 10, 64)
	if e != nil {
		err = fmt.Errorf(
			"could not parse '%s' Box Size in cell %s!%s",
			strings.TrimSuffix(size, "FO"),
			rp.activity,
			xls.RcToAxis(row, colRaccoSize),
		)
		return
	}
	boxSize := int(isize)
	info := fmt.Sprintf("Install. %s: %s (%dFO)", boxCatergory, boxType, boxSize)

	nbFiber, e := sh.Cell(row, colRaccoFiber).Int()
	if e != nil {
		err = fmt.Errorf(
			"could not parse Bpe Nb Fiber from '%s' in cell %s!%s",
			sh.Cell(row, colRaccoFiber).Value,
			rp.activity,
			xls.RcToAxis(row, colRaccoFiber),
		)
		return
	}

	nbSplice, e := sh.Cell(row, colRaccoSplice).Int()
	if e != nil {
		err = fmt.Errorf(
			"could not parse Bpe Nb Splice from '%s' in cell %s!%s",
			sh.Cell(row, colRaccoSplice).Value,
			rp.activity,
			xls.RcToAxis(row, colRaccoSplice),
		)
		return
	}

	nBox = &box{
		nbFiber:  nbFiber,
		nbSplice: nbSplice,
	}

	isDone := sh.Cell(row, colRaccoStatus).Value
	var done, todo bool
	switch strings.ToLower(isDone) {
	case "ok":
		done = true
		todo = true
	case "na", "annule", "supprime", "suprime":
		todo = false
	case "", "nok", "ko", "blocage", "en cours":
		todo = true
	default:
		err = fmt.Errorf(
			"unknown Status '%s' in cell %s!%s",
			isDone,
			rp.activity,
			xls.RcToAxis(row, colRaccoStatus),
		)
		return
	}

	var idate time.Time
	if done {
		date, e := sh.Cell(row, colRaccoDate).GetTime(false)
		if e != nil {
			err = fmt.Errorf(
				"could not parse date from '%s' in cell %s!%s",
				sh.Cell(row, colRaccoDate).Value,
				rp.activity,
				xls.RcToAxis(row, colRaccoDate),
			)
			return
		}
		idate = GetMonday(date)
	}

	// get relevant chapters
	catChapters := catalog.GetCategoryChapters(rp.activity)
	uBoxCategory := strings.ToUpper(boxCatergory)
	switch uBoxCategory {
	case "PM":
		mainChapter, optChapter = catChapters["PM"][0], catChapters["PM"][1]
		qty1 = nbSplice / mainChapter.Size
		// check for missing modules
		qty2 = 0
		if qty1*mainChapter.Size < nbSplice {
			qty1++
			nbMissingSplice := qty1*mainChapter.Size - nbSplice
			qty2 = nbMissingSplice / optChapter.Size
		}

	case "BPE", "PBO":
		mainChapter, optChapter, e = getRaccoBoxChapters(catalog, rp.activity, boxCatergory, boxType)
		qty1, qty2 = 1, nbSplice
		if e != nil {
			panic(e.Error())
		}

	default:
		err = fmt.Errorf("illegal box category '%s'", boxCatergory)
		return
	}

	items = append(items,
		bpu.NewItem(rp.activity, name, info, idate, mainChapter, qty1, todo, done),
	)
	if optChapter != nil {
		items = append(items,
			bpu.NewItem(rp.activity, name, info, idate, optChapter, qty2, todo, done),
		)
	}
	return
}

// getRaccoBoxChapters returns Article applicable for given Bpe or Pbo type
func getRaccoBoxChapters(catalog *bpu.Catalog, activity, cat, boxType string) (boxChapter, spliceChapter *bpu.Article, err error) {
	// box lookup
	box := catalog.GetBox(cat, boxType)
	if box == nil {
		err = fmt.Errorf("unknow box type '%s' for category '%s'", boxType, cat)
		return
	}

	catChapters := catalog.GetCategoryChapters(activity)
	if catChapters == nil {
		err = fmt.Errorf("unknow activity '%s'", activity)
		return
	}

	switch strings.ToUpper(cat) {
	case "BPE":
		boxChapter, err = catChapters.GetChapterForSize(cat, box.Size)
		if err != nil {
			return
		}
		spliceChapter, err = catChapters.GetChapterForSize(cat+" Splice", box.Size)
		if err != nil {
			return
		}
	case "PBO":
		boxChapter, err = catChapters.GetChapterForSize(cat+" "+box.Usage, box.Size)
		if err != nil {
			return
		}
	default:
		err = fmt.Errorf("category '%s' is not handled", cat)
		return
	}
	return
}

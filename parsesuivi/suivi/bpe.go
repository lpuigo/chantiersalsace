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

type Bpe struct {
	Name        string
	Type        string
	Category    string
	Size        int
	NbFiber     int
	NbSplice    int
	BpeValue    float64
	SpliceValue float64
	PriceName   string
	ToDo        bool
	Done        bool
	Date        time.Time
}

func NewItemFromRaccoXLSRow(sh *xlsx.Sheet, row int, catalog *bpu.Catalog) ([]*bpu.Item, error) {
	var chapter1, chapter2 *bpu.Chapter
	var nb1, nb2 int
	var e error

	activity := "Racco"
	name := sh.Cell(row, colRaccoName).Value
	boxType := sh.Cell(row, colRaccoBoxName).Value
	boxCatergory := sh.Cell(row, colRaccoBoxType).Value

	if !catalog.IsBoxDefined(boxCatergory, boxType) {
		return nil, fmt.Errorf(
			"unknown Bpe Type '%s'for Category '%s' in line %s:%d",
			boxType,
			boxCatergory,
			sheetRacco,
			row+1,
		)
	}

	size := sh.Cell(row, colRaccoSize).Value
	if !strings.HasSuffix(size, "FO") {
		return nil, fmt.Errorf(
			"unexpected Box Size format '%s' in cell %s!%s",
			size,
			sheetRacco,
			xls.RcToAxis(row, colRaccoSize),
		)
	}
	isize, e := strconv.ParseInt(strings.TrimSuffix(size, "FO"), 10, 64)
	if e != nil {
		return nil, fmt.Errorf(
			"could not parse Box Size from '%s' in cell %s!%s",
			strings.TrimSuffix(size, "FO"),
			sheetRacco,
			xls.RcToAxis(row, colRaccoSize),
		)
	}
	boxSize := int(isize)
	info := fmt.Sprintf("%s - %s (%dFO)", boxCatergory, boxType, boxSize)

	nbSplice, e := sh.Cell(row, colRaccoSplice).Int()
	if e != nil {
		return nil, fmt.Errorf(
			"could not parse Bpe Nb Splice from '%s' in cell %s!%s",
			sh.Cell(row, colRaccoSplice).Value,
			sheetRacco,
			xls.RcToAxis(row, colRaccoSplice),
		)
	}

	isDone := sh.Cell(row, colRaccoStatus).Value
	var done, todo bool
	switch strings.ToLower(isDone) {
	case "ok":
		done = true
		todo = true
	case "na", "annule", "supprime":
		todo = false
	default:
		todo = true
	}
	var idate time.Time
	if done {
		date, e := sh.Cell(row, colRaccoDate).GetTime(false)
		if e != nil {
			err := fmt.Errorf(
				"could not parse Box Size from '%s' in cell %s!%s",
				strings.TrimSuffix(size, "FO"),
				sheetRacco,
				xls.RcToAxis(row, colRaccoSize),
			)
			return nil, err
		}
		idate = GetMonday(date)
	}

	// get relevant chapter
	switch strings.ToUpper(boxCatergory) {
	case "PM":
		chapter1, chapter2 = catalog.GetRaccoPmPrices()
		nb1 = nbSplice / chapter1.Size
		// check for missing modules
		nb2 = 0
		if nb1*chapter1.Size < nbSplice {
			nb1++
			nbMissingSplice := nb1*chapter1.Size - nbSplice
			nb2 = nbMissingSplice / chapter2.Size
		}

	case "BPE", "PBO":
		chapter1, chapter2, e = catalog.GetRaccoBoxPrice(boxCatergory, boxType)
		nb1, nb2 = 1, nbSplice
		if e != nil {
			panic(e.Error())
		}
	default:
		return nil, fmt.Errorf(
			"undefined Racco configuration found in line %s!%d",
			sh.Cell(row, colRaccoSplice).Value,
			sheetRacco,
			row+1,
		)
	}

	res := []*bpu.Item{
		bpu.NewItem(activity, name, info, idate, chapter1, nb1, todo, done),
	}
	if chapter2 != nil {
		res = append(res,
			bpu.NewItem(activity, name, info, idate, chapter2, nb2, todo, done),
		)
	}
	return res, nil
}

func NewBpeFromXLSRow(sh *xlsx.Sheet, row int) (*Bpe, error) {
	bpe := &Bpe{}
	bpe.Name = sh.Cell(row, colRaccoName).Value
	bpe.Type = sh.Cell(row, colRaccoBoxName).Value
	bpe.Category = sh.Cell(row, colRaccoBoxType).Value
	size := sh.Cell(row, colRaccoSize).Value
	if !strings.HasSuffix(size, "FO") {
		return nil, fmt.Errorf("unexpected Bpe Size format '%s' in cell %s!%s", size, sheetRacco, xls.RcToAxis(row, colRaccoSize))
	}
	isize, e := strconv.ParseInt(strings.TrimSuffix(size, "FO"), 10, 64)
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe Size from '%s' in cell %s!%s", strings.TrimSuffix(size, "FO"), sheetRacco, xls.RcToAxis(row, colRaccoSize))
	}
	bpe.Size = int(isize)

	bpe.NbFiber, e = sh.Cell(row, colRaccoFiber).Int()
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe Nb Fiber from '%s' in cell %s!%s", sh.Cell(row, colRaccoFiber).Value, sheetRacco, xls.RcToAxis(row, colRaccoFiber))
	}

	bpe.NbSplice, e = sh.Cell(row, colRaccoSplice).Int()
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe Nb Splice from '%s' in cell %s!%s", sh.Cell(row, colRaccoSplice).Value, sheetRacco, xls.RcToAxis(row, colRaccoSplice))
	}

	done := sh.Cell(row, colRaccoStatus).Value
	switch strings.ToLower(done) {
	case "ok":
		bpe.Done = true
		bpe.ToDo = true
	case "na", "annule", "supprime":
		bpe.ToDo = false
	default:
		bpe.ToDo = true
	}

	if !bpe.Done {
		return bpe, nil
	}

	date, e := sh.Cell(row, colRaccoDate).GetTime(false)
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe End Date from '%s' in cell %s!%s", row+1, sh.Cell(row, colRaccoDate).Value, sheetRacco, xls.RcToAxis(row, colRaccoDate))
	}
	bpe.Date = GetMonday(date)
	return bpe, nil
}

func (b *Bpe) CheckFiber(fiber int) bool {
	return b.NbFiber == fiber
}

func (b *Bpe) CheckSplice(splice int) bool {
	return b.NbSplice == splice
}

func (b *Bpe) IsSro() bool {
	return strings.HasPrefix(strings.ToLower(b.Category), "pm")
}

func (b *Bpe) SetValues(pc *bpu.Catalog) {
	if b.IsSro() {
		b.SetSroValues(pc)
		return
	}
	pBox, pSplice, err := pc.GetRaccoBoxPrice(b.Category, b.Type)
	if err != nil {
		panic(err.Error())
	}

	b.PriceName = pBox.Name
	b.BpeValue = pBox.Price
	if pSplice != nil {
		b.SpliceValue = pSplice.Price
	}
}

func (b *Bpe) GetSroNumbers(pc *bpu.Catalog) (nbSro, nbMissingModule int) {
	p, mp := pc.GetRaccoPmPrices()
	b.PriceName = p.Name
	nbSro = b.NbSplice / p.Size

	// check for missing modules
	nbMissingModule = 0
	if nbSro*p.Size < b.NbSplice {
		nbSro++
		nbMissingSplice := nbSro*p.Size - b.NbSplice
		nbMissingModule = nbMissingSplice / mp.Size
	}
	return
}

func (b *Bpe) SetSroValues(pc *bpu.Catalog) {
	p, mp := pc.GetRaccoPmPrices()
	nbSro, nbMissingModule := b.GetSroNumbers(pc)
	b.BpeValue = p.GetBpeValue()*float64(nbSro) + mp.GetBpeValue()*float64(nbMissingModule)
}

func (b *Bpe) String() string {
	date := ""
	if !b.Date.IsZero() {
		date = fmt.Sprintf("  Date:%s", b.Date.Format("06-01-02"))
	}
	return fmt.Sprintf("Size:%3d (%7.2f€) Fiber:%3d  Splice:%3d (%8.2f€)  ToDo:%6t  Done:%6t%s",
		b.Size,
		b.BpeValue,
		b.NbFiber,
		b.NbSplice,
		b.SpliceValue,
		b.ToDo,
		b.Done,
		date)
}

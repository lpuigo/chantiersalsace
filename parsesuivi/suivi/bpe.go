package suivi

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/tealeg/xlsx"
	"strconv"
	"strings"
	"time"
)

type Bpe struct {
	Name        string
	Type        string
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

func NewBpeFromXLSRow(sh *xlsx.Sheet, row int) (*Bpe, error) {
	bpe := &Bpe{}
	bpe.Name = sh.Cell(row, colBpeName).Value
	bpe.Type = sh.Cell(row, colBpeType).Value
	size := sh.Cell(row, colBpeSize).Value
	if !strings.HasSuffix(size, "FO") {
		return nil, fmt.Errorf("unexpected Bpe Size format '%s' line %d", size, row+1)
	}
	isize, e := strconv.ParseInt(strings.TrimSuffix(size, "FO"), 10, 64)
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe Size from '%s' line %d", strings.TrimSuffix(size, "FO"), row+1)
	}
	bpe.Size = int(isize)

	bpe.NbFiber, e = sh.Cell(row, colBpeFiber).Int()
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe Nb Fiber from '%s' line %d", sh.Cell(row, colBpeFiber).Value, row+1)
	}

	bpe.NbSplice, e = sh.Cell(row, colBpeSplice).Int()
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe Nb Splice from '%s' line %d", sh.Cell(row, colBpeSplice).Value, row+1)
	}

	done := sh.Cell(row, colBpeStatus).Value
	switch strings.ToLower(done) {
	case "ok":
		bpe.Done = true
		bpe.ToDo = true
	case "na":
		bpe.ToDo = false
	default:
		bpe.ToDo = true
	}
	if strings.ToLower(done) == "ok" {

	}

	if !bpe.Done {
		return bpe, nil
	}

	date, e := sh.Cell(row, colBpeDate).GetTime(false)
	if e != nil {
		return nil, fmt.Errorf("could not parse Bpe End Date from '%s' line %d", sh.Cell(row, colBpeDate).Value, row+1)
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
	return strings.HasPrefix(strings.ToLower(b.Name), "sro")
}

func (b *Bpe) SetValues(pc *bpu.Bpu) {
	if b.IsSro() {
		b.SetSroValues(pc)
		return
	}
	p := pc.GetBpePrice(b.Type)
	b.PriceName = p.Name
	b.BpeValue = p.GetBpeValue()
	b.SpliceValue = p.GetSpliceValue(b.NbSplice)
}

func (b *Bpe) SetSroValues(pc *bpu.Bpu) {
	p, mp := pc.GetSroPrice()
	b.PriceName = p.Name
	nb := b.NbSplice / p.Size
	b.BpeValue = p.GetBpeValue() * float64(nb)

	// check for missing modules
	if nb*p.Size < b.NbSplice {
		nbMissingSplice := (nb+1)*p.Size - b.NbSplice
		nbMissingModule := nbMissingSplice / mp.Size
		b.BpeValue += p.GetBpeValue() + mp.GetBpeValue()*float64(nbMissingModule)
	}
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

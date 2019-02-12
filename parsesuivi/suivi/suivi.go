package suivi

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"time"
)

type Suivi struct {
	Bpes      []*Bpe
	BeginDate time.Time
	LastDate  time.Time
}

func (s *Suivi) String() string {
	res := ""
	for id, bpe := range s.Bpes {
		res += fmt.Sprintf("%3d:%s\n", id, bpe.String())
	}
	return res
}

func (s *Suivi) Add(b *Bpe) {
	s.Bpes = append(s.Bpes, b)
	if s.BeginDate.IsZero() {
		s.BeginDate = GetMonday(time.Now())
		s.LastDate = s.BeginDate
	}
	if !b.Done {
		return
	}
	if b.Date.Before(s.BeginDate) {
		s.BeginDate = b.Date
	}
	if b.Date.After(s.LastDate) {
		s.LastDate = b.Date
	}
}

const (
	colBpeName   int = 1
	colBpeSize   int = 6
	colBpeOpe    int = 7
	colBpeFiber  int = 8
	colBpeSplice int = 9
	colBpeStatus int = 10
	colBpeDate   int = 14
)

func NewSuiviFromXLS(file string, pc *bpu.Bpu) (s *Suivi, err error) {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		return
	}
	if len(xf.Sheets) < 2 {
		err = fmt.Errorf("could not find sheet with Bpe details in '%s'", file)
		return
	}
	bsh := xf.Sheets[1]

	s = &Suivi{}

	row := 0
	inProgress := true
	var nbFiber, nbSplice, bpeRow int
	var bpe *Bpe
	for inProgress {
		row++
		bpeName := bsh.Cell(row, colBpeName).Value
		bpeOpe := bsh.Cell(row, colBpeOpe).Value
		if bpeOpe == "" && bpeName == "" {
			inProgress = false
			continue
		}
		if bpeName != "" { // This Row contains BPE definition, parse and process it
			if bpe != nil {
				if !bpe.CheckFiber(nbFiber) {
					err = fmt.Errorf("invalid Nb Fiber for Bpe on line %d", bpeRow+1)
					return
				}
				if !bpe.CheckSplice(nbSplice) {
					err = fmt.Errorf("invalid Nb Splice for Bpe on line %d", bpeRow+1)
					return
				}
			}
			bpeRow = row
			bpe, err = NewBpeFromXLSRow(bsh, row)
			if err != nil {
				return
			}
			bpe.SetValues(pc)
			s.Add(bpe)
			nbFiber = 0
			nbSplice = 0
			continue
		}
		// this row contains BPE detail, check for fiber and splice number

		sf := bsh.Cell(row, colBpeFiber).Value
		if sf != "" {
			nbf, e := bsh.Cell(row, colBpeFiber).Int()
			if e != nil {
				err = fmt.Errorf("could not parse Nb Fiber from '%s' line %d", sf, row+1)
				return
			}
			nbFiber += nbf
		}
		ss := bsh.Cell(row, colBpeSplice).Value != ""
		if ss {
			nbs, e := bsh.Cell(row, colBpeSplice).Int()
			if e != nil {
				err = fmt.Errorf("could not parse Nb Splice from '%s' line %d", ss, row+1)
				return
			}
			nbSplice += nbs
		}
	}

	return
}

const (
	suiviSheetName string = "Suivi"
)

func (s *Suivi) WriteSuiviXLS(file string) error {
	xf, err := excelize.OpenFile(file)
	if err != nil {
		return err
	}

	if xf.GetSheetIndex(suiviSheetName) == 0 {
		xf.NewSheet(suiviSheetName)
	}

	fTodo := func(bpe *Bpe) bool { return bpe.ToDo == true }
	fNbBpe := func(bpe *Bpe) int { return 1 }
	fNbFiber := func(bpe *Bpe) int { return bpe.NbFiber }
	fNbSplice := func(bpe *Bpe) int { return bpe.NbSplice }
	fValue := func(bpe *Bpe) float64 { return bpe.BpeValue + bpe.SpliceValue }
	nbBPE := s.CountInt(fNbBpe, fTodo)
	nbFiber := s.CountInt(fNbFiber, fTodo)
	nbSplice := s.CountInt(fNbSplice, fTodo)
	nbValue := s.CountFloat(fValue, fTodo)

	// Set Dates header
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(0, 0), "Dates")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, 0), "Nb BPE Total")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(2, 0), "Nb BPE Installés")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(3, 0), "Nb Fibre Total")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(4, 0), "Nb Fibre Installées")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(5, 0), "Nb Epissure Total")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(6, 0), "Nb Epissure Effectuées")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(7, 0), "Valeur € Total")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(8, 0), "Valeur € Réalisée")
	dates := s.Dates()
	for i, d := range dates {
		fDone := func(bpe *Bpe) bool { return bpe.ToDo && bpe.Done && !bpe.Date.After(d) }
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(0, i+1), d)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, i+1), nbBPE)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(2, i+1), s.CountInt(fNbBpe, fDone))
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(3, i+1), nbFiber)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(4, i+1), s.CountInt(fNbFiber, fDone))
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(5, i+1), nbSplice)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(6, i+1), s.CountInt(fNbSplice, fDone))
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(7, i+1), nbValue)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(8, i+1), s.CountFloat(fValue, fDone))
	}
	return xf.Save()
}

func (s *Suivi) Dates() []time.Time {
	res := []time.Time{}
	for d := s.BeginDate; d.Before(s.LastDate); d = d.AddDate(0, 0, 7) {
		res = append(res, d)
	}
	return res
}

func (s *Suivi) CountInt(val func(bpe *Bpe) int, filter func(bpe *Bpe) bool) int {
	res := 0
	for _, b := range s.Bpes {
		if filter(b) {
			res += val(b)
		}
	}
	return res
}

func (s *Suivi) CountFloat(val func(bpe *Bpe) float64, filter func(bpe *Bpe) bool) float64 {
	res := 0.0
	for _, b := range s.Bpes {
		if filter(b) {
			res += val(b)
		}
	}
	return res
}
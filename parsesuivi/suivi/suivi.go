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
	Items      []*bpu.Item
	BeginDate time.Time
	LastDate  time.Time
}

func (s *Suivi) String() string {
	res := ""
	for id, item := range s.Items {
		res += fmt.Sprintf("%3d:%s\n", id, item.String())
	}
	return res
}

func (s *Suivi) Add(items ...*bpu.Item) {
	s.Items = append(s.Items, items...)
	for _, item := range items {
		if s.BeginDate.IsZero() {
			s.BeginDate = GetMonday(time.Now())
			s.LastDate = s.BeginDate
		}
		if !item.Done {
			return
		}
		if item.Date.Before(s.BeginDate) {
			s.BeginDate = item.Date
		}
		if item.Date.After(s.LastDate) {
			s.LastDate = item.Date
		}
	}
}

const (
	sheetTirage  string = "Tirage"
	sheetRacco   string = "Racco"
	sheetMeasure string = "Mesures"
)

func NewSuiviFromXLS(file string, catalog *bpu.Catalog) (s *Suivi, err error) {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		return
	}
	// TODO parse Tirage Tab sheetTirage
	// parse Racco Tab sheetRacco
	bsh := xf.Sheet[sheetRacco]
	if bsh == nil {
		err = fmt.Errorf("onglet '%s' introuvable", sheetRacco)
		return
	}
	s, err = parseRaccoTab(bsh, catalog)

	//TODO parse Mesure Tab sheetMeasure

	return
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

func parseRaccoTab(bsh *xlsx.Sheet, catalog *bpu.Catalog) (s *Suivi, err Error) {
	s = &Suivi{}
	parseErr := Error{}

	row := 0
	inProgress := true
	var nbFiber, nbSplice, bpeRow int
	var bpe *Bpe
	for inProgress {
		row++
		bpeName := bsh.Cell(row, colRaccoName).Value
		bpeOpe := bsh.Cell(row, colRaccoOpe).Value
		if bpeOpe == "" && bpeName == "" {
			inProgress = false
			continue
		}
		if bpeName != "" { // This Row contains BPE definition, parse and process it
			// check finished BPE
			if bpe != nil {
				if !bpe.CheckFiber(nbFiber) {
					parseErr.Add(fmt.Errorf("invalid Nb Fiber for Bpe in cell %s!%s", sheetRacco, xls.RcToAxis(bpeRow, colRaccoFiber)))
				}
				if !bpe.CheckSplice(nbSplice) {
					parseErr.Add(fmt.Errorf("invalid Nb Splice for Bpe in cell %s!%s", sheetRacco, xls.RcToAxis(bpeRow, colRaccoSplice)))
				}
			}
			bpeRow = row
			items, e := NewItemFromRaccoXLSRow(bsh, row, catalog)
			if e != nil {
				parseErr.Add(e)
			} else {
				s.Add(items...)
				nbFiber = 0
				nbSplice = 0
			}
			continue
		}
		// this row contains BPE detail, check for fiber and splice number

		sf := bsh.Cell(row, colRaccoFiber).Value
		if sf != "" {
			nbf, e := bsh.Cell(row, colRaccoFiber).Int()
			if e != nil {
				parseErr.Add(fmt.Errorf("could not parse Nb Fiber in cell %s!%s", sheetRacco, xls.RcToAxis(row, colRaccoFiber)))
			}
			nbFiber += nbf
		}
		ss := bsh.Cell(row, colRaccoSplice).Value != ""
		if ss {
			nbs, e := bsh.Cell(row, colRaccoSplice).Int()
			if e != nil {
				parseErr.Add(fmt.Errorf("could not parse Nb Splice in cell %s!%s", sheetRacco, xls.RcToAxis(row, colRaccoSplice)))
			}
			nbSplice += nbs
		}
	}

	if parseErr.HasError() {
		err = parseErr
	}
	return
}

const (
	suiviSheetName      string = "Suivi"
	progressSheetName   string = "Avancement"
	attachmentSheetName string = "Attachement"
)

func (s *Suivi) WriteSuiviXLS(file string) error {
	xf, err := excelize.OpenFile(file)
	if err != nil {
		return err
	}
	s.writeSuiviSheet(xf)
	return xf.Save()
}

func (s *Suivi) WriteAttachmentXLS(file string, priceCatalog *bpu.Catalog) error {
	xf, err := excelize.OpenFile(file)
	if err != nil {
		return err
	}

	s.writeAttachmentSheet(xf, priceCatalog)
	s.writeProgressSheet(xf)
	xf.UpdateLinkedValue()
	return xf.Save()
}

func (s *Suivi) writeSuiviSheet(xf *excelize.File) {
	if xf.GetSheetIndex(suiviSheetName) == 0 {
		xf.NewSheet(suiviSheetName)
	}

	fTodo := func(bpe *Bpe) bool { return bpe.ToDo }
	fNbBpe := func(bpe *Bpe) int { return 1 }
	fNbFiber := func(bpe *Bpe) int { return bpe.NbFiber }
	fNbSplice := func(bpe *Bpe) int { return bpe.NbSplice }
	fValue := func(bpe *Bpe) float64 { return bpe.BpeValue + bpe.SpliceValue }

	nbBPE := s.CountInt(fNbBpe, fTodo)
	nbFiber := s.CountInt(fNbFiber, fTodo)
	nbSplice := s.CountInt(fNbSplice, fTodo)
	nbValue := s.CountFloat(fValue, fTodo)

	// Set Dates header
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(0, 0), "Semaines")
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
}

func (s *Suivi) writeAttachmentSheet(xf *excelize.File, priceCatalog *bpu.Catalog) {
	if xf.GetSheetIndex(attachmentSheetName) == 0 {
		xf.NewSheet(attachmentSheetName)
	}

	row := 0
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 0), "Ref.")
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 1), "Prix Unit. Boitier")
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 2), "Quantité Boitier")
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 3), "Boitiers Réalisés")
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 4), "Prix Unit. Epissure")
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 5), "Quantité Epissure")
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 6), "Epissures Réalisées")
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 7), "Montant HT")

	row++
	// SRO
	ps, pm := priceCatalog.GetRaccoPmPrices()
	fTDSro := func(bpe *Bpe) bool { return bpe.ToDo && bpe.IsSro() }
	fSro := func(bpe *Bpe) bool { return bpe.ToDo && bpe.Done && bpe.IsSro() }
	fNbSro := func(bpe *Bpe) int {
		nbSro, _ := bpe.GetSroNumbers(priceCatalog)
		return nbSro
	}
	fNbSroMissingModule := func(bpe *Bpe) int {
		_, nbMissingModule := bpe.GetSroNumbers(priceCatalog)
		return nbMissingModule
	}
	fSroValue := func(bpe *Bpe) float64 { return bpe.BpeValue }
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 0), ps.Name)
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 1), ps.Price)
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 2), s.CountInt(fNbSro, fTDSro))
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 3), s.CountInt(fNbSro, fSro))
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 4), pm.Price)
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 5), s.CountInt(fNbSroMissingModule, fTDSro))
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 6), s.CountInt(fNbSroMissingModule, fSro))
	xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 7), s.CountFloat(fSroValue, fSro))

	// Bpe
	fNbBpe := func(bpe *Bpe) int { return 1 }
	fNbSplice := func(bpe *Bpe) int { return bpe.NbSplice }
	fValue := func(bpe *Bpe) float64 { return bpe.BpeValue + bpe.SpliceValue }
	for _, priceCat := range priceCatalog.Chapters {
		for _, p := range priceCat {
			row++
			fTDBpe := func(bpe *Bpe) bool { return bpe.ToDo && bpe.PriceName == p.Name }
			fBpe := func(bpe *Bpe) bool { return bpe.ToDo && bpe.Done && bpe.PriceName == p.Name }
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 0), p.Name)
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 1), p.Price)
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 2), s.CountInt(fNbBpe, fTDBpe))
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 3), s.CountInt(fNbBpe, fBpe))
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 4), p.GetSpliceValue(1))
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 5), s.CountInt(fNbSplice, fTDBpe))
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 6), s.CountInt(fNbSplice, fBpe))
			xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 7), s.CountFloat(fValue, fBpe))
		}
	}

}

func (s *Suivi) writeProgressSheet(xf *excelize.File) {
	if xf.GetSheetIndex(progressSheetName) == 0 {
		xf.NewSheet(progressSheetName)
	}

	row := 0
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 0), "Bpe")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 1), "Type Boitier")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 2), "Taille Tronçon")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 3), "Nb Epissure")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 4), "Ref Catalogue")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 5), "Installé")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 6), "Semaine")
	for _, b := range s.Bpes {
		if !b.ToDo {
			continue
		}
		row++
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 0), b.Name)
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 1), b.Type)
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 2), fmt.Sprintf("%dFO", b.Size))
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 3), b.NbSplice)
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 4), b.PriceName)
		if b.Done {
			xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 5), "Oui")
			xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 6), b.Date)
		} else {
			xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 5), "")
			xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 6), "")
		}
	}
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

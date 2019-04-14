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
	Items     []*bpu.Item
	BeginDate time.Time
	LastDate  time.Time
	Catalog   *bpu.Catalog
}

func NewSuivi(catalog *bpu.Catalog) *Suivi {
	s := &Suivi{}
	s.BeginDate = GetMonday(time.Now())
	s.LastDate = s.BeginDate
	s.Catalog = catalog
	return s
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
		if !item.Done {
			continue
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

	s = NewSuivi(catalog)

	// TODO parse Tirage Tab sheetTirage
	rsh := xf.Sheet[sheetTirage]
	if rsh == nil {
		err = fmt.Errorf("onglet '%s' introuvable", sheetTirage)
		return
	}
	perr := ParseTab(rsh, NewPullingParser(), s)
	if perr.HasError() {
		return nil, perr
	}

	// parse Racco Tab sheetRacco
	rsh = xf.Sheet[sheetRacco]
	if rsh == nil {
		err = fmt.Errorf("onglet '%s' introuvable", sheetRacco)
		return
	}
	perr = ParseTab(rsh, NewRaccoParser(), s)
	if perr.HasError() {
		return nil, perr
	}

	// parse Mesure Tab sheetMeasure
	msh := xf.Sheet[sheetMeasure]
	if msh == nil {
		err = fmt.Errorf("onglet '%s' introuvable", sheetMeasure)
		return
	}
	perr = ParseTab(msh, NewMeasurementParser(), s)
	if perr.HasError() {
		return nil, perr
	}

	return
}

const (
	suiviSheetName      string = "Suivi"
	progressSheetName   string = "Avancement"
	attachmentSheetName string = "Attachement"
	global              string = "Global"
)

func (s *Suivi) WriteSuiviXLS(file string) error {
	xf, err := excelize.OpenFile(file)
	if err != nil {
		return err
	}
	s.writeSuiviSheet(xf)
	s.writeProgressSheet(xf)
	return xf.Save()
}

func (s *Suivi) WriteAttachmentXLS(file string) error {
	xf, err := excelize.OpenFile(file)
	if err != nil {
		return err
	}

	s.writeAttachmentSheet(xf)
	xf.UpdateLinkedValue()
	return xf.Save()
}

func (s *Suivi) writeSuiviSheet(xf *excelize.File) {
	if xf.GetSheetIndex(suiviSheetName) == 0 {
		xf.NewSheet(suiviSheetName)
	}

	fAll := func(item *bpu.Item) bool { return true }
	fTodo := func(item *bpu.Item) bool { return item.Todo }
	fNbQty := func(item *bpu.Item) int { return item.Quantity }
	fPrice := func(item *bpu.Item) float64 { return item.Price() }

	articleNames := s.Catalog.GetArticleNames(sheetTirage)
	articleNames = append(articleNames, s.Catalog.GetArticleNames(sheetRacco)...)
	articleNames = append(articleNames, s.Catalog.GetArticleNames(sheetMeasure)...)

	nbTot := make(map[string]int)
	valTot := make(map[string]float64)
	valTot[global] = s.CountFloat(s.Items, fPrice, fTodo)
	articleNameItems := s.GetItems(fTodo)
	for _, articleName := range articleNames {
		nbTot[articleName] = s.CountInt(articleNameItems[articleName], fNbQty, fAll)
		valTot[articleName] = s.CountFloat(articleNameItems[articleName], fPrice, fAll)
	}

	// Set Dates header
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(0, 1), "Semaines")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, 0), "Suivi financier")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, 1), "€ Total")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(2, 1), "€ Fait")
	row := 2
	for _, articleName := range articleNames {
		if nbTot[articleName] == 0 {
			continue
		}
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+1, 0), articleName)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+1, 1), "Nb Total")
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+2, 1), "Nb")
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+3, 1), "€ Total")
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+4, 1), "€ Fait")
		row += 4
	}

	for col, d := range s.Dates() {
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(0, col+2), d)
		fDone := func(item *bpu.Item) bool { return item.Done && !item.Date.After(d) }
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, col+2), valTot[global])
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(2, col+2), s.CountFloat(s.Items, fPrice, fDone))
		row := 2
		for _, articleName := range articleNames {
			if nbTot[articleName] == 0 {
				continue
			}
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+1, col+2), nbTot[articleName])
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+2, col+2), s.CountInt(articleNameItems[articleName], fNbQty, fDone))
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+3, col+2), valTot[articleName])
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+4, col+2), s.CountFloat(articleNameItems[articleName], fPrice, fDone))
			row += 4
		}
	}
}

func (s *Suivi) writeAttachmentSheet(xf *excelize.File) {
	//if xf.GetSheetIndex(attachmentSheetName) == 0 {
	//	xf.NewSheet(attachmentSheetName)
	//}
	//
	//row := 0
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 0), "Ref.")
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 1), "Prix Unit. Boitier")
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 2), "Quantité Boitier")
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 3), "Boitiers Réalisés")
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 4), "Prix Unit. Epissure")
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 5), "Quantité Epissure")
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 6), "Epissures Réalisées")
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 7), "Montant HT")
	//
	//row++
	//// SRO
	//ps, pm := priceCatalog.GetRaccoPmPrices()
	//fTDSro := func(bpe *Bpe) bool { return bpe.ToDo && bpe.IsSro() }
	//fSro := func(bpe *Bpe) bool { return bpe.ToDo && bpe.Done && bpe.IsSro() }
	//fNbSro := func(bpe *Bpe) int {
	//	nbSro, _ := bpe.GetSroNumbers(priceCatalog)
	//	return nbSro
	//}
	//fNbSroMissingModule := func(bpe *Bpe) int {
	//	_, nbMissingModule := bpe.GetSroNumbers(priceCatalog)
	//	return nbMissingModule
	//}
	//fSroValue := func(bpe *Bpe) float64 { return bpe.BpeValue }
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 0), ps.Name)
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 1), ps.Price)
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 2), s.CountInt(fNbSro, fTDSro))
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 3), s.CountInt(fNbSro, fSro))
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 4), pm.Price)
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 5), s.CountInt(fNbSroMissingModule, fTDSro))
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 6), s.CountInt(fNbSroMissingModule, fSro))
	//xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 7), s.CountFloat(fSroValue, fSro))
	//
	//// Bpe
	//fNbBpe := func(bpe *Bpe) int { return 1 }
	//fNbSplice := func(bpe *Bpe) int { return bpe.NbSplice }
	//fValue := func(bpe *Bpe) float64 { return bpe.BpeValue + bpe.SpliceValue }
	//for _, priceCat := range priceCatalog.Chapters {
	//	for _, p := range priceCat {
	//		row++
	//		fTDBpe := func(bpe *Bpe) bool { return bpe.ToDo && bpe.PriceName == p.Name }
	//		fBpe := func(bpe *Bpe) bool { return bpe.ToDo && bpe.Done && bpe.PriceName == p.Name }
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 0), p.Name)
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 1), p.Price)
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 2), s.CountInt(fNbBpe, fTDBpe))
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 3), s.CountInt(fNbBpe, fBpe))
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 4), p.GetSpliceValue(1))
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 5), s.CountInt(fNbSplice, fTDBpe))
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 6), s.CountInt(fNbSplice, fBpe))
	//		xf.SetCellValue(attachmentSheetName, xls.RcToAxis(row, 7), s.CountFloat(fValue, fBpe))
	//	}
	//}

}

func (s *Suivi) writeProgressSheet(xf *excelize.File) {
	if xf.GetSheetIndex(progressSheetName) == 0 {
		xf.NewSheet(progressSheetName)
	}

	row := 0
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 0), "Item")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 1), "Info")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 2), "Code BPU")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 3), "Quantité")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 4), "Prix")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 5), "Installé")
	xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 6), "Semaine")
	for _, item := range s.Items {
		if !(item.Todo && item.Quantity > 0) {
			continue
		}
		row++
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 0), item.Name)
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 1), item.Info)
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 2), item.Article.Name)
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 3), item.Quantity)
		xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 4), item.Price())
		if item.Done {
			xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 5), "Oui")
			xf.SetCellValue(progressSheetName, xls.RcToAxis(row, 6), item.Date)
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

func (s *Suivi) GetItems(keep func(item *bpu.Item) bool) map[string][]*bpu.Item {
	res := make(map[string][]*bpu.Item)
	for _, item := range s.Items {
		if keep(item) {
			res[item.Article.Name] = append(res[item.Article.Name], item)
		}
	}
	return res
}

func (s *Suivi) CountInt(items []*bpu.Item, val func(item *bpu.Item) int, keep func(item *bpu.Item) bool) int {
	res := 0
	for _, item := range items {
		if keep(item) {
			res += val(item)
		}
	}
	return res
}

func (s *Suivi) CountFloat(items []*bpu.Item, val func(item *bpu.Item) float64, keep func(item *bpu.Item) bool) float64 {
	res := 0.0
	for _, item := range items {
		if keep(item) {
			res += val(item)
		}
	}
	return res
}

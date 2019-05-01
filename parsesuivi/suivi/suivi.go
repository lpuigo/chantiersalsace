package suivi

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"strings"
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
	catActivity  string = "Activité "
	sheetTirage  string = "Tirage"
	sheetRacco   string = "Racco"
	sheetMeasure string = "Mesures"
)

func NewSuiviFromXLS(file string, catalog *bpu.Catalog) (s *Suivi, perr ParsingError) {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		perr.Add(err, true)
		return
	}

	s = NewSuivi(catalog)

	tabFound := false
	// TODO parse Tirage Tab sheetTirage
	rsh := xf.Sheet[sheetTirage]
	if rsh == nil {
		perr.Add(fmt.Errorf("onglet '%s' non traité", sheetTirage), false)
	} else {
		tabFound = true
		pullErr := ParseTab(rsh, NewPullingParser(), s)
		perr.Append(pullErr)
	}

	// parse Racco Tab sheetRacco
	rsh = xf.Sheet[sheetRacco]
	if rsh == nil {
		perr.Add(fmt.Errorf("onglet '%s' non traité", sheetRacco), false)
	} else {
		tabFound = true
		raccoErr := ParseTab(rsh, NewRaccoParser(), s)
		perr.Append(raccoErr)
	}

	// parse Mesure Tab sheetMeasure
	msh := xf.Sheet[sheetMeasure]
	if msh == nil {
		perr.Add(fmt.Errorf("onglet '%s' non traité", sheetMeasure), false)
	} else {
		tabFound = true
		measErr := ParseTab(msh, NewMeasurementParser(), s)
		perr.Append(measErr)
	}

	if !tabFound {
		perr.Add(fmt.Errorf("aucun onglet traité"), true)
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

func (s *Suivi) writeSuiviSheet(xf *excelize.File) {
	if xf.GetSheetIndex(suiviSheetName) == 0 {
		xf.NewSheet(suiviSheetName)
	}

	fAllBpu := func(item *bpu.Item) bool { return !strings.HasPrefix(item.Article.Name, catActivity) }
	fAct := func(item *bpu.Item) bool { return strings.HasPrefix(item.Article.Name, catActivity) }
	fTodo := func(item *bpu.Item) bool { return item.Todo && !strings.HasPrefix(item.Article.Name, catActivity) }
	fNbQty := func(item *bpu.Item) int { return item.Quantity }
	fPrice := func(item *bpu.Item) float64 { return item.Price() }

	articleNames := s.Catalog.GetArticleNames(sheetTirage)
	articleNames = append(articleNames, s.Catalog.GetArticleNames(sheetRacco)...)
	articleNames = append(articleNames, s.Catalog.GetArticleNames(sheetMeasure)...)

	nbTot := make(map[string]int)
	valTot := make(map[string]float64)
	actTot := make(map[string]float64)
	valTot[global] = s.CountFloat(s.Items, fPrice, fTodo)
	articleNameItems := s.GetItems(fTodo)
	for _, articleName := range articleNames {
		nbTot[articleName] = s.CountInt(articleNameItems[articleName], fNbQty, fAllBpu)
		valTot[articleName] = s.CountFloat(articleNameItems[articleName], fPrice, fAllBpu)
		actTot[articleName] = s.CountFloat(articleNameItems[articleName], fPrice, fAct)
	}

	// Set Dates header
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(0, 1), "Semaines")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, 0), "Suivi financier")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, 1), "€ Total")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(2, 1), "€ Fait")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(3, 1), "€ /Sem")
	xf.SetCellValue(suiviSheetName, xls.RcToAxis(4, 1), "%")
	row := 4
	for _, articleName := range articleNames {
		if !(nbTot[articleName] != 0 && !strings.HasPrefix(articleName, catActivity)) {
			continue
		}
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+1, 0), articleName)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+1, 1), "Nb Total")
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+2, 1), "Nb")
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+3, 1), "%")
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+4, 1), "€ Total")
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+5, 1), "€ Fait")
		row += 5
	}
	prevWeekProd := 0.0
	for col, d := range s.Dates() {
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(0, col+2), d)
		fDone := func(item *bpu.Item) bool {
			return item.Done && !item.Date.After(d) && !strings.HasPrefix(item.Article.Name, catActivity)
		}
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(1, col+2), valTot[global])
		weekProd := s.CountFloat(s.Items, fPrice, fDone)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(2, col+2), weekProd)
		xf.SetCellValue(suiviSheetName, xls.RcToAxis(3, col+2), weekProd-prevWeekProd)
		if valTot[global] > 0 {
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(4, col+2), weekProd/valTot[global])
		} else {
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(4, col+2), 0)
		}
		prevWeekProd = weekProd
		row := 4
		for _, articleName := range articleNames {
			if !(nbTot[articleName] != 0 && !strings.HasPrefix(articleName, catActivity)) {
				continue
			}
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+1, col+2), nbTot[articleName])
			nb := s.CountInt(articleNameItems[articleName], fNbQty, fDone)
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+2, col+2), nb)
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+3, col+2), float64(nb)/float64(nbTot[articleName]))
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+4, col+2), valTot[articleName])
			xf.SetCellValue(suiviSheetName, xls.RcToAxis(row+5, col+2), s.CountFloat(articleNameItems[articleName], fPrice, fDone))
			row += 5
		}
	}
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

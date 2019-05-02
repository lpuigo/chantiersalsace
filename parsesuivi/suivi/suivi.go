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
	workSheetName       string = "Travail"
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
	s.writeWorkSheet(xf)
	s.writeProgressSheet(xf)
	return xf.Save()
}

func (s *Suivi) writeSuiviSheet(xf *excelize.File) {
	if xf.GetSheetIndex(suiviSheetName) == 0 {
		xf.NewSheet(suiviSheetName)
	}

	fArticleName := func(item *bpu.Item) string { return item.Article.Name }
	fAllBpu := func(item *bpu.Item) bool { return true }
	fTodo := func(item *bpu.Item) bool { return item.Todo }
	fNbQty := func(item *bpu.Item) int { return item.Quantity }
	fPrice := func(item *bpu.Item) float64 { return item.Price() }

	articleNames := s.Catalog.GetArticleNames(sheetTirage)
	articleNames = append(articleNames, s.Catalog.GetArticleNames(sheetRacco)...)
	articleNames = append(articleNames, s.Catalog.GetArticleNames(sheetMeasure)...)

	nbTot := make(map[string]int)
	valTot := make(map[string]float64)
	valTot[global] = s.CountFloat(s.Items, fPrice, fTodo)
	articleNameItems := s.GetItems(fArticleName, fTodo)
	for _, articleName := range articleNames {
		nbTot[articleName] = s.CountInt(articleNameItems[articleName], fNbQty, fAllBpu)
		valTot[articleName] = s.CountFloat(articleNameItems[articleName], fPrice, fAllBpu)
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
		if nbTot[articleName] == 0 {
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
			return item.Done && !item.Date.After(d)
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
			if nbTot[articleName] == 0 {
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

func (s *Suivi) writeWorkSheet(xf *excelize.File) {
	if xf.GetSheetIndex(workSheetName) == 0 {
		xf.NewSheet(workSheetName)
	}

	fArticleName := func(item *bpu.Item) string { return item.Article.Name }
	fActivity := func(item *bpu.Item) string { return item.Activity }
	fAllItems := func(item *bpu.Item) bool { return true }
	fTodoItems := func(item *bpu.Item) bool { return item.Todo }
	fWorkQty := func(item *bpu.Item) int { return item.WorkQuantity }
	fWork := func(item *bpu.Item) float64 { return item.Work() }

	activities := []string{sheetTirage, sheetRacco, sheetMeasure}
	articleNames := []string{}
	for _, activity := range activities {
		articleNames = append(articleNames, s.Catalog.GetArticleNames(activity)...)
	}

	nbAct := make(map[string]int)
	valAct := make(map[string]float64)
	activityItems := s.GetItems(fActivity, fTodoItems)
	for _, activity := range activities {
		nbAct[activity] = s.CountInt(activityItems[activity], fWorkQty, fAllItems)
		valAct[activity] = s.CountFloat(activityItems[activity], fWork, fAllItems)
	}

	nbTot := make(map[string]int)
	valTot := make(map[string]float64)
	valTot[global] = s.CountFloat(s.Items, fWork, fTodoItems)
	articleNameItems := s.GetItems(fArticleName, fTodoItems)
	for _, articleName := range articleNames {
		nbTot[articleName] = s.CountInt(articleNameItems[articleName], fWorkQty, fAllItems)
		valTot[articleName] = s.CountFloat(articleNameItems[articleName], fWork, fAllItems)
	}

	// Set Info cols
	// Global
	xf.SetCellValue(workSheetName, xls.RcToAxis(0, 1), "Semaines")
	xf.SetCellValue(workSheetName, xls.RcToAxis(1, 0), "Travail")
	xf.SetCellValue(workSheetName, xls.RcToAxis(1, 1), "Total")
	xf.SetCellValue(workSheetName, xls.RcToAxis(2, 1), "Fait")
	xf.SetCellValue(workSheetName, xls.RcToAxis(3, 1), "%")
	xf.SetCellValue(workSheetName, xls.RcToAxis(4, 1), "/Sem")
	xf.SetCellValue(workSheetName, xls.RcToAxis(5, 1), "Effectif")
	xf.SetCellValue(workSheetName, xls.RcToAxis(6, 1), "% Eff.")
	// per Activities
	row := 6
	for _, activity := range activities {
		if nbAct[activity] == 0 {
			row += 5
			continue
		}
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+1, 0), activity)
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+1, 1), "Total")
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+2, 1), "%")
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+3, 1), "/Sem")
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+4, 1), "Effectif")
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+5, 1), "% Eff.")
		row += 5
	}

	// per Article Names
	for _, articleName := range articleNames {
		if nbTot[articleName] == 0 {
			continue
		}
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+1, 0), articleName)
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+1, 1), "Total")
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+2, 1), "%")
		xf.SetCellValue(workSheetName, xls.RcToAxis(row+3, 1), "/Sem")
		row += 3
	}
	prevWeekProd := 0.0
	prevWeekProdFor := make(map[string]float64)
	for col, d := range s.Dates() {
		xf.SetCellValue(workSheetName, xls.RcToAxis(0, col+2), d)
		fDoneItems := func(item *bpu.Item) bool {
			return item.Done && !item.Date.After(d)
		}
		xf.SetCellValue(workSheetName, xls.RcToAxis(1, col+2), valTot[global])
		weekProd := s.CountFloat(s.Items, fWork, fDoneItems)
		xf.SetCellValue(workSheetName, xls.RcToAxis(2, col+2), weekProd)
		if valTot[global] > 0 {
			xf.SetCellValue(workSheetName, xls.RcToAxis(3, col+2), weekProd/valTot[global])
		} else {
			xf.SetCellValue(workSheetName, xls.RcToAxis(3, col+2), 0)
		}
		xf.SetCellValue(workSheetName, xls.RcToAxis(4, col+2), weekProd-prevWeekProd)
		prevWeekProd = weekProd
		row := 6
		for _, activity := range activities {
			if nbAct[activity] == 0 {
				row += 5
				continue
			}
			xf.SetCellValue(workSheetName, xls.RcToAxis(row+1, col+2), valAct[activity])
			nb := s.CountFloat(activityItems[activity], fWork, fDoneItems)
			if valAct[activity] > 0 {
				xf.SetCellValue(workSheetName, xls.RcToAxis(row+2, col+2), nb/valAct[activity])
			} else {
				xf.SetCellValue(workSheetName, xls.RcToAxis(row+2, col+2), 0.0)
			}
			xf.SetCellValue(workSheetName, xls.RcToAxis(row+3, col+2), nb-prevWeekProdFor[activity])
			prevWeekProdFor[activity] = nb
			row += 5
		}
		for _, articleName := range articleNames {
			if nbTot[articleName] == 0 {
				continue
			}
			xf.SetCellValue(workSheetName, xls.RcToAxis(row+1, col+2), valTot[articleName])
			nb := s.CountFloat(articleNameItems[articleName], fWork, fDoneItems)
			xf.SetCellValue(workSheetName, xls.RcToAxis(row+2, col+2), nb)
			xf.SetCellValue(workSheetName, xls.RcToAxis(row+3, col+2), nb-prevWeekProdFor[articleName])
			prevWeekProdFor[articleName] = nb
			row += 3
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

func (s *Suivi) GetItems(info func(item *bpu.Item) string, keep func(item *bpu.Item) bool) map[string][]*bpu.Item {
	res := make(map[string][]*bpu.Item)
	for _, item := range s.Items {
		if keep(item) {
			res[info(item)] = append(res[info(item)], item)
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

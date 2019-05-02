package bpu

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"sort"
	"strings"
)

type Catalog struct {
	Chapters map[string]CategoryChapters // map[Activity]map[Category][]*Article
	Boxes    map[string]map[string]*Box  // map[Category]map[Name]*Box
}

const (
	bpuPriceSheetName = "Prices"
	bpuBoxeSheetName  = "Boxes"
)

func NewCatalogFromXLS(file string) (bpu *Catalog, err error) {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		return
	}
	priceSheet := xf.Sheet[bpuPriceSheetName]
	if priceSheet == nil {
		err = fmt.Errorf("could not find '%s' sheet in '%s'", bpuPriceSheetName, file)
		return
	}
	boxSheet := xf.Sheet[bpuBoxeSheetName]
	if boxSheet == nil {
		err = fmt.Errorf("could not find '%s' sheet in '%s'", bpuBoxeSheetName, file)
		return
	}
	bpu = &Catalog{
		Chapters: map[string]CategoryChapters{},
		Boxes:    map[string]map[string]*Box{},
	}
	err = bpu.parseChapters(priceSheet)
	if err != nil {
		return
	}
	err = bpu.parseBoxes(boxSheet)
	return
}

func (catalog *Catalog) GetCategoryChapters(activity string) CategoryChapters {
	return catalog.Chapters[strings.ToUpper(activity)]
}

const (
	colPricesActivity int = iota
	colPricesCategory
	colPricesName
	colPricesSize
	colPricesPrice
	colPricesWork
	colPricesEnd
)

func (catalog *Catalog) parseChapters(sheet *xlsx.Sheet) (err error) {
	entryFound := true
	for row := 1; entryFound; row++ {
		activity := sheet.Cell(row, colPricesActivity).Value
		if activity == "" { // check for data ending (first column is empty => we are done)
			entryFound = false
			continue
		}
		cat := sheet.Cell(row, colPricesCategory).Value
		nChapter, err := parseChapterRow(sheet, row)
		if err != nil {
			return err
		}
		cat = strings.ToUpper(cat)
		activity = strings.ToUpper(activity)
		currentActivityCatChapters, found := catalog.Chapters[activity]
		if !found {
			currentActivityCatChapters = NewCategoryChapters()
			catalog.Chapters[activity] = currentActivityCatChapters
		}
		currentActivityCatChapters[cat] = append(currentActivityCatChapters[cat], nChapter)
	}

	// sort price categories by ascending size
	//for _, actCatChapt := range catalog.Chapters {
	//	actCatChapt.SortChapters()
	//}
	return
}

func parseChapterRow(sh *xlsx.Sheet, row int) (p *Article, err error) {
	cSize := sh.Cell(row, colPricesSize)
	size, e := cSize.Int()
	if e != nil {
		err = fmt.Errorf("could not get size info '%s' in sheet '%s!%s'", cSize.Value, bpuPriceSheetName, xls.RcToAxis(row, colPricesSize))
		return
	}
	cPrice := sh.Cell(row, colPricesPrice)
	price, e := cPrice.Float()
	if e != nil {
		err = fmt.Errorf("could not get price info '%s' in sheet '%s!%s'", cSize.Value, bpuPriceSheetName, xls.RcToAxis(row, colPricesPrice))
		return
	}
	cWork := sh.Cell(row, colPricesWork)
	work, e := cWork.Float()
	if e != nil {
		err = fmt.Errorf("could not get work info '%s' in sheet '%s!%s'", cSize.Value, bpuPriceSheetName, xls.RcToAxis(row, colPricesWork))
		return
	}
	p = NewChapter()
	p.Name = sh.Cell(row, colPricesName).Value
	p.Size = size
	p.Price = price
	p.Work = work
	return
}

const (
	colBoxesCategory int = iota
	colBoxesName
	colBoxesSize
	colBoxesUsage
	colBoxesEnd
)

func (catalog *Catalog) parseBoxes(sheet *xlsx.Sheet) (err error) {
	entryFound := true
	for row := 1; entryFound; row++ {
		cat := strings.ToUpper(sheet.Cell(row, colBoxesCategory).Value)
		boxName := strings.ToUpper(sheet.Cell(row, colBoxesName).Value)
		if boxName == "" { // check for data ending (first column is empty => we are done)
			entryFound = false
			continue
		}
		sizeCell := sheet.Cell(row, colBoxesSize)
		size, e := sizeCell.Int()
		if e != nil {
			err = fmt.Errorf("could not get size info '%s' in sheet '%s!%s'", sizeCell.Value, bpuBoxeSheetName, xls.RcToAxis(row, colBoxesSize))
			return
		}
		nBox := NewBox()
		nBox.Name = boxName
		nBox.Size = size
		nBox.Usage = sheet.Cell(row, colBoxesUsage).Value
		catBox, found := catalog.Boxes[cat]
		if !found {
			catBox = make(map[string]*Box)
			catalog.Boxes[cat] = catBox
		}
		catBox[nBox.Name] = nBox
	}
	return
}

//func (bpu *Catalog) String() string {
//	res := ""
//	for _, p := range bpu.BpePrices {
//		res += fmt.Sprintf("%3d : %v\n", p.Size, p.Article)
//	}
//	return res
//}

func (catalog *Catalog) IsBoxDefined(cat, boxName string) bool {
	catBox, found := catalog.Boxes[strings.ToUpper(cat)]
	if !found {
		return false
	}
	_, found = catBox[strings.ToUpper(boxName)]
	return found
}

func (catalog *Catalog) GetBox(cat, boxName string) *Box {
	catBox, found := catalog.Boxes[strings.ToUpper(cat)]
	if !found {
		return nil
	}
	return catBox[strings.ToUpper(boxName)]
}

func (catalog *Catalog) GetArticleNames(activity string) []string {
	res := []string{}
	for _, catChapters := range catalog.GetCategoryChapters(activity) {
		for _, chapter := range catChapters {
			res = append(res, chapter.Name)
		}
	}
	sort.Strings(res)
	return res
}

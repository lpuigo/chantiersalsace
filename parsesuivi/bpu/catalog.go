package bpu

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"sort"
	"strings"
)

type Catalog struct {
	Chapters map[string][]*Chapter
	Boxes    map[string]map[string]*Box
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
		Chapters: map[string][]*Chapter{},
		Boxes:    map[string]map[string]*Box{},
	}
	err = bpu.parseChapters(priceSheet)
	if err != nil {
		return
	}
	err = bpu.parseBoxes(boxSheet)
	return
}

const (
	colPricesCategory int = iota
	colPricesName
	colPricesSize
	colPricesPrice
	colPricesEnd
)

func (catalog *Catalog) parseChapters(sheet *xlsx.Sheet) (err error) {
	entryFound := true
	for row := 1; entryFound; row++ {
		if sheet.Cell(row, colPricesCategory).Value == "" { // check for data ending (first column is empty => we are done)
			entryFound = false
			continue
		}
		cat := sheet.Cell(row, colPricesCategory).Value
		nPrice, err := parseChapterRow(sheet, row)
		if err != nil {
			return err
		}
		cat = strings.ToUpper(cat)
		catalog.Chapters[cat] = append(catalog.Chapters[cat], nPrice)
	}

	// sort price categories by ascending size
	for cat, prices := range catalog.Chapters {
		sort.Slice(prices, func(i, j int) bool {
			return prices[i].Size < prices[j].Size
		})
		catalog.Chapters[cat] = prices
	}
	return
}

func parseChapterRow(sh *xlsx.Sheet, row int) (p *Chapter, err error) {
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
	p = NewChapter()
	p.Name = sh.Cell(row, colPricesName).Value
	p.Size = size
	p.Price = price
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
//		res += fmt.Sprintf("%3d : %v\n", p.Size, p.Chapter)
//	}
//	return res
//}

// GetRaccoBoxPrice returns Chapter applicable for given Bpe type (bpeType must be declared)
func (catalog *Catalog) GetRaccoBoxPrice(cat, name string) (box, splice *Chapter, err error) {
	// box lookup
	cat = strings.ToUpper(cat)
	name = strings.ToUpper(name)
	catBox, found := catalog.Boxes[cat]
	if !found {
		err = fmt.Errorf("unknow box category '%s'", cat)
		return
	}
	b, found := catBox[name]
	if !found {
		err = fmt.Errorf("unknow box name '%s' for category '%s'", name, cat)
		return
	}

	// box price lookup
	catPrice, found := catalog.Chapters[cat]
	if !found {
		err = fmt.Errorf("unknow price category '%s'", cat)
		return
	}

	for _, p := range catPrice {
		if b.Size <= p.Size {
			box = p
			break
		}
	}
	if box == nil {
		err = fmt.Errorf("no minimal price size found for '%d'", b.Size)
		return
	}

	// splice price lookup
	spliceCat := cat + " SPLICE"
	catPrice, found = catalog.Chapters[spliceCat]
	if !found {
		// no Splice for this category, leave splice price as nil
		return
	}

	for _, p := range catPrice {
		if b.Size <= p.Size {
			splice = p
			break
		}
	}
	if splice == nil {
		err = fmt.Errorf("no minimal splice price size found for '%d'", b.Size)
		return
	}

	return
}

// GetRaccoPmPrices returns applicable PM Chapter & missing ModulePrice
func (catalog *Catalog) GetRaccoPmPrices() (*Chapter, *Chapter) {
	return catalog.Chapters["PM"][0], catalog.Chapters["PM"][1]
}

func (catalog *Catalog) IsBoxDefined(cat, name string) bool {
	catBox, found := catalog.Boxes[strings.ToUpper(cat)]
	if !found {
		return false
	}
	_, found = catBox[strings.ToUpper(name)]
	return found
}
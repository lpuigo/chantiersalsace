package bpu

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"sort"
)

type Bpu struct {
	Prices map[string][]*Price
	Boxes  map[string]int
}

const (
	colPricesCategory int = iota
	colPricesName
	colPricesSize
	colPricesPrice
	colPricesEnd
)
const (
	colBoxesCategory int = iota
	colBoxesName
	colBoxesSize
	colBoxesType
	colBoxesEnd
)

const (
	bpuPriceSheetName = "Prices"
	bpuBoxeSheetName  = "Boxes"
)

func NewBpuFromXLS(file string) (bpu *Bpu, err error) {
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
	bpu = &Bpu{
		Prices: map[string][]*Price{},
		Boxes:  map[string]int{},
	}
	err = bpu.parseBpePrices(priceSheet)
	if err != nil {
		return
	}
	err = bpu.parseBoxes(boxSheet)
	return
}

func (bpu *Bpu) parseBpePrices(sheet *xlsx.Sheet) (err error) {
	header := make([]string, colPricesEnd)
	for col := colPricesCategory; col < colPricesEnd; col++ {
		header[col] = sheet.Cell(0, col).Value
	}

	entryFound := true
	for row := 1; entryFound; row++ {
		if sheet.Cell(row, colPricesCategory).Value == "" { // check for data ending (first column is empty => we are done)
			entryFound = false
			continue
		}
		cat := sheet.Cell(row, colPricesCategory).Value
		nPrice, err := parsePriceRow(sheet, row, header)
		if err != nil {
			return
		}
		bpu.Prices[cat] = append(bpu.Prices[cat], nPrice)
	}

	// sort price categories by ascending size
	for cat, prices := range bpu.Prices {
		sort.Slice(prices, func(i, j int) bool {
			return prices[i].Size < prices[j].Size
		})
		bpu.Prices[cat] = prices
	}
	return
}

func parsePriceRow(sh *xlsx.Sheet, row int, header []string) (p *Price, err error) {
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
	p = NewPrice()
	p.Name = sh.Cell(row, colPricesName).Value
	p.Size = size
	p.Price = price
	return
}

func (bpu *Bpu) parseBoxes(sheet *xlsx.Sheet) (err error) {
	bpu.Boxes = make(map[string]int)
	header := make([]string, colBoxesEnd)
	for col := colBoxesName; col < colBoxesEnd; col++ {
		header[col] = sheet.Cell(0, col).Value
	}

	entryFound := true
	for row := 1; entryFound; row++ {
		boxName := sheet.Cell(row, colBoxesName).Value
		if boxName == "" { // check for data ending (first column is empty => we are done)
			entryFound = false
			continue
		}
		sizeCell := sheet.Cell(row, colBoxesSize)
		size, e := sizeCell.Int()
		if e != nil {
			err = fmt.Errorf("could not get size info '%s' in sheet '%s(%s)'", sizeCell.Value, bpuBoxeSheetName, xls.RcToAxis(row, 0))
			return
		}
		bpu.Boxes[boxName] = size
	}
	return
}

//func (bpu *Bpu) String() string {
//	res := ""
//	for _, p := range bpu.BpePrices {
//		res += fmt.Sprintf("%3d : %v\n", p.Size, p.Price)
//	}
//	return res
//}

// GetBpePrice returns Price applicable for given Bpe type (bpeType must be declared)
func (bpu *Bpu) GetBpePrice(bpeType string) *Price {
	size := bpu.Boxes[bpeType]
	for _, p := range bpu.BpePrices {
		if size <= p.Size {
			return p
		}
	}
	return bpu.BpePrices[len(bpu.BpePrices)-1]
}

// GetSroPrice returns applicable Sro Price & missing ModulePrice
func (bpu *Bpu) GetSroPrice() (*Price, *Price) {
	return bpu.SroPrices[0], bpu.SroPrices[1]
}

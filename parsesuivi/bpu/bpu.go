package bpu

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"sort"
)

type Bpu struct {
	BpePrices []*Price
	SroPrices []*Price
	Boxes     map[string]int
}

const (
	colPricesType int = iota
	colPricesName
	colPricesSize
	colPricesBox
	colPricesSplice
	colPricesEnd
)
const (
	colBoxesName int = iota
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
	bpu = &Bpu{}
	err = bpu.parseBpePrices(priceSheet)
	if err != nil {
		return
	}
	err = bpu.parseBoxes(boxSheet)
	return
}

func (bpu *Bpu) parseBpePrices(sheet *xlsx.Sheet) (err error) {
	header := make([]string, colPricesEnd)
	for col := colPricesType; col < colPricesEnd; col++ {
		header[col] = sheet.Cell(0, col).Value
	}

	entryFound := true
	var entry *Price
	for row := 1; entryFound; row++ {
		if sheet.Cell(row, colPricesType).Value == "" { // check for data ending (first column is empty => we are done)
			entryFound = false
			continue
		}
		switch sheet.Cell(row, colPricesType).Value {
		case "bpe":
			entry, err = parseBpePriceRow(sheet, row, header)
			if err != nil {
				return
			}
			bpu.BpePrices = append(bpu.BpePrices, entry)
		case "sro":
			entry, err = parseBpePriceRow(sheet, row, header)
			if err != nil {
				return
			}
			bpu.SroPrices = append(bpu.SroPrices, entry)
		}
	}

	sort.Slice(bpu.BpePrices, func(i, j int) bool {
		return bpu.BpePrices[i].Size < bpu.BpePrices[j].Size
	})
	return
}

func parseBpePriceRow(sh *xlsx.Sheet, row int, header []string) (p *Price, err error) {
	cSize := sh.Cell(row, colPricesSize)
	size, e := cSize.Int()
	if e != nil {
		err = fmt.Errorf("could not get size info '%s' in sheet '%s(%s)'", cSize.Value, bpuPriceSheetName, xls.RcToAxis(row, 0))
		return
	}
	p = NewPrice(int(size))
	p.Name = sh.Cell(row, colPricesName).Value
	for col := colPricesBox; col < colPricesEnd; col++ {
		cell := sh.Cell(row, col)
		fval, e := cell.Float()
		if e != nil {
			err = fmt.Errorf("could not get value '%s' in sheet '%s(%s)'", cell.Value, bpuPriceSheetName, xls.RcToAxis(row, col))
			return
		}
		p.Income[header[col]] = fval
	}
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

func (bpu *Bpu) String() string {
	res := ""
	for _, p := range bpu.BpePrices {
		res += fmt.Sprintf("%3d : %v\n", p.Size, p.Income)
	}
	return res
}

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

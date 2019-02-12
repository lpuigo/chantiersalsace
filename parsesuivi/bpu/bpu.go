package bpu

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"sort"
)

type Bpu struct {
	Prices []*Price
}

const (
	bpuPriceSheetName = "Prices"
)

func NewBpuFromXLS(file string) (bpu *Bpu, err error) {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		return
	}
	psh := xf.Sheet[bpuPriceSheetName]
	if psh == nil {
		err = fmt.Errorf("could not find '%s' sheet in '%s'", bpuPriceSheetName, file)
		return
	}

	header := make([]string, 3)
	for col := 0; col < 3; col++ {
		header[col] = psh.Cell(0, col).Value
	}

	bpu = &Bpu{}

	entryFound := true
	for row := 1; entryFound; row++ {
		ssize := psh.Cell(row, 0).Value
		if ssize == "" {
			entryFound = false
			continue
		}
		size, e := psh.Cell(row, 0).Int()
		if e != nil {
			err = fmt.Errorf("could not get size info '%s' in sheet '%s(%s)'", ssize, bpuPriceSheetName, xls.RcToAxis(row, 0))
			return
		}
		entry := NewPrice(int(size))
		for col := 1; col < 3; col++ {
			fval, e := psh.Cell(row, col).Float()
			if e != nil {
				err = fmt.Errorf("could not get value '%s' in sheet '%s(%s)'", psh.Cell(row, col).Value, bpuPriceSheetName, xls.RcToAxis(row, col))
				return
			}
			entry.Income[header[col]] = fval
		}
		bpu.Prices = append(bpu.Prices, entry)
	}

	sort.Slice(bpu.Prices, func(i, j int) bool {
		return bpu.Prices[i].Size < bpu.Prices[j].Size
	})

	return
}

func (bpu *Bpu) String() string {
	res := ""
	for _, p := range bpu.Prices {
		res += fmt.Sprintf("%3d : %v\n", p.Size, p.Income)
	}
	return res
}

// GetPrice returns Price applicable for given size ( prev Price.size > size >= target price.size)
func (bpu *Bpu) GetPrice(size int) *Price {
	for _, p := range bpu.Prices {
		if size <= p.Size {
			return p
		}
	}
	return bpu.Prices[len(bpu.Prices)-1]
}

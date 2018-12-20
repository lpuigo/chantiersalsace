package parsezacable

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"testing"
)

const (
	textfile      = `C:\Users\Laurent\Golang\src\github.com\lpuig\ewin\chantiersalsace\parsezacable\test\ZACABLE_1.xlsx`
	testsheetname = "PBO-68-048-DXA-1010"
)

func openXLSFile(t *testing.T, file string) *xlsx.File {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		t.Fatalf("could not process xlsx file: %v", err)
	}
	return xf
}

func TestZone_ParseXLSSheet(t *testing.T) {
	xf := openXLSFile(t, textfile)

	sheetname := testsheetname
	xs, found := xf.Sheet[sheetname]
	if !found {
		t.Fatalf("could not find '%s' sheet", sheetname)
	}

	zone := NewZone(xs.Name)
	err := zone.ParseXLSSheet(xs)
	if err != nil {
		t.Fatalf("parse sheet returns: %v", err)
	}

	fmt.Printf("Zone : '%s' (short: '%s')\n", zone.FullName, zone.Name)
	for _, sfn := range zone.GetSiteFullNames() {
		s := zone.GetSiteByFullName(sfn)
		fmt.Println(s.String())
	}

}

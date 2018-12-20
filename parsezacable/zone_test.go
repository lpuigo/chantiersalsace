package parsezacable

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strings"
	"testing"
)

const (
	testfile      = `C:\Users\Laurent\Golang\src\github.com\lpuig\ewin\chantiersalsace\parsezacable\test\ZACABLE_1.xlsx`
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
	xf := openXLSFile(t, testfile)

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

	fmt.Printf("Zone : '%s'\n", zone.Name)
	for _, sfn := range zone.GetSiteFullNames() {
		s := zone.GetSiteByFullName(sfn)
		fmt.Println(s.String())
	}
}

func TestZone_ParseXLSFile(t *testing.T) {
	zone := NewZone(testsheetname)
	err := zone.ParseXLSFile(strings.TrimPrefix(testfile, ".xlsx"))
	if err != nil {
		t.Fatalf("parse XLS File returns: %v", err)
	}

	fmt.Printf("Zone : '%s'\n", zone.Name)
	for _, sfn := range zone.GetSiteFullNames() {
		s := zone.GetSiteByFullName(sfn)
		fmt.Println(s.String())
	}
}

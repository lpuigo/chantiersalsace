package parsezacable

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"testing"
)

const (
	testfile           = `C:\Users\Laurent\Golang\src\github.com\lpuig\ewin\chantiersalsace\parsezacable\test\ZACABLE_1.xlsx`
	testresultfile     = `C:\Users\Laurent\Golang\src\github.com\lpuig\ewin\chantiersalsace\parsezacable\test\DXA_Suivi.xlsx`
	testsheetname      = "PBO-68-048-DXA-1010"
	testblobpattern    = `C:\Users\Laurent\Desktop\DXA\ZACABLE*\ZACABLE*.xlsx`
	testblobresultfile = `C:\Users\Laurent\Desktop\DXA\DXA_Suivi.xlsx`
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
	zone := NewZone(GetShortZoneName(testsheetname))
	err := zone.ParseXLSFile(testfile)
	if err != nil {
		t.Fatalf("parse XLS File returns: %v", err)
	}

	fmt.Printf("Zone : '%s'\n", zone.Name)
	for _, sfn := range zone.GetSiteFullNames() {
		s := zone.GetSiteByFullName(sfn)
		fmt.Println(s.String())
	}
}

func TestZone_WriteXLS(t *testing.T) {
	zone := NewZone(GetShortZoneName(testsheetname))
	err := zone.ParseXLSFile(testfile)
	if err != nil {
		t.Fatalf("parse XLS File returns: %v", err)
	}

	err = zone.WriteXLS(testresultfile)
	if err != nil {
		t.Fatal("zone.WriteXLS returns:", err)
	}
}

func TestZone_ParseBlob(t *testing.T) {
	zone := NewZone(GetShortZoneName(testsheetname))
	err := zone.ParseBlob(testblobpattern)
	if err != nil {
		t.Fatalf("zone.ParseBlob returns: %v", err)
	}

	err = zone.WriteXLS(testblobresultfile)
	if err != nil {
		t.Fatal("zone.WriteXLS returns:", err)
	}
}

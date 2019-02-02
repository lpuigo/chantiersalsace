package xlsxconvert

import (
	"path/filepath"
	"testing"
)

const (
	testDir string = `C:\Users\Laurent\Desktop\CCPE_DES_PM3_BPE`
	testXLS string = `_PT-182061.xls`
)

func TestOleXlsToCsv(t *testing.T) {
	file := filepath.Join(testDir, testXLS)
	err := OleXlsToCsv(file)
	if len(err) > 0 && err[0] != nil {
		t.Fatal(err[0])
	}
}

func TestConvertDir(t *testing.T) {
	if err := ConvertDir(testDir); err != nil {
		t.Fatalf(err.Error())
	}
}

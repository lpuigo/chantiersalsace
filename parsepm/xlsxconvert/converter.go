package xlsxconvert

import (
	"fmt"
	"github.com/aswjh/excel"
	"path/filepath"
	"strings"
)

const (
	blobpattern string = `*.xls`
)

func OleXlsToCsv(inFile string) []error {
	option := excel.Option{"Visible": false, "DisplayAlerts": false}
	//xl, _ := excel.New(option)
	xl, _ := excel.Open(inFile, option)
	defer xl.Quit()

	fileExt := filepath.Ext(inFile)
	outFile := strings.TrimSuffix(inFile, fileExt) + ".xlsx"

	return xl.SaveAs(outFile, "xlsx") //xl.SaveAs("test_excel", "html")
}

func ConvertDir(dir string) error {
	parseBlobPattern := filepath.Join(dir, blobpattern)
	files, err := filepath.Glob(parseBlobPattern)
	if err != nil {
		return err
	}
	for _, f := range files {
		fmt.Printf("Converting '%s' ... ", f)
		errs := OleXlsToCsv(f)
		if len(errs) > 1 && errs[0] != nil {
			fmt.Printf("%s\n", errs[0].Error())
			continue
		}
		fmt.Printf("OK\n")
	}
	return nil
}

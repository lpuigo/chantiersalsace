package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/zone"
	"log"
)

const (
	testBPEDir  string = `C:\Users\Laurent\Desktop\CCPE_DES_PM3_BPE`
	testROPXlsx string = `C:\Users\Laurent\Desktop\CCPE_DES_PM3_ROP\CCPE_DES_PM3_ROP.xlsx`
	testXLS     string = `PM3`
)

func main() {
	pm := zone.New()
	err := pm.ParseBPEDir(testBPEDir)
	if err != nil {
		log.Fatal("could not parse :", err)
	}

	err = pm.ParseROPXLS(testROPXlsx)
	if err != nil {
		log.Fatal("could not parse ROP file:", err)
	}

	fmt.Print(pm.Sro.Tree("- ", "", 0))

	err = pm.WriteXLS(testBPEDir, testXLS)
	if err != nil {
		log.Fatal("could not write XLSx :", err)
	}
}

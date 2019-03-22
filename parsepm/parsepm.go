package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/zone"
	"log"
	"os"
)

const (
	testBPEDir  string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\MBeck - AUDUN-LE-TICHE\Info\BPE`
	testROPXlsx string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\MBeck - AUDUN-LE-TICHE\Info\CCPE_DES_PM3_ROP.xlsx`
	testXLS     string = `PM3`
)

func main() {
	pm := zone.New()
	err := pm.ParseBPEDir(testBPEDir)
	if err != nil {
		log.Fatal("could not parse BPE Directory:", err)
	}

	if exists(testROPXlsx) {
		// If ROP File exist, parse it to create BPE Tree
		err = pm.ParseROPXLS(testROPXlsx)
		if err != nil {
			log.Fatal("could not parse ROP file:", err)
		}

		fmt.Print(pm.Sro.Tree("- ", "", 0))

	} else {
		// Otherwise, scan Nodes list to create BPE Tree
		pm.CreateBPETree()
	}

	err = pm.WriteXLS(testBPEDir, testXLS)
	if err != nil {
		log.Fatal("could not write XLSx :", err)
	}
}

func exists(file string) bool {
	_, err := os.Stat(file)
	return !os.IsNotExist(err)
}

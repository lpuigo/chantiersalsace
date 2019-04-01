package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/zone"
	"log"
	"os"
)

const (
	testDir     string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\DMaussand - KEDANGE\Info\CCAM_KED_PM03\`
	testBPEDir  string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\DMaussand - KEDANGE\Info\CCAM_KED_PM03\CCAM_KED_PM03_BPE`
	testROPXlsx string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\DMaussand - KEDANGE\Info\CCAM_KED_PM03\CCAM_KED_PM03_ROP\CCAM_KED_PM03_ROP.xlsx`
	testXLS     string = `KED_PM03`
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

	err = pm.WriteXLS(testDir, testXLS)
	if err != nil {
		log.Fatal("could not write XLSx :", err)
	}
}

func exists(file string) bool {
	_, err := os.Stat(file)
	return !os.IsNotExist(err)
}

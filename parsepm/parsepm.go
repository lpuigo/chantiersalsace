package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/zone"
	"log"
	"os"
	"path/filepath"
)

const (
	//testDir     string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\JTestevuide - DESSELING PM3\Infos PM3`
	//testBPEDir  string = `CCPE_DES_PM3_BPE`
	//testROPXlsx string = `CCPE_DES_PM3_ROP\CCPE_DES_PM3_ROP.xlsx`
	//testXLS     string = `DES_PM3`
	//testDir     string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\DMaussand - KEDANGE\Info\CCAM_KED_PM03\`
	//testBPEDir  string = `CCAM_KED_PM03_BPE`
	//testROPXlsx string = `CCAM_KED_PM03_ROP\CCAM_KED_PM03_ROP.xlsx`
	//testXLS     string = `KED_PM03`
	testDir     string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Infos Chantiers\DMaussand - KEDANGE\Info\CCAM_KED_PM09\`
	testBPEDir  string = `CCAM_KED_PM09_BPE`
	testROPXlsx string = `CCAM_KED_PM09_ROP\CCAM_KED_PM09_ROP.xlsx`
	testXLS     string = `KED_PM09`
)

func main() {

	pm := zone.New()
	err := pm.ParseBPEDir(filepath.Join(testDir, testBPEDir))
	if err != nil {
		log.Fatal("could not parse BPE Directory:", err)
	}

	ropFile := filepath.Join(testDir, testROPXlsx)
	if exists(ropFile) {
		// If ROP File exist, parse it to create BPE Tree
		err = pm.ParseROPXLS(ropFile)
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

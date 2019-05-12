package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/zone"
	"log"
	"os"
	"path/filepath"
)

const (
	// DES_PM03
	//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\MBeck - GUE_PM03\Infos`
	//testBPEDir      string = `CCAM_GUE_PM3_BPE`
	//testROPXlsx     string = `CCAM_GUE_PM3_ROP.xlsx`
	//testCable94Xlsx string = ``
	//testXLS         string = `GUE_PM03`

	// DES_PM03
	testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\JTestevuide - DESSELING PM3\Infos PM3`
	testBPEDir      string = `CCPE_DES_PM3_BPE`
	testROPXlsx     string = `CCPE_DES_PM3_ROP\CCPE_DES_PM3_ROP.xlsx`
	testCable94Xlsx string = ``
	testXLS         string = `DES_PM03`

	// KED_PM03
	//testDir     string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\DMaussand - KEDANGE\Info\CCAM_KED_PM03\`
	//testBPEDir  string = `CCAM_KED_PM03_BPE`
	//testROPXlsx string = `CCAM_KED_PM03_ROP\CCAM_KED_PM03_ROP.xlsx`
	//testCable94Xlsx string = `Quantité_CCAM_KED_PM03\CCAM_KED_PM03_9_4.xlsx`
	//testXLS     string = `KED_PM03`

	// KED_PM08
	//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\DMaussand - KEDANGE\Info\CCAM-KED-PM08\CCAM_KED_PM08`
	//testBPEDir      string = `CCAM_KED_PM08_BPE`
	//testROPXlsx     string = `CCAM_KED_PM08_ROP\CCAM_KED_PM08_ROP.xlsx`
	//testCable94Xlsx string = `Quantité_CCAM_KED_PM08\CCAM_KED_PM08_9_4.xlsx`
	//testXLS         string = `KED_PM08`

	// KED_PM09
	//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\DMaussand - KEDANGE\Info\CCAM_KED_PM09\`
	//testBPEDir      string = `CCAM_KED_PM09_BPE`
	//testROPXlsx     string = `CCAM_KED_PM09_ROP\CCAM_KED_PM09_ROP.xlsx`
	//testCable94Xlsx string = `Quantité_CCAM_KED_PM09\CCAM_KED_PM09_9_4.xlsx`
	//testXLS         string = `KED_PM09`
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

	if testCable94Xlsx != "" {
		cable94File := filepath.Join(testDir, testCable94Xlsx)
		if !exists(cable94File) {
			log.Fatal("cable file '%s' does not exist\n", cable94File)
		}
		err = pm.ParseQuantiteCableXLS(cable94File)
		if err != nil {
			log.Fatal("could not parse Quantité Cable 9.4 file:", err)
		}
	}

	err = pm.WriteXLS(testDir, testXLS)
	if err != nil {
		log.Fatal("could not write XLSx :", err)
	}

	err = pm.WriteJSON(testDir, testXLS)
	if err != nil {
		log.Fatal("could not write JSON file :", err)
	}
}

func exists(file string) bool {
	_, err := os.Stat(file)
	return !os.IsNotExist(err)
}

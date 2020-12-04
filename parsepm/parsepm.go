package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/zone"
	"log"
	"os"
	"path/filepath"
)

const (
	// Sogetrel Fibre ==================================================================================================
	testClient  string = "Sogetrel Fibre"
	testManager string = "CHAUFFERT Nicolas"

	testXLS                string = `SRO 10_018_098`
	testDir                string = `C:\Users\Laurent\OneDrive\Documents\TEMPORAIRE\Sogetrel\Chantier Fibre Aube\2020-11-30 SRO 10_018_098\`
	testBPEDir             string = `PLANS DE SOUDURE 10_018_098`
	testROPXlsx            string = `20200914-SRO-10-018-098-ROP-EXCEL.xlsx`
	testCable94Xlsx        string = ``
	testCableOptiqueC2Xlsx string = ``

	// Axians Moselle ==================================================================================================
	//testClient  string = "Axians Moselle"
	//testManager string = "David MAUSSAND"
	////testManager string = "Matthieu BECK"
	//
	//testXLS                string = `SARLB_PM04`
	//testDir                string = `C:\Users\Laurent\OneDrive\Documents\TEMPORAIRE\Moselle\2020-11-25 SAR_PRO\CCAL_SAR_PM04\`
	//testBPEDir             string = `CCAL_SAR_PM04_BPE`
	//testROPXlsx            string = `CCAL_SAR_PM04_ROP\CCAL_SAR_PM04_ROP.xlsx`
	//testCable94Xlsx        string = `9.4.xlsx` // to activate Pulling Activity
	//testCableOptiqueC2Xlsx string = ``

	// Common ==========================================================================================================
	testSiteId int = 57
)

var EnableDestBPECable = map[string]string{
	//"ELINE": "CABLE_%dFO_IMMEUBLE_M6_G657A2",
}

func main() {
	pm := zone.New()
	// configure activity
	pm.DoPulling = true

	pm.DoJunctions = true
	pm.DoEline = true
	pm.DoOtherThanEline = true

	pm.DoMeasurement = true

	// Uncomment zone BlobPattern setup for Sogetrel
	pm.BlobPattern = zone.Blobpattern_Sogetrel // for Sogetrel Worksite

	log.Printf("Parse BPE directory\n")
	err := pm.ParseBPEDir(filepath.Join(testDir, testBPEDir))
	if err != nil {
		log.Fatal("could not parse BPE Directory:", err)
	}

	log.Printf("Parse ROP file\n")
	ropFile := filepath.Join(testDir, testROPXlsx)
	if exists(ropFile) {
		// If ROP File exist, parse it to create BPE Tree
		err = pm.ParseROPXLS(ropFile) // rp.zone.Nodes["PT 204558"]
		if err != nil {
			log.Fatal("could not parse ROP file:", err)
		}

		fmt.Print(pm.Sro.Tree("- ", "", 0))

	} else {
		fmt.Printf("Unable to find ROP file '%s'... creating BPE Tree\n", ropFile)
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

	if testCableOptiqueC2Xlsx != "" {
		cableC2File := filepath.Join(testDir, testCableOptiqueC2Xlsx)
		if !exists(cableC2File) {
			log.Fatal("cable file '%s' does not exist\n", cableC2File)
		}
		err = pm.ParseQuantiteCableOptiqueC2Xlsx(cableC2File)
		if err != nil {
			log.Fatal("could not parse Quantité Cable Optique C2 file:", err)
		}
	}

	pm.CheckConsistency()

	// Force CableType on selected Troncons (used for Immeuble Pulling activity)
	if len(EnableDestBPECable) > 0 {
		pm.EnableCables(EnableDestBPECable)
	}

	err = pm.WriteXLS(testDir, testXLS)
	if err != nil {
		log.Printf("could not write XLSx : %s", err)
	}

	err = pm.WriteJSON(testDir, testXLS, testClient, testManager, testSiteId)
	if err != nil {
		log.Fatal("could not write JSON file :", err)
	}
}

func exists(file string) bool {
	_, err := os.Stat(file)
	return !os.IsNotExist(err)
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Archive
//

//testXLS         string = `VOL_PM03`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCCE_VOL_PM03`
//testBPEDir      string = `CCCE_VOL_PM03_BPE`
//testROPXlsx     string = `CCCE_VOL_PM03_ROP\CCCE_VOL_PM03_ROP.xlsx`
//testCable94Xlsx string = ``
////
//testXLS         string = `MET_PM04`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCPP_MET_PM04`
//testBPEDir      string = `CCPP_MET_PM04_BPE`
//testROPXlsx     string = `CCPP_MET_PM04_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `LON_PM05`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCDUF_LON_PM05`
//testBPEDir      string = `CCDUF_LON_PM05_BPE`
//testROPXlsx     string = `CCDUF_LON_PM05_ROP\CCDUF_LON_PM05_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `GUE_PM17`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCAM_GUE_PM17`
//testBPEDir      string = `CCAM_GUE _PM17_BPE`
//testROPXlsx     string = `CCAM_GUE _PM17_ROP\CCAM_GUE_PM17_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `BOUL_PM04`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCPB_BOU_PM04`
//testBPEDir      string = `CCPB_BOU_PM04_BPE`
//testROPXlsx     string = `CCPB_BOU_PM04_ROP\CCPB_BOU_PM4_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `BOUL_PM03`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCPB_BOU_PM03`
//testBPEDir      string = `CCPB_BOU_PM03_BPE`
//testROPXlsx     string = `CCPB_BOU_PM03_ROP\CCPB_BOU_PM03_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `BOUL_PM01`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCPB_BOU_PM01`
//testBPEDir      string = `CCPB_BOU_PM1_BPE`
//testROPXlsx     string = `CCPB_BOU_PM1_ROP\CCPB_BOU_PM1_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `BOUL_PM05`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCPB_BOU_PM05`
//testBPEDir      string = `CCPB_BOU_PM05_BPE`
//testROPXlsx     string = `CCPB_BOU_PM05_ROP\CCPB_BOU_PM05_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `GUE_PM14`
//testDir         string = `C:\Users\Laurent\Desktop\DOSSIERS MOSELLE\CCAM_GUE_PM14`
//testBPEDir      string = `CCAM_GUE_PM14_BPE`
//testROPXlsx     string = `CCAM_GUE_PM14_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `GUE_PM05`
//testDir         string = `C:\Users\Laurent\Desktop\GUE_PM05`
//testBPEDir      string = `CCAM_GUE_PM5_BPE`
//testROPXlsx     string = `CCAM_GUE_PM5_ROP.xlsx`
//testCable94Xlsx string = ``
//
//testXLS         string = `BOU_PM7`
//testDir         string = `C:\Users\Laurent\Desktop\CCB_BOU_PM8`
//testBPEDir      string = `CCB_BOU_PM08_BPE`
//testROPXlsx     string = `CCB_BOU_PM08_ROP\CCB_BOU_PM08_ROP.xlsx`
//testCable94Xlsx string = `9-4.xlsx`
//
//testXLS         string = `BOU_PM7`
//testDir         string = `C:\Users\Laurent\Desktop\CCB_BOU_PM7`
//testBPEDir      string = `CCB_BOU_PM07_BPE`
//testROPXlsx     string = `CCB_BOU_PM07_ROP\CCB_BOU_PM07_ROP.xlsx`
//testCable94Xlsx string = `Quantité_CCB_BOU_PM07\9-4.xlsx`
//
//testXLS         string = `BOU_PM13`
//testDir         string = `C:\Users\Laurent\Desktop\CCB_BOU_PM13`
//testBPEDir      string = `CCB_BOU_PM13_BPE`
//testROPXlsx     string = `CCB_BOU_PM13_ROP\CCB_BOU_PM13.xlsx`
//testCable94Xlsx string = `CCB_BOU_QTE\9_4.xlsx`
//
//testXLS         string = `BOU_PM12`
//testDir         string = `C:\Users\Laurent\Desktop\CCB_BOU_PM12`
//testBPEDir      string = `CCB_BOU_PM12_BPE`
//testROPXlsx     string = `CCB_BOU_PM12_ROP\CCB_BOU_PM12_ROP.xlsx`
//testCable94Xlsx string = `qtités_12\9_4.xlsx`
//
//testXLS         string = `BOU_PM11`
//testDir         string = `C:\Users\Laurent\Desktop\CCB_BOU_PM11`
//testBPEDir      string = `CCB_BOU_PM11_BPE`
//testROPXlsx     string = `CCB_BOU_PM11_ROP\CCB_BOU_PM11_ROP.xlsx`
//testCable94Xlsx string = `qtités_11\9_4.xlsx`

//testXLS         string = `BOU_PM10`
//testDir         string = `C:\Users\Laurent\Desktop\CCB_BOU_PM10`
//testBPEDir      string = `CCB_BOU_PM10_BPE`
//testROPXlsx     string = `CCB_BOU_PM10_ROP\CCB_BOU_PM10_ROP.xlsx`
//testCable94Xlsx string = `qtités_10\9_4.xlsx`

//testXLS         string = `KOE_PM04`
//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\DMaussand - KOE_PM04\Infos`
//testBPEDir      string = `CCAM_KOE_PM4_BPE`
//testROPXlsx     string = `CCAM_KOE_PM4_ROP\CCAM_KOE_PM4_ROP.xlsx`
//testCable94Xlsx string = `CCAM_KOE_PM4_quantités\94.xlsx`

//testXLS         string = `GUE_PM20`
//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\MBeck - GUE_PM20\info`
//testBPEDir      string = `CCAM_GUE_PM20_BPE`
//testROPXlsx     string = `CCAM_GUE_PM20_ROP\CCAM_GUE_PM20_ROP\CCAM_GUE_PM20_ROP.xlsx`
//testCable94Xlsx string = ``

//testXLS         string = `SOL_PM01`
//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\JTestevuide - SOLGNE\info`
//testBPEDir      string = `CCSM_SOL_PM1_BPE`
//testROPXlsx     string = `CCSM_SOl_PM1_ROP\CCSM_SOL_PM01_ROP.xlsx`
//testCable94Xlsx string = `CCSM_SOL_PM1_Qtté\9,4 PM1.xlsx`

//testXLS         string = `GUE_TR`
//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\MBeck - GUENANGE\Info`
//testBPEDir      string = `CCAM_GUE_TR_BPE`
//testROPXlsx     string = `CCAM_GUE_TR_ROP.xlsx`
//testCable94Xlsx string = ``

//testXLS         string = `AUD_TR`
//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\MBeck - AUDUN-LE-TICHE\Info`
//testBPEDir      string = `BPE`
//testROPXlsx     string = `CCPHVA_AUD_TR_ROP.xlsx`
//testCable94Xlsx string = ``

// GUE_PM03
//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\MBeck - GUE_PM03\Infos`
//testBPEDir      string = `CCAM_GUE_PM3_BPE`
//testROPXlsx     string = `CCAM_GUE_PM3_ROP.xlsx`
//testCable94Xlsx string = ``
//testXLS         string = `GUE_PM03`

// DES_PM03
//testDir         string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\JTestevuide - DESSELING PM3\Infos PM3`
//testBPEDir      string = `CCPE_DES_PM3_BPE`
//testROPXlsx     string = `CCPE_DES_PM3_ROP\CCPE_DES_PM3_ROP.xlsx`
//testCable94Xlsx string = ``
//testXLS         string = `DES_PM03`

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

package main

import (
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/suivi"
	"log"
	"path/filepath"
)

const (
	// Moselle PM3
	//testDir           string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\JTestevuide - DESSELING PM3\Suivi PM3\`
	//bpuFile           string = `BPU.xlsx`
	//suiviFile         string = `DES_PM3_suivi_04-09.xlsx`
	//suiviOutFile      string = `PM3_Suivi.xlsx`

	// Moselle PM3
	testDir      string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\DMaussand - KEDANGE\Suivi`
	bpuFile      string = `BPU.xlsx`
	suiviFile    string = `KED_PM09_suivi_19-04-26  S16.xlsx`
	suiviOutFile string = `KED_PM09_Suivi.xlsx`

	// Alsace DXC
	//testDir           string = `C:\Users\Laurent\Desktop\Suivi`
	//bpuFile           string = `BPU Axians Alsace.xlsx`
	//suiviFile         string = `DXC_Suivi Equipe v2 - MAJ27 FEB S09.xlsx`
	//suiviOutFile      string = `DXC_Suivi.xlsx`

	// Alsace ECF
	//testDir           string = `C:\Users\Laurent\Desktop\Suivi`
	//bpuFile           string = `BPU Axians Alsace.xlsx`
	//suiviFile         string = `ECF_suivi_19-03-20 S12.xlsx`
	//suiviOutFile      string = `ECF_Suivi.xlsx`

	// Alsace ECE
	//testDir           string = `C:\Users\Laurent\Desktop\Suivi`
	//bpuFile           string = `BPU Axians Alsace.xlsx`
	//suiviFile         string = `ECE_suivi_19-02-20.xlsx`
	//suiviOutFile      string = `ECE_Suivi.xlsx`
)

func main() {
	pricecat, err := bpu.NewCatalogFromXLS(filepath.Join(testDir, bpuFile))
	if err != nil {
		log.Fatalf("could not create PriceCatalog: %s", err.Error())
	}

	_ = pricecat
	//fmt.Print(pricecat.String())

	progress, perr := suivi.NewSuiviFromXLS(filepath.Join(testDir, suiviFile), pricecat)
	if perr.HasError() {
		log.Fatalf("%s: could not create progress: %s", suiviFile, perr.Error())
	}

	err = progress.WriteSuiviXLS(filepath.Join(testDir, suiviOutFile))
	if err != nil {
		log.Fatalf("%s: could not Write Suivi XLS: %s", suiviOutFile, err.Error())
	}

	//err = progress.WriteAttachmentXLS(filepath.Join(testDir, attachmentOutFile))
	//if err != nil {
	//	log.Fatalf("%s: could not Write Suivi XLS: %s", attachmentOutFile, err.Error())
	//}
}

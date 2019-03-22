package main

import (
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/suivi"
	"log"
	"path/filepath"
)

const (
	// Moselle PM3
	//testDir           string = `C:\Users\Laurent\Desktop\Suivi PM3`
	//bpuFile           string = `BPU Axians Moselle.xlsx`
	//suiviFile         string = `PM3_suivi 02_15.xlsx`
	//suiviOutFile      string = `PM3_Suivi.xlsx`
	//attachmentOutFile string = `PM3_Attachement.xlsx`

	// Alsace DXC
	testDir           string = `C:\Users\Laurent\Desktop\Suivi`
	bpuFile           string = `BPU Axians Alsace.xlsx`
	suiviFile         string = `DXC_Suivi Equipe v2 - MAJ27 FEB S09.xlsx`
	suiviOutFile      string = `DXC_Suivi.xlsx`
	attachmentOutFile string = `DXC_Attachement.xlsx`

	// Alsace ECF
	//testDir           string = `C:\Users\Laurent\Desktop\Suivi`
	//bpuFile           string = `BPU Axians Alsace.xlsx`
	//suiviFile         string = `ECF_suivi_19-03-20 S12.xlsx`
	//suiviOutFile      string = `ECF_Suivi.xlsx`
	//attachmentOutFile string = `ECF_Attachement.xlsx`

	// Alsace ECE
	//testDir           string = `C:\Users\Laurent\Desktop\Suivi`
	//bpuFile           string = `BPU Axians Alsace.xlsx`
	//suiviFile         string = `ECE_suivi_19-02-20.xlsx`
	//suiviOutFile      string = `ECE_Suivi.xlsx`
	//attachmentOutFile string = `ECE_Attachement.xlsx`
)

func main() {
	pricecat, err := bpu.NewBpuFromXLS(filepath.Join(testDir, bpuFile))
	if err != nil {
		log.Fatalf("could not create PriceCatalog: %s", err.Error())
	}

	_ = pricecat
	//fmt.Print(pricecat.String())

	progress, err := suivi.NewSuiviFromXLS(filepath.Join(testDir, suiviFile), pricecat)
	if err != nil {
		log.Fatalf("%s: could not create progress: %s", suiviFile, err.Error())
	}

	err = progress.WriteSuiviXLS(filepath.Join(testDir, suiviOutFile))
	if err != nil {
		log.Fatalf("%s: could not Write Suivi XLS: %s", suiviOutFile, err.Error())
	}

	err = progress.WriteAttachmentXLS(filepath.Join(testDir, attachmentOutFile), pricecat)
	if err != nil {
		log.Fatalf("%s: could not Write Suivi XLS: %s", attachmentOutFile, err.Error())
	}
}

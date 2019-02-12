package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/suivi"
	"log"
	"path/filepath"
)

const (
	testDir   string = `C:\Users\Laurent\Desktop\Suivi`
	bpuFile          = `BPU Axians Alsace v1.xlsx`
	suiviFile        = `DXC_Suivi Equipe v2 - MAJ 08 FEB S06.xlsx`
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
		log.Fatalf("could not create progress: %s", err.Error())
	}

	fmt.Print(progress.String())

}

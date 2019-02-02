package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/tealeg/xlsx"
	"gopkg.in/src-d/go-vitess.v1/vt/log"
	"os"
	"path/filepath"
	"sort"
)

type PM struct {
	Nodes []*node.Node
}

const (
	blobpattern string = `*PT*.xlsx`
)

func (p *PM) ParseDir(dir string) error {
	parseBlobPattern := filepath.Join(dir, blobpattern)
	files, err := filepath.Glob(parseBlobPattern)
	if err != nil {
		return err
	}
	for _, f := range files {
		n := node.NewNode()
		err := n.ParseXLS(f)
		if err != nil {
			return fmt.Errorf("parsing '%s' returned error : %s\n", filepath.Base(f), err.Error())
		}
		fmt.Printf("'%s' parsed\n", n.PtName)
		p.Nodes = append(p.Nodes, n)
	}
	return nil
}

func (p *PM) WriteXLS(dir, name string) error {
	if len(p.Nodes) == 0 {
		return fmt.Errorf("PM is empty, nothing to write to XLSx")
	}
	file := filepath.Join(dir, name+"_suivi.xlsx")

	xlsx.SetDefaultFont(11, "Calibri")
	xls := xlsx.NewFile()
	sheet, err := xls.AddSheet(name)
	if err != nil {
		return err
	}

	p.Nodes[0].WriteHeader(sheet)
	sort.Slice(p.Nodes, func(i, j int) bool {
		return p.Nodes[i].PtName < p.Nodes[j].PtName
	})
	for _, s := range p.Nodes {
		s.WriteXLS(sheet)
	}

	of, err := os.Create(file)
	if err != nil {
		return err
	}
	defer of.Close()

	return xls.Write(of)
}

const (
	testDir string = `C:\Users\Laurent\Desktop\CCPE_DES_PM3_BPE`
	testXLS string = `PM3`
)

func main() {
	pm := PM{}

	err := pm.ParseDir(testDir)
	if err != nil {
		log.Fatal("could not parse :", err)
	}

	err = pm.WriteXLS(testDir, testXLS)
	if err != nil {
		log.Fatal("could not write XLSx :", err)
	}
}

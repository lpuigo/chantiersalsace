package zone

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"strings"
)

type Zone struct {
	Nodes     []*node.Node
	Sro       *node.Node
	NodeRoots []*node.Node
}

func New() *Zone {
	z := &Zone{
		Nodes: []*node.Node{},
		Sro:   node.NewNode(),
	}
	z.Sro.Name = "SRO"
	z.Sro.PtName = "SRO"
	z.Sro.BPEType = "SRO"
	z.Sro.CableIn = node.NewCable("Aduction")
	return z
}

const (
	blobpattern string = `*PT*.xlsx`
)

func (z *Zone) GetNodeByPtName(ptname string) *node.Node {
	for _, n := range z.Nodes {
		if n.PtName == ptname {
			return n
		}
	}
	return nil
}

func (z *Zone) ParseBPEDir(dir string) error {
	parseBlobPattern := filepath.Join(dir, blobpattern)
	files, err := filepath.Glob(parseBlobPattern)
	if err != nil {
		return err
	}
	for _, f := range files {
		if strings.HasPrefix(filepath.Base(f), "~") {
			continue
		}
		n := node.NewNode()
		fname := filepath.Base(f)
		_ = fname
		err := n.ParseBPEXLS(f)
		if err != nil {
			return fmt.Errorf("parsing '%s' returned error : %s\n", filepath.Base(f), err.Error())
		}
		fmt.Printf("'%s' parsed\n", n.PtName)
		z.Nodes = append(z.Nodes, n)
	}
	return nil
}

func (z *Zone) WriteXLS(dir, name string) error {
	if len(z.Nodes) == 0 {
		return fmt.Errorf("zone is empty, nothing to write to XLSx")
	}
	file := filepath.Join(dir, name+"_suivi.xlsx")

	xlsx.SetDefaultFont(11, "Calibri")
	xls := xlsx.NewFile()
	sheet, err := xls.AddSheet(name)
	if err != nil {
		return err
	}

	z.Nodes[0].WriteHeader(sheet)

	z.Sro.WriteXLS(sheet)

	of, err := os.Create(file)
	if err != nil {
		return err
	}
	defer of.Close()

	return xls.Write(of)
}

func (z *Zone) ParseROPXLS(file string) error {
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}

	var sheet *xlsx.Sheet
	for _, sh := range xls.Sheets {
		if strings.HasPrefix(sh.Name, "TAB") {
			sheet = sh
		}
	}
	if sheet == nil {
		return fmt.Errorf("could not find Tab sheet")
	}

	//parse sheet
	rp := NewRopParser(sheet, z)
	rp.pos = Pos{1, 6}

	rp.ParseRop()

	return nil
}

func (z *Zone) CreateBPETree() {
	// Create Cable list with source / dest node
	cables := map[string]Link{}
	for _, nod := range z.Nodes {
		if nod.CableIn != nil && nod.CableIn.Name != "" {
			link := cables[nod.CableIn.Name]
			link.Dest = nod
			cables[nod.CableIn.Name] = link
		}
		for cableName, cable := range nod.CablesOut {
			if nod.CableIn != nil && nod.CableIn.Name == cable.Name {
				continue
			}
			link := cables[cableName]
			link.Source = nod
			cables[cableName] = link
		}
	}
	// Populate Nodes Children
	for _, link := range cables {
		if link.Source != nil && link.Dest != nil {
			link.Source.Children = append(link.Source.Children, link.Dest)
			link.Dest.IsChild = true
		}
	}
	// Detect Root Nodes
	for _, nod := range z.Nodes {
		if nod.IsChild {
			continue
		}
		z.NodeRoots = append(z.NodeRoots, nod)
	}
}

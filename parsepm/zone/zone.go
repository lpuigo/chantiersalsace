package zone

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

type Zone struct {
	Nodes     node.Nodes
	Troncons  node.Troncons
	Sro       *node.Node
	NodeRoots []*node.Node
}

func New() *Zone {
	z := &Zone{
		Nodes:     node.NewNodes(),
		Troncons:  node.NewTroncons(),
		NodeRoots: []*node.Node{},
		Sro:       node.NewNode(),
	}
	z.Sro.Name = "SRO"
	z.Sro.PtName = "SRO"
	z.Sro.BPEType = "SRO"
	z.Sro.TronconIn = node.NewTroncon("Aduction")
	return z
}

const (
	blobpattern string = `*PT*.xlsx`
)

func (z *Zone) ParseBPEDir(dir string) error {
	parseBlobPattern := filepath.Join(dir, blobpattern)
	files, err := filepath.Glob(parseBlobPattern)
	if err != nil {
		return err
	}
	for _, f := range files {
		// skip XLS temp files
		if strings.HasPrefix(filepath.Base(f), "~") {
			continue
		}
		fname := filepath.Base(f)
		_ = fname
		n, err := z.ParseBPEXLS(f)
		if err != nil {
			return fmt.Errorf("parsing '%s' returned error : %s\n", filepath.Base(f), err.Error())
		}
		fmt.Printf("'%s' parsed\n", n.PtName)
		z.Nodes.Add(n)
	}

	return nil
}

const (
	rowBpePtName = 1
	colBpePtName = 8
	rowBPEType   = 4
	colBPEType   = 1
	rowAddress   = 2
	colAddress   = 8

	colFiberNumIn   = 11
	colFiberNumOut  = 19
	colCableNameIn  = 3
	colCableNameOut = 24
	colOperation    = 13
	colTubulure     = 17
	colCableDict    = 18
)

func (z *Zone) ParseBPEXLS(file string) (*node.Node, error) {
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return nil, err
	}

	sheet := xls.Sheets[0]
	if !strings.HasPrefix(sheet.Name, "Plan ") {
		return nil, fmt.Errorf("Unexpected Sheet name: '%s'", sheet.Name)
	}

	// n.Name
	n := node.NewNode()
	n.PtName = sheet.Cell(rowBpePtName, colBpePtName).Value
	n.BPEType = sheet.Cell(rowBPEType, colBPEType).Value
	// n.LocationType
	n.Address = sheet.Cell(rowAddress, colAddress).Value

	var tronconIn, tronconOut string
	var CableDictZone bool
	// Scan all fiber info rows
	for row := 9; row < sheet.MaxRow; row++ {
		if CableDictZone {
			tronconInfo := sheet.Cell(row, colCableDict).Value
			if tronconInfo == "" {
				continue
			}
			infos := strings.Split(tronconInfo, "-")
			if len(infos) < 2 {
				return nil, fmt.Errorf("could not parse Troncon Info line %d : '%s'", row+1, tronconInfo)
			}
			nt := node.NewTroncon(infos[1])
			capa := strings.Split(infos[0], " ")[0]
			nbFo, err := strconv.ParseInt(capa, 10, 64)
			if err != nil {
				return nil, fmt.Errorf("could not parse Troncon Capa Info line %d : %s", row+1, err.Error())
			}
			nt.Capa = int(nbFo)
			n.TronconsOut[infos[1]] = nt
			continue
		}
		fiberIn := sheet.Cell(row, colFiberNumIn).Value
		fiberOut := sheet.Cell(row, colFiberNumOut).Value
		ope := sheet.Cell(row, colOperation).Value
		nTronconIn := sheet.Cell(row, colCableNameIn).Value
		if nTronconIn != "" && tronconIn != nTronconIn {
			if n.TronconIn != nil {
				return nil, fmt.Errorf("multiple Troncon In found line %d : %s", row+1, err.Error())
			}
			tronconIn = nTronconIn
			n.TronconIn = z.Troncons.Get(tronconIn)
			n.TronconIn.NodeDest = n
		}
		nTronconOut := sheet.Cell(row, colCableNameOut).Value
		if nTronconOut != "" && tronconOut != nTronconOut {
			tronconOut = nTronconOut
			tro := z.Troncons.Get(tronconOut)
			tro.NodeSource = n
			n.TronconsOut[tronconOut] = tro
		}

		if fiberIn != "" || fiberOut != "" { // Input or Output Troncon info available, process it
			n.AddOperation(tronconIn, ope, fiberOut, tronconOut)
		}

		tube := sheet.Cell(row, colTubulure).Value
		if strings.HasPrefix(tube, "Affectation des") {
			CableDictZone = true
			row += 2
		}
	}
	return n, nil
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

	node.NewNode().WriteHeader(sheet)

	if len(z.Sro.Children) > 0 {
		z.Sro.WriteXLS(sheet)
	} else {
		for _, rootnode := range z.NodeRoots {
			rootnode.WriteXLS(sheet)
		}
	}

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
	// Create Troncon list with source / dest node
	cables := map[string]Link{}
	for _, nod := range z.Nodes {
		if nod.TronconIn != nil && nod.TronconIn.Name != "" {
			link := cables[nod.TronconIn.Name]
			link.Dest = nod
			cables[nod.TronconIn.Name] = link
		}
		for cableName, cable := range nod.TronconsOut {
			if nod.TronconIn != nil && nod.TronconIn.Name == cable.Name {
				continue
			}
			link := cables[cableName]
			link.Source = nod
			cables[cableName] = link
		}
	}
	// Populate Nodes Children
	for cableName, link := range cables {
		if link.Source != nil {
			if link.Dest != nil {
				link.Source.Children = append(link.Source.Children, link.Dest)
				link.Dest.IsChild = true
			} else {
				link.Source.AddPMChild(link.Source.TronconsOut[cableName])
			}
		}
	}
	// Detect and attach root nodes to new PM (TODO : detect PM using PM Splice file)
	for _, nod := range z.Nodes {
		if nod.IsChild {
			continue
		}
		if nod.TronconIn != nil && nod.TronconIn.Name != "" {
			z.NodeRoots = append(z.NodeRoots, node.NewPMNode(nod))
		} else {
			z.NodeRoots = append(z.NodeRoots, nod)
		}
	}
}

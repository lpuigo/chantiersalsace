package node

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"sort"
	"strconv"
	"strings"
)

type Node struct {
	Name         string // 3001 (??)
	PtName       string //  PT 182002 (1,8)
	BPEType      string // TENIO T1 (4, 1)
	LocationType string // Chambre Orange (??)
	Address      string // 0, FERME DU TOUPET AZOUDANGE (2,8)

	CableIn   *Cable
	CablesOut Cables

	Children []*Node
	IsChild  bool
}

func NewNode() *Node {
	n := &Node{
		CablesOut: NewCables(),
	}
	return n
}

func NewPMNode(child *Node) *Node {
	pm := NewNode()
	pm.Name = "PM"
	pm.PtName = "PT PM"
	pm.BPEType = "PM"
	pm.LocationType = "PM"
	//pm.Address = ""
	//pm.CableIn = nil
	if child != nil && child.CableIn.Name != "" {
		pm.Children = []*Node{child}
		//pm.IsChild = false
		child.IsChild = true

		cableIn := NewCable("")
		cableIn.Operation["Epissure->"+child.CableIn.Name] = child.CableIn.Capa
		pm.CableIn = cableIn
		pm.CablesOut[""] = cableIn
		cableOut := NewCable(child.CableIn.Name)
		cableOut.Capa = child.CableIn.Capa
		pm.CablesOut[child.CableIn.Name] = cableOut
	}
	return pm
}

const (
	rowPtName  = 1
	colPtName  = 8
	rowBPEType = 4
	colBPEType = 1
	rowAddress = 2
	colAddress = 8

	colFiberNumIn   = 11
	colFiberNumOut  = 19
	colCableNameIn  = 3
	colCableNameOut = 24
	colOperation    = 13
	colTubulure     = 17
	colCableDict    = 18
)

func (n *Node) AddChild(cn *Node) {
	for _, cc := range n.Children {
		if cc.PtName == cn.PtName {
			return
		}
	}
	n.Children = append(n.Children, cn)
}

func (n *Node) AddPMChild(cable *Cable) {
	cpm := NewPMNode(nil)
	cpm.CableIn = cable
	n.Children = append(n.Children, cpm)
}

func (n *Node) GetChildren() []*Node {
	sort.Slice(n.Children, func(i, j int) bool {
		return n.Children[i].PtName < n.Children[j].PtName
	})
	return n.Children
}

func (n *Node) ParseBPEXLS(file string) error {
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}

	sheet := xls.Sheets[0]
	if !strings.HasPrefix(sheet.Name, "Plan ") {
		return fmt.Errorf("Unexpected Sheet name: '%s'", sheet.Name)
	}

	// n.Name
	n.PtName = sheet.Cell(rowPtName, colPtName).Value
	n.BPEType = sheet.Cell(rowBPEType, colBPEType).Value
	// n.LocationType
	n.Address = sheet.Cell(rowAddress, colAddress).Value

	var cableIn, cableOut string
	var CableDict bool
	cableIns := NewCables()
	// Scan all fiber info rows
	for row := 9; row < sheet.MaxRow; row++ {
		if CableDict {
			cableInfo := sheet.Cell(row, colCableDict).Value
			if cableInfo == "" {
				continue
			}
			infos := strings.Split(cableInfo, "-")
			if len(infos) < 2 {
				return fmt.Errorf("could not parse Cable Info line %d : '%s'", row+1, cableInfo)
			}
			nc := NewCable(infos[1])
			capa := strings.Split(infos[0], " ")[0]
			nbFo, err := strconv.ParseInt(capa, 10, 64)
			if err != nil {
				return fmt.Errorf("could not parse Cable Info line %d : %s", row+1, err.Error())
			}
			nc.Capa = int(nbFo)
			n.CablesOut[infos[1]] = nc
			continue
		}
		fiberIn := sheet.Cell(row, colFiberNumIn).Value
		fiberOut := sheet.Cell(row, colFiberNumOut).Value
		ope := sheet.Cell(row, colOperation).Value
		ncableIn := sheet.Cell(row, colCableNameIn).Value
		if ncableIn != "" {
			cableIn = ncableIn
		}
		ncableOut := sheet.Cell(row, colCableNameOut).Value
		if ncableOut != "" {
			cableOut = ncableOut
		}
		tube := sheet.Cell(row, colTubulure).Value

		if fiberIn != "" || fiberOut != "" { // Input or Output Cable info available, process it
			cableIns.Add(cableIn, ope, fiberOut, cableOut)
		}

		if strings.HasPrefix(tube, "Affectation des") {
			CableDict = true
			row += 2
		}
	}

	if len(cableIns) == 0 {
		return fmt.Errorf("could not find cable In info")
	}
	for _, cable := range cableIns {
		if len(cable.Operation) > 0 {
			if n.CableIn != nil {
				return fmt.Errorf("could not define unique cable In info")
			}
			n.CableIn = cable
		}
	}
	return nil
}

func (n *Node) String(co Cables) string {
	res := ""
	res += fmt.Sprintf("%s : cableIn=%s", n.PtName, n.CableIn.String(n.CablesOut))
	return res
}

func (n *Node) Tree(prefix, header string, level int) string {
	res := fmt.Sprintf("%s%d '%s' (%s): %d children\n", header, level, n.Name, n.PtName, len(n.Children))
	for _, cn := range n.GetChildren() {
		res += cn.Tree(prefix, header+prefix, level+1)
	}
	return res
}

func (n *Node) WriteHeader(xs *xlsx.Sheet) {
	type col struct {
		title string
		width float64
	}

	cols := []col{
		{"Nom Site", 12},
		{"Adresse", 40},
		{"Type Boitier", 15},
		{"Type Site", 17},
		{"Ref Site", 10},
		{"Cable entrant", 15},
		{"Taille", 8},

		{"Opérations", 20},
		{"Nb Fibre Sortant", 15},
		{"Nb Epissure", 15},
	}

	r := xs.AddRow()
	for i, ci := range cols {
		r.AddCell().SetString(ci.title)
		xs.Col(i).Width = ci.width
	}
}

const (
	colAerien       string = "fffde9d9"
	colSouterrain   string = "ffdfedda"
	colSansEpissure string = "ffb7dee8"
	colImmeuble     string = "ffe4dfec"
	coldefault      string = "ffff8800"
)

func (n *Node) WriteXLS(xs *xlsx.Sheet) {
	n.writeSiteInfo(xs.AddRow())
	for _, opname := range n.CableIn.Operations() {
		r := xs.AddRow()
		n.writeSitePrefix(r)
		r.AddCell().SetString(n.GetOperationCapa(opname))
		r.AddCell().SetString(opname)
		epi, other := n.CableIn.GetOperationNumbers(opname)
		r.AddCell().SetInt(other + epi)
		r.AddCell().SetInt(epi)

		st := xlsx.NewStyle()
		//st.Fill = *xlsx.NewFill("solid", s.Color, "00000000")
		st.Font = *xlsx.NewFont(10, "Calibri")
		st.Font.Color = "FF6F6F6F"
		st.ApplyFont = true
		for i := 6; i < 10; i++ {
			r.Cells[i].SetStyle(st)
		}
	}

	for _, cnode := range n.GetChildren() {
		cnode.WriteXLS(xs)
	}
}

func (n *Node) writeSiteInfo(r *xlsx.Row) {
	r.AddCell().SetString(n.PtName)
	r.AddCell().SetString(n.Address)
	r.AddCell().SetString(n.BPEType)
	r.AddCell().SetString(n.LocationType)
	r.AddCell().SetString(n.Name)
	r.AddCell().SetString(n.CableIn.Name)
	r.AddCell().SetString(n.CableIn.CapaString())
	r.AddCell().SetString("TOTAL")
	epi, other := n.GetNumbers()
	r.AddCell().SetInt(other + epi)
	r.AddCell().SetInt(epi)

	color := colSouterrain
	//locType := strings.ToLower(strings.TrimSpace(n.LocationType))
	switch {
	case epi == 0:
		color = colSansEpissure
		//case strings.HasPrefix(locType, "poteau"):
		//	color = colAerien
		//case strings.HasPrefix(locType, "app"):
		//	color = colAerien
		//case strings.HasPrefix(locType, "ancr"):
		//	color = colAerien
		//case strings.HasPrefix(locType, "imm"):
		//	color = colImmeuble
	}
	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", color, "00000000")
	st.ApplyFill = true
	for i := 0; i < 10; i++ {
		r.Cells[i].SetStyle(st)
	}

}

func (n *Node) GetNumbers() (nbEpi, nbOther int) {
	return n.CableIn.GetNumbers()
}

func (n *Node) writeSitePrefix(r *xlsx.Row) {
	for i := 0; i < 6; i++ {
		r.AddCell()
	}
}

func (n *Node) GetOperationCapa(ope string) string {
	if !strings.Contains(ope, "->") {
		return ""
	}
	cname := strings.Split(ope, "->")[1]
	return n.CablesOut[cname].CapaString()
}

func (n *Node) SetOperationFromChildren() {
	for _, cn := range n.Children {
		n.CableIn.Capa += cn.CableIn.Capa
		key := "Epissure->" + cn.CableIn.Name
		n.CableIn.Operation[key] = cn.CableIn.Capa

		n.CablesOut[cn.CableIn.Name] = cn.CableIn
	}
}

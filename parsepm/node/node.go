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
	DistFromPM   int

	TronconIn   *Troncon
	TronconsOut Troncons
	Operation   map[string]int

	StartDrawer string
	EndDrawer   string
	SplicePT    []string

	Children []*Node
	IsChild  bool
}

func NewNode() *Node {
	n := &Node{
		TronconsOut: NewTroncons(),
		Operation:   map[string]int{},
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
	//pm.TronconIn = nil
	if child != nil && child.TronconIn.Name != "" {
		pm.Children = []*Node{child}
		//pm.IsChild = false
		child.IsChild = true

		tronconIn := NewTroncon("")
		pm.Operation["Epissure->"+child.TronconIn.Name] = child.TronconIn.Capa
		pm.TronconIn = tronconIn
		pm.TronconsOut[""] = tronconIn
		tronconOut := NewTroncon(child.TronconIn.Name)
		tronconOut.Capa = child.TronconIn.Capa
		pm.TronconsOut[child.TronconIn.Name] = tronconOut
	}
	return pm
}

func (n *Node) SetLocationType() {
	nbEpi, _ := n.GetNumbers()
	if nbEpi > 0 {
		n.LocationType = "BPE"
		return
	}
	n.LocationType = "PBO"
}

// GetToBeMeasuredFiber returns number of measurement to be done on receiver node
//
// - for PM node : number of "Epissure" Operation
//
// - for BPE/PBO node : number of "Attente" Operation
func (n *Node) GetToBeMeasuredFiber() int {
	if n.DistFromPM == 0 {
		return 0
	}
	if n.LocationType == "PM" {
		return n.Operation["Epissure"]
	}
	return n.Operation["Attente"]
}

func (n *Node) AddOperation(tronconIn, ope, fiberOut, tronconOut string) {
	if ope == "Love" || ope == "" {
		return
	}

	key := strings.Title(strings.ToLower(ope))
	if fiberOut != "" {
		if tronconIn == "" {
			key += "<-" + tronconOut
		} else {
			key += "->" + tronconOut
		}
	}
	n.Operation[key]++
}

func (n *Node) AddNbOperations(ope, tronconOut string, nb int) {
	key := strings.Title(strings.ToLower(ope)) + "->" + tronconOut
	n.Operation[key] += nb
}

func (n *Node) AddChild(cn *Node) {
	for _, cc := range n.Children {
		if cc.PtName == cn.PtName {
			return
		}
	}
	n.Children = append(n.Children, cn)
	cn.TronconIn.NodeSource = n
	n.TronconsOut.Add(cn.TronconIn)
}

func (n *Node) AddDrawerInfo(drawer string) {
	if n.StartDrawer == "" {
		n.StartDrawer = drawer
		n.EndDrawer = drawer
	} else {
		if n.StartDrawer > drawer {
			n.StartDrawer = drawer
		}
		if n.EndDrawer < drawer {
			n.EndDrawer = drawer
		}
	}
}

func (n *Node) AddPMChild(cable *Troncon) {
	cpm := NewPMNode(nil)
	cpm.TronconIn = cable
	n.Children = append(n.Children, cpm)
}

func (n *Node) GetChildren() []*Node {
	sort.Slice(n.Children, func(i, j int) bool {
		return n.Children[i].PtName < n.Children[j].PtName
	})
	return n.Children
}

func (n *Node) Operations() []string {
	res := []string{}
	for ope, _ := range n.Operation {
		res = append(res, ope)
	}
	sort.Strings(res)
	return res
}

func (n *Node) String(co Troncons) string {
	res := ""
	res += fmt.Sprintf("%s : cableIn=%s", n.PtName, n.TronconIn.String(n.TronconsOut))
	for _, ope := range n.Operations() {
		if !strings.Contains(ope, "->") {
			res += fmt.Sprintf("\n\t%s : %d", ope, n.Operation[ope])
			continue
		}
		cname := strings.Split(ope, "->")[1]
		res += fmt.Sprintf("\n\t%s (%s): %d", ope, co[cname].CapaString(), n.Operation[ope])
	}
	return res
}

// GetNumbers returns numbers of Epissure and Attente (or others) to be done on receiver node
func (n *Node) GetNumbers() (nbEpi, nbOther int) {
	for ope, _ := range n.Operation {
		e, o := n.GetOperationNumbers(ope)
		nbEpi += e
		nbOther += o
	}
	return
}

func (n *Node) GetOperationNumbers(ope string) (nbEpi, nbOther int) {
	lope := strings.ToLower(ope)
	switch {
	case strings.HasPrefix(lope, "epissure"):
		nbEpi += n.Operation[ope]
	default:
		nbOther += n.Operation[ope]
	}
	return
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

func (n *Node) ParseBPEXLS(file string, troncons Troncons) error {
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}

	sheet := xls.Sheets[0]
	if !strings.HasPrefix(sheet.Name, "Plan ") {
		return fmt.Errorf("Unexpected Sheet name: '%s'", sheet.Name)
	}

	// n.Name
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
				return fmt.Errorf("could not parse Troncon Info line %d : '%s'", row+1, tronconInfo)
			}
			if len(infos) > 2 {
				infos = []string{infos[0], strings.Join(infos[1:], "-")}
			}
			nt := troncons.Get(infos[1])
			capa := strings.Split(infos[0], " ")[0]
			nbFo, err := strconv.ParseInt(capa, 10, 64)
			if err != nil {
				return fmt.Errorf("could not parse Troncon Capa Info line %d : %s", row+1, err.Error())
			}
			nt.Capa = int(nbFo)
			if n.TronconIn == nil {
				return fmt.Errorf("Node has no TronconIn")
			}
			if infos[1] != n.TronconIn.Name {
				n.TronconsOut[infos[1]] = nt
			}
			continue
		}
		fiberIn := sheet.Cell(row, colFiberNumIn).Value
		fiberOut := sheet.Cell(row, colFiberNumOut).Value
		ope := sheet.Cell(row, colOperation).Value
		nTronconIn := sheet.Cell(row, colCableNameIn).Value
		if nTronconIn != "" && tronconIn != nTronconIn {
			if !(ope == "Love" && fiberIn != "" && fiberOut != "") {
				if n.TronconIn != nil {
					return fmt.Errorf("multiple Troncon In found line %d : %s", row+1, nTronconIn)
				}
				tronconIn = nTronconIn
				n.TronconIn = troncons.Get(tronconIn)
				n.TronconIn.NodeDest = n
			}
		}
		nTronconOut := sheet.Cell(row, colCableNameOut).Value
		if nTronconOut != "" && tronconOut != nTronconOut {
			tronconOut = nTronconOut
			tro := troncons.Get(tronconOut)
			tro.NodeSource = n
			n.TronconsOut[tronconOut] = tro
		}

		if fiberIn != "" || fiberOut != "" { // Input or Output Troncon info available, process it
			n.AddOperation(tronconIn, ope, fiberOut, tronconOut)
		}

		// detect "Affectation des tubulures" blocks (troncon list)
		tube := sheet.Cell(row, colTubulure).Value
		if strings.HasPrefix(tube, "Affectation des") {
			CableDictZone = true
			row += 2
		}
	}
	n.SetLocationType()
	return nil
}

func (n *Node) Tree(prefix, header string, level int) string {
	res := fmt.Sprintf("%s%d '%s' (%s): %d children\n", header, level, n.Name, n.PtName, len(n.Children))
	for _, cn := range n.GetChildren() {
		res += cn.Tree(prefix, header+prefix, level+1)
	}
	return res
}

func (n *Node) GetOperationCapa(ope string) string {
	if !strings.Contains(ope, "->") {
		return ""
	}
	cname := strings.Split(ope, "->")[1]
	return n.TronconsOut[cname].CapaString()
}

// SetOperationFromChildren defines reciever node Operations (Epissure) based on children's TroconIn capacity
//
// applicable on PM or SRO Node only
func (n *Node) SetOperationFromChildren() {
	for _, cn := range n.Children {
		n.TronconIn.Capa += cn.TronconIn.Capa
		key := "Epissure->" + cn.TronconIn.Name
		n.Operation[key] = cn.TronconIn.Capa
	}
}

// SetSplicePTs sets SplicePT on receiver node and its children
func (n *Node) SetSplicePTs(splicePT ...string) {
	if n.GetToBeMeasuredFiber() > 0 {
		n.SplicePT = splicePT
	}
	splicedChildren := n.getSplicedChildren()
	for _, cnode := range n.Children {
		nSplicePT := splicePT
		if splicedChildren[cnode.PtName] {
			nSplicePT = append(nSplicePT, n.PtName)
		}
		cnode.SetSplicePTs(nSplicePT...)
	}
}

func (n *Node) getSplicedChildren() map[string]bool {
	res := map[string]bool{}
	for ope, _ := range n.Operation {
		if strings.HasPrefix(ope, "Epissure->") {
			dest := strings.TrimPrefix(ope, "Epissure->")
			if dest != "" {
				tronconDest, found := n.TronconsOut[dest]
				if !found {
					fmt.Printf("Error : unable to find dest %s\n", dest)
				} else {
					res[tronconDest.NodeDest.PtName] = true
				}
			}
		}
	}
	return res
}

// SpliceTRs returns slice of next Troncons with Epissure operation (empty slice if not found)
func (n *Node) SpliceTRs() []*Troncon {
	res := []*Troncon{}
	for ope, _ := range n.Operation {
		if strings.HasPrefix(ope, "Epissure->") {
			res = append(res, n.TronconsOut[strings.TrimPrefix(ope, "Epissure->")])
		}
	}
	return res
}

// GetTronconPassage returns next Troncon with Passage operation, or nil if not found
//
// (unicity is not checked, first troncon found will be returned)
func (n *Node) GetTronconPassage() *Troncon {
	for ope, _ := range n.Operation {
		if strings.HasPrefix(ope, "Passage->") {
			return n.TronconsOut[strings.TrimPrefix(ope, "Passage->")]
		}
	}
	return nil
}

func (n *Node) WriteRaccoHeader(xs *xlsx.Sheet) {
	cols := []col{
		{"Nom Site", 12},
		{"Adresse", 40},
		{"Type Boitier", 15},
		{"Type Site", 17},
		{"Ref Site", 10},
		{"Troncon entrant", 15},
		{"Taille", 8},
		{"Opérations", 20},
		{"Nb Fibre Sortant", 15},
		{"Nb Epissure", 15},

		{"Statut", 15},
		{"Acteur(s)", 15},
		{"N° Déplacement", 15},
		{"Début", 15},
		{"Fin", 15},
	}
	addHeaderRow(xs, cols)
}

const (
	colAerien       string = "fffde9d9"
	colPM           string = "fffde9d9"
	colSouterrain   string = "ffdfedda"
	colPBO          string = "ffdfedda"
	colSansEpissure string = "ffb7dee8"
	colBPE          string = "ffb7dee8"
	colImmeuble     string = "ffe4dfec"
	coldefault      string = "ffff8800"
	colError        string = "ffff0000"

	nbColRacco   int = 10
	nbColMeasure int = 6
)

func (n *Node) WriteRaccoXLS(xs *xlsx.Sheet) {
	n.writeSiteRaccoInfo(xs.AddRow())
	for _, opname := range n.Operations() {
		r := xs.AddRow()
		writeSitePrefix(r, 6)
		r.AddCell().SetString(n.GetOperationCapa(opname))
		r.AddCell().SetString(opname)
		epi, other := n.GetOperationNumbers(opname)
		r.AddCell().SetInt(other + epi)
		r.AddCell().SetInt(epi)

		st := xlsx.NewStyle()
		//st.Fill = *xlsx.NewFill("solid", s.Color, "00000000")
		st.Font = *xlsx.NewFont(10, "Calibri")
		st.Font.Color = "FF6F6F6F"
		st.ApplyFont = true
		addStyleOnRow(r, st, nbColRacco)
	}

	for _, cnode := range n.GetChildren() {
		cnode.WriteRaccoXLS(xs)
	}
}

func (n *Node) writeSiteRaccoInfo(r *xlsx.Row) {
	epi, other := n.GetNumbers()
	color := colBPE
	//locType := strings.ToLower(strings.TrimSpace(n.LocationType))
	switch n.LocationType {
	case "PM":
		color = colPM
	case "PBO":
		color = colPBO
	}

	r.AddCell().SetString(n.PtName)
	r.AddCell().SetString(n.Address)
	r.AddCell().SetString(n.BPEType)
	r.AddCell().SetString(n.LocationType)
	r.AddCell().SetString(n.Name)
	r.AddCell().SetString(n.TronconIn.Name)
	r.AddCell().SetString(n.TronconIn.CapaString())
	r.AddCell().SetString("TOTAL")
	r.AddCell().SetInt(other + epi)
	r.AddCell().SetInt(epi)

	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", color, "00000000")
	st.ApplyFill = true
	addStyleOnRow(r, st, nbColRacco)
}

func (n *Node) WriteMesuresHeader(xs *xlsx.Sheet) {
	cols := []col{
		{"PT cible", 12},
		{"Nb Fibres", 15},
		{"Nb Episs.", 15},
		{"Distance", 15},
		{"Conn. Deb.", 20},
		{"Conn. Fin.", 40},

		{"Statut", 15},
		{"Acteur(s)", 15},
		{"N° Déplacement", 15},
		{"Début", 15},
		{"Fin", 15},
	}
	addHeaderRow(xs, cols)
}

func (n *Node) WriteMesuresXLS(xs *xlsx.Sheet, nodes Nodes) {
	wf := n.GetToBeMeasuredFiber()
	if wf > 0 {
		n.writeMesuresInfo(xs, wf)
		for i, ptName := range n.SplicePT {
			pt := nodes[ptName]
			r := xs.AddRow()
			writeSitePrefix(r, 2)
			r.AddCell().SetInt(i + 1)
			r.AddCell().SetInt(pt.DistFromPM)
			r.AddCell().SetString(ptName)
			r.AddCell().SetString(pt.Address)

			st := xlsx.NewStyle()
			st.Font = *xlsx.NewFont(10, "Calibri")
			st.Font.Color = "FF6F6F6F"
			st.ApplyFont = true
			addStyleOnRow(r, st, nbColMeasure)
		}
	}

	for _, cnode := range n.GetChildren() {
		cnode.WriteMesuresXLS(xs, nodes)
	}
}

func (n *Node) writeMesuresInfo(xs *xlsx.Sheet, nbWaiting int) {
	r := xs.AddRow()
	r.AddCell().SetString(n.PtName)
	r.AddCell().SetInt(nbWaiting)
	r.AddCell().SetInt(len(n.SplicePT))
	r.AddCell().SetInt(n.DistFromPM)
	r.AddCell().SetString(n.StartDrawer)
	r.AddCell().SetString(n.EndDrawer)

	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", colPM, "00000000")
	st.ApplyFill = true
	addStyleOnRow(r, st, nbColMeasure)
}

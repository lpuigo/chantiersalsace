package node

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"sort"
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

func (n *Node) AddOperation(tronconIn, ope, fiberOut, tronconOut string) {
	if ope == "Love" || ope == "" {
		return
	}
	if tronconIn != "" {
		n.TronconIn.Capa++
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

func (n *Node) AddChild(cn *Node) {
	for _, cc := range n.Children {
		if cc.PtName == cn.PtName {
			return
		}
	}
	n.Children = append(n.Children, cn)
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

	r := xs.AddRow()
	for i, ci := range cols {
		r.AddCell().SetString(ci.title)
		xs.Col(i).Width = ci.width
	}
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

	nbCol int = 10
)

func (n *Node) WriteXLS(xs *xlsx.Sheet) {
	n.writeSiteInfo(xs.AddRow())
	for _, opname := range n.Operations() {
		r := xs.AddRow()
		n.writeSitePrefix(r)
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
		for i := 6; i < 10; i++ {
			r.Cells[i].SetStyle(st)
		}
	}

	for _, cnode := range n.GetChildren() {
		cnode.WriteXLS(xs)
	}
}

func (n *Node) writeSiteInfo(r *xlsx.Row) {
	epi, other := n.GetNumbers()
	color := colBPE
	locationType := "BPE"
	//locType := strings.ToLower(strings.TrimSpace(n.LocationType))
	switch {
	case n.Name == "SRO":
		color = colPM
		locationType = "PM"
	case epi == 0:
		color = colPBO
		locationType = "PBO"
		//case strings.HasPrefix(locType, "poteau"):
		//	color = colAerien
		//case strings.HasPrefix(locType, "app"):
		//	color = colAerien
		//case strings.HasPrefix(locType, "ancr"):
		//	color = colAerien
		//case strings.HasPrefix(locType, "imm"):
		//	color = colImmeuble
	}

	r.AddCell().SetString(n.PtName)
	r.AddCell().SetString(n.Address)
	r.AddCell().SetString(n.BPEType)
	//r.AddCell().SetString(n.LocationType) // attribute not set at parse time ... use business rule to assert value instead
	r.AddCell().SetString(locationType)
	r.AddCell().SetString(n.Name)
	r.AddCell().SetString(n.TronconIn.Name)
	r.AddCell().SetString(n.TronconIn.CapaString())
	r.AddCell().SetString("TOTAL")
	r.AddCell().SetInt(other + epi)
	r.AddCell().SetInt(epi)

	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", color, "00000000")
	st.ApplyFill = true
	for i := 0; i < nbCol; i++ {
		r.Cells[i].SetStyle(st)
	}
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
	return n.TronconsOut[cname].CapaString()
}

func (n *Node) SetOperationFromChildren() {
	for _, cn := range n.Children {
		n.TronconIn.Capa += cn.TronconIn.Capa
		key := "Epissure->" + cn.TronconIn.Name
		n.Operation[key] = cn.TronconIn.Capa

		n.TronconsOut[cn.TronconIn.Name] = cn.TronconIn
	}
}

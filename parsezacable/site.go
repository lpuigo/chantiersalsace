package parsezacable

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"sort"
	"strings"
)

const (
	lovage  string = "<Lovage>"
	passage string = "<Passage>"
)

type Dest struct {
	cable string
	capa  string
}

type Link map[string]int //Map[Ope (row, 4)]int

func NewLink() Link {
	return make(Link)
}

// GetNumbers returns Nb of Epi and Others operation for the given Link
func (l Link) GetNumbers() (nbEpi, nbOthers int) {
	for op, n := range l {
		switch op {
		case "EPI":
			nbEpi += n
		default:
			nbOthers += n
		}
	}
	return
}

type Site struct {
	FullName     string // PBO-68-048-DXA-1010 (0-0)
	Name         string // 1010
	BPEType      string // TENIO T1 (0,1)
	LocationType string // Chambre Orange (1,0)
	LocationRef  string // 68248/43 (1,1)

	CapaIn  string // 144FO (>CAPA , 0)
	CableIn string // CDI-68-048-DXA-1010 (>CABLE , 1)

	Links map[Dest]Link //Map[DestSite (row, 9)]Link
}

func NewSite(fname string) *Site {
	s := &Site{
		FullName: fname,
		Name:     GetShortSiteName(fname), //keep 1010 in PBO-68-048-DXA-1010
		Links:    make(map[Dest]Link),
	}
	return s
}

func GetShortSiteName(fullname string) string {
	bs := strings.Split(fullname, "-")
	return strings.Join(bs[4:], "-")
}

// GetNumbers returns Nb of Epi and Others operation for the given Site
func (s Site) GetNumbers() (nbEpi, nbOthers int) {
	for _, l := range s.Links {
		lepi, lother := l.GetNumbers()
		nbEpi += lepi
		nbOthers += lother
	}
	return
}

func (s Site) GetLinksNames() []Dest {
	ll := []Dest{}
	for l, _ := range s.Links {
		ll = append(ll, l)
	}
	sort.Slice(ll, func(i, j int) bool {
		return ll[i].cable < ll[j].cable
	})
	return ll
}

func (s Site) String() string {
	res := fmt.Sprintf("Site %s (short %s):\n", s.FullName, s.Name)
	res += fmt.Sprintf("\tBPEType: '%s', Location Type: '%s', Ref: '%s'\n", s.BPEType, s.LocationType, s.LocationRef)
	res += fmt.Sprintf("\tCapa In: '%s', Cable In: '%s'\n", s.CapaIn, s.CableIn)
	for _, l := range s.GetLinksNames() {
		res += fmt.Sprintf("\t\tLien vers '%s': %v\n", l, s.Links[l])
	}
	return res
}

func (s *Site) AddLink(ope string, cableout Dest) {
	if ope == "" {
		ope = "<none>"
	}
	//if cableout.cable == "" {
	switch ope {
	case "LOV":
		cableout.cable = lovage
		cableout.capa = ""
	case "PAS":
		cableout.cable = passage
		cableout.capa = ""
	default:
		if cableout.cable == "" {
			return // others ope should have a defined cableout, skip this link
		}
	}
	l, found := s.Links[cableout]
	if !found {
		l = NewLink()
		s.Links[cableout] = l
	}

	l[ope]++
}

func (s *Site) ParseXLSSheet(xsh *xlsx.Sheet) error {
	//name := xsh.Cell(0, 0).Value
	//if name != s.FullName {
	//	return fmt.Errorf("site fullname does not match XLS info ('%s' vs '%s)", s.FullName, name)
	//}

	cableout := Dest{}
	ope := ""
	s.BPEType = xsh.Cell(0, 1).Value
	s.LocationType = xsh.Cell(1, 0).Value
	s.LocationRef = xsh.Cell(1, 1).Value
	s.CapaIn = xsh.Cell(4, 0).Value
	s.CableIn = xsh.Cell(4, 1).Value

	for row := 4; row < len(xsh.Rows); row++ {
		cable := xsh.Cell(row, 1).Value
		if cable == "" {
			break
		}
		nope := xsh.Cell(row, 4).Value
		if nope == "" {
			cableout = Dest{}
			ope = nope
			continue
		}
		if nope != ope {
			cableout = Dest{}
			ope = nope
		}
		nco := xsh.Cell(row, 7).Value
		ncocapa := xsh.Cell(row, 8).Value
		if nco != "" && nco != cableout.cable {
			cableout.cable = nco
			cableout.capa = ncocapa
		}
		if cableout.cable == "" && ope != "LOV" && xsh.Cell(row, 9).Value == s.FullName {
			ope = "LOV"
		}
		s.AddLink(ope, cableout)
	}

	return nil
}

func (s Site) WriteHeader(xs *xlsx.Sheet) {
	type col struct {
		title string
		width float64
	}

	cols := []col{
		{"Nom Site", 22},
		{"Nom Syno", 12},
		{"Type Boitier", 15},
		{"Type Site", 17},
		{"Ref Site", 10},
		{"Cable entrant", 23},
		{"Taille", 10},

		{"Cable sortant", 20},
		{"Nb Fibre Sortant", 15},
		{"Nb Epissure", 15},
	}

	r := xs.AddRow()
	for i, ci := range cols {
		r.AddCell().SetString(ci.title)
		xs.Col(i).Width = ci.width
	}
}

func (s Site) writeSiteInfo(r *xlsx.Row) {
	r.AddCell().SetString(s.FullName)
	r.AddCell().SetString(s.Name)
	r.AddCell().SetString(s.BPEType)
	r.AddCell().SetString(s.LocationType)
	r.AddCell().SetString(s.LocationRef)
	r.AddCell().SetString(s.CableIn)
	r.AddCell().SetString(s.CapaIn)
	r.AddCell().SetString("TOTAL")
	epi, other := s.GetNumbers()
	r.AddCell().SetInt(other + epi)
	r.AddCell().SetInt(epi)

	color := "FFdfedda"
	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", color, "00000000")
	st.ApplyFill = true
	for i := 0; i < 10; i++ {
		r.Cells[i].SetStyle(st)
	}
}

func (s Site) writeSitePrefix(r *xlsx.Row) {
	for i := 0; i < 6; i++ {
		r.AddCell()
	}
}

func (s Site) WriteXLS(xs *xlsx.Sheet) {
	s.writeSiteInfo(xs.AddRow())
	for _, l := range s.GetLinksNames() {
		link := s.Links[l]
		r := xs.AddRow()
		s.writeSitePrefix(r)
		r.AddCell().SetString(l.capa)
		r.AddCell().SetString(l.cable)
		epi, other := link.GetNumbers()
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
}

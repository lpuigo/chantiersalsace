package parsezacable

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"sort"
	"strings"
)

type Link map[string]int //Map[Ope (row, 4)]int

func NewLink() Link {
	return make(Link)
}

type Site struct {
	FullName     string // PBO-68-048-DXA-1010 (0-0)
	Name         string // 1010
	BPEType      string // TENIO T1 (0,1)
	LocationType string // Chambre Orange (1,0)
	LocationRef  string // 68248/43 (1,1)

	CapaIn  string // 144FO (>CAPA , 0)
	CableIn string // CDI-68-048-DXA-1010 (>CABLE , 1)

	Links map[string]Link //Map[DestSite (row, 9)]Link
}

func NewSite(fname string) *Site {
	s := &Site{
		FullName: fname,
		Name:     GetShortName(fname), //keep 1010 in PBO-68-048-DXA-1010
		Links:    make(map[string]Link),
	}
	return s
}

func GetShortName(fullname string) string {
	return strings.Split(fullname, "-")[4]
}

func (s Site) String() string {
	res := fmt.Sprintf("Site %s (short %s):\n", s.FullName, s.Name)
	res += fmt.Sprintf("\tBPEType: '%s', Location Type: '%s', Ref: '%s'\n", s.BPEType, s.LocationType, s.LocationRef)
	res += fmt.Sprintf("\tCapa In: '%s', Cable In: '%s'\n", s.CapaIn, s.CableIn)
	ll := []string{}
	for l, _ := range s.Links {
		ll = append(ll, l)
	}
	sort.Strings(ll)
	for _, l := range ll {
		res += fmt.Sprintf("\t\tLien vers '%s': %v\n", l, s.Links[l])
	}
	return res
}

func (s *Site) AddLink(ope, cableout string) {
	if ope == "" {
		ope = "<none>"
	}
	if cableout == "" {
		cableout = "<none>"
	}
	l, found := s.Links[cableout]
	if !found {
		l = NewLink()
		s.Links[cableout] = l
	}

	l[ope]++
}

func (s *Site) ParseXLSSheet(xsh *xlsx.Sheet) error {
	name := xsh.Cell(0, 0).Value
	if name != s.FullName {
		return fmt.Errorf("site fullname does not match XLS info ('%s' vs '%s)", s.FullName, name)
	}

	cableout := ""

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
		ope := xsh.Cell(row, 4).Value
		if ope == "" {
			continue
		}
		//destsite := xsh.Cell(row, 9).Value
		nco := xsh.Cell(row, 7).Value
		if nco != "" && nco != cableout {
			cableout = nco
		}
		s.AddLink(ope, cableout)
	}

	return nil
}

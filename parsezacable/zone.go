package parsezacable

import (
	"github.com/tealeg/xlsx"
	"sort"
	"strings"
)

type Zone struct {
	FullName string // PBO-68-048-DXA
	Name     string // DXA

	Sites []*Site
	Index map[string]int
}

func NewZone(fullname string) *Zone {
	bs := strings.Split(fullname, "-")

	z := &Zone{
		FullName: strings.Join(bs[:4], "-"),
		Name:     bs[3], //keep DXA in PBO-68-048-DXA
		Sites:    []*Site{},
		Index:    make(map[string]int),
	}
	return z
}

// Add adds the given Site to Zone.
// If 'Site.FullName' already exist in Zone, the new site replace the previous one.
func (z *Zone) Add(s *Site) {
	si, found := z.Index[s.FullName]
	if !found {
		z.Index[s.FullName] = len(z.Sites)
		z.Sites = append(z.Sites, s)
		return
	}
	z.Sites[si] = s
}

// GetSiteByFullName returns the given named site (nil if none exists)
func (z *Zone) GetSiteByFullName(fname string) *Site {
	si, found := z.Index[fname]
	if !found {
		return nil
	}
	return z.Sites[si]
}

func (z *Zone) GetSiteFullNames() []string {
	res := []string{}
	for sn, _ := range z.Index {
		res = append(res, sn)
	}
	sort.Strings(res)
	return res
}

func (z *Zone) ParseXLSSheet(xsh *xlsx.Sheet) error {
	s := NewSite(xsh.Name)
	err := s.ParseXLSSheet(xsh)
	if err != nil {
		return err
	}
	z.Add(s)
	return nil
}

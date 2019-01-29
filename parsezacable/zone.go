package parsezacable

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

type Zone struct {
	Name string // DXA

	Sites     []*Site
	Index     map[string]int
	SynoIndex map[string]int
}

func NewZone(name string) *Zone {
	z := &Zone{
		Name:      name,
		Sites:     []*Site{},
		Index:     make(map[string]int),
		SynoIndex: make(map[string]int),
	}
	return z
}

// GetShortZoneName returns ZoneName out of Box Ref (ex : PBO-68-048-DXA-1010 => DXA)
func GetShortZoneName(fullname string) string {
	return strings.Split(fullname, "-")[3]
}

// Add adds the given Site to Zone.
// If 'Site.FullName' already exist in Zone, the new site replace the previous one.
func (z *Zone) Add(s *Site) {
	si, found := z.Index[s.FullName]
	if !found {
		numSite := len(z.Sites)
		z.Index[s.FullName] = numSite
		z.SynoIndex[s.Name] = numSite
		z.Sites = append(z.Sites, s)
		return
	}
	z.Sites[si] = s
}

// GetSiteByFullName returns the given full-named site (nil if none exists)
func (z *Zone) GetSiteByFullName(fname string) *Site {
	si, found := z.Index[fname]
	if !found {
		return nil
	}
	return z.Sites[si]
}

// GetSiteByName returns the given syno-named site (nil if none exists)
func (z *Zone) GetSiteByName(fname string) *Site {
	si, found := z.SynoIndex[fname]
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

func (z *Zone) ParseXLSFile(file string) error {
	fmt.Printf("Processing file '%s'\n", file)
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		return fmt.Errorf("could not process xlsx file: %v", err)
	}

	for _, xs := range xf.Sheets {
		fmt.Printf("\tParsing sheet %s\n", xs.Name)
		err := z.ParseXLSSheet(xs)
		if err != nil {
			return fmt.Errorf("could not parse sheet %s: %v", xs.Name, err)
		}
	}
	return nil
}

func (z *Zone) ParseBlob(pattern string) error {
	files, err := filepath.Glob(pattern)
	if err != nil {
		return err
	}
	for _, f := range files {
		err = z.ParseXLSFile(f)
		if err != nil {
			return err
		}
	}
	return nil
}

func (z *Zone) WriteXLS(file string) error {
	if len(z.Sites) == 0 {
		return fmt.Errorf("zone is empty, nothing to write to XLSx")
	}
	of, err := os.Create(file)
	if err != nil {
		return err
	}
	defer of.Close()

	xlsx.SetDefaultFont(11, "Calibri")
	oxf := xlsx.NewFile()
	oxs, err := oxf.AddSheet(z.Name)
	if err != nil {
		return err
	}
	z.Sites[0].WriteHeader(oxs)
	sort.Slice(z.Sites, func(i, j int) bool {
		return z.Sites[i].Name < z.Sites[j].Name
	})
	for _, s := range z.Sites {
		s.WriteXLS(oxs)
	}

	return oxf.Write(of)
}

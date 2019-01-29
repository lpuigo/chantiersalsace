package site

import "github.com/tealeg/xlsx"

type Ascendent interface {
	Name() string
	Hierarchy() string
}

type Site struct {
	Type      string
	Id        string
	BPEType   string
	Operation string
	Ref       string
	Ref2      string
	FiberOut  string
	FiberIn   string
	Lenght    string

	Color string

	Parent   Ascendent
	Children []*Site
}

func (s Site) Name() string {
	return s.Id
}

func (s Site) Hierarchy() string {
	if s.Parent == nil {
		return s.Id
	}
	return s.Parent.Hierarchy() + " > " + s.Id
}

func (s Site) NbSite() int {
	res := 1
	for _, c := range s.Children {
		res += c.NbSite()
	}
	return res
}

type SRO struct {
	name     string
	Children []*Site
}

func (sr SRO) Hierarchy() string {
	return sr.name
}

func (sr SRO) Name() string {
	return sr.name
}

func NewSRO(name string) SRO {
	return SRO{
		name:     name,
		Children: []*Site{},
	}
}

func (s Site) WriteXLSHeader(xs *xlsx.Sheet) {

	type col struct {
		title string
		width float64
	}

	cols := []col{
		{"Type", 8},
		{"Id Parent", 18},
		{"Longueur", 10},
		{"Nb Fibre In", 14},
		{"Id Site", 10},
		{"Type Boitier", 15},
		{"Ref", 10},
		{"Ref2", 10},
		{"Nb Fibre Sortant", 14},
		{"Operation", 12},
		{"Nb Cable Sortant", 18},
		{"Status", 12},
		{"Date", 12},
	}

	r := xs.AddRow()
	for i, ci := range cols {
		r.AddCell().SetString(ci.title)
		xs.Col(i).Width = ci.width
	}
}

func (s Site) WriteXLSRow(xs *xlsx.Sheet) {
	r := xs.AddRow()

	r.AddCell().SetString(s.Type)
	r.AddCell().SetString(s.Parent.Name())
	r.AddCell().SetString(s.Lenght)
	r.AddCell().SetString(s.FiberIn)
	r.AddCell().SetString(s.Id)
	r.AddCell().SetString(s.BPEType)
	r.AddCell().SetString(s.Ref)
	r.AddCell().SetString(s.Ref2)
	r.AddCell().SetString(s.FiberOut)
	r.AddCell().SetString(s.Operation)
	r.AddCell().SetInt(len(s.Children))

	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", s.Color, "00000000")
	st.ApplyFill = true
	for i := 4; i < 11; i++ {
		r.Cells[i].SetStyle(st)
	}

	for _, csite := range s.Children {
		csite.WriteXLSRow(xs)
	}
}

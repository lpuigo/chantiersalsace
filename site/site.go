package site

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

	Parent   *Site
	Children []*Site
}

type SRO struct {
	Name     string
	Children []*Site
}

func NewSRO(name string) SRO {
	return SRO{
		Name:     name,
		Children: []*Site{},
	}
}

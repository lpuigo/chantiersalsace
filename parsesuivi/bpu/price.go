package bpu

type Price struct {
	Name   string
	Size   int
	Income map[string]float64
}

func (p *Price) GetBpeValue() float64 {
	return p.Income["box"]
}

func (p *Price) GetSpliceValue(nbsplice int) float64 {
	return p.Income["splice"] * float64(nbsplice)
}

func NewPrice(size int) *Price {
	p := &Price{
		Size:   size,
		Income: make(map[string]float64),
	}
	return p
}

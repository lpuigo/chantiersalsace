package bpu

type Price struct {
	Name  string
	Size  int
	Price float64
}

func NewPrice() *Price {
	p := &Price{}
	return p
}

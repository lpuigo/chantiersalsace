package bpu

type Chapter struct {
	Name  string
	Size  int
	Price float64
}

func NewChapter() *Chapter {
	p := &Chapter{}
	return p
}

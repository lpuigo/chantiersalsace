package bpu

type Article struct {
	Name  string
	Size  int
	Price float64
}

func NewChapter() *Article {
	p := &Article{}
	return p
}

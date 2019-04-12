package bpu

type Box struct {
	Name  string
	Size  int
	Usage string
}

func NewBox() *Box {
	b := &Box{}
	return b
}

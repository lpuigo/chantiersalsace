package node

type Cables []*Cable

func NewCables() Cables {
	return make([]*Cable, 0)
}

func (cs *Cables) Add(c *Cable) {
	*cs = append(*cs, c)
}

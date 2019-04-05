package node

type Cable struct {
	Type     string
	Capa     int
	Length   int
	Troncons []*Troncon
}

func NewCable(tr *Troncon) *Cable {
	nc := &Cable{}
	//nc.AddTroncon(tr, 20)
	return nc
}

func (c *Cable) AddTroncon(t *Troncon, loveDist int) {
	c.Length += t.NodeDest.DistFromPM - t.NodeSource.DistFromPM + loveDist
	c.Troncons = append(c.Troncons, t)
}

func (c *Cable) LastTroncon() *Troncon {
	if len(c.Troncons) > 0 {
		return c.Troncons[len(c.Troncons)-1]
	}
	return nil
}

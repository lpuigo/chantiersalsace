package node

import (
	"fmt"
)

type Troncon struct {
	Name              string
	Capa              int
	CableType         string
	LoveLength        int
	UndergroundLength int
	AerialLength      int
	FacadeLength      int

	NodeSource *Node
	NodeDest   *Node
}

func NewTroncon(name string) *Troncon {
	return &Troncon{Name: name}
}

func (c Troncon) String(co Troncons) string {
	res := ""
	res += fmt.Sprintf("Troncon '%s' (%s)", c.Name, c.CapaString())
	return res
}

func (c Troncon) CapaString() string {
	return fmt.Sprintf("%dFO", c.Capa)
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

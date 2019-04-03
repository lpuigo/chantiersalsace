package node

import "gopkg.in/src-d/go-vitess.v1/vt/log"

type Nodes map[string]*Node

func NewNodes() Nodes {
	return make(map[string]*Node)
}

func (ns Nodes) Add(n *Node) {
	if _, found := ns[n.PtName]; found {
		log.Fatal("adding already existing node %s", n.PtName)
	}
	ns[n.PtName] = n
}

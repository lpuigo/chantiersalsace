package node

type Nodes map[string]*Node

func NewNodes() Nodes {
	return make(map[string]*Node)
}

func (ns Nodes) Add(n *Node) (new bool) {
	_, found := ns[n.PtName]
	if !found {
		ns[n.PtName] = n
		return true
	}
	return false
}

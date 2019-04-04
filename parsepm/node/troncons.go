package node

type Troncons map[string]*Troncon

func NewTroncons() Troncons {
	return make(map[string]*Troncon)
}

func (ts Troncons) Add(t *Troncon) (new bool) {
	_, found := ts[t.Name]
	if !found {
		ts[t.Name] = t
		return true
	}
	return false
}

// Get returns the Troncon having given name (create a new one if not exists)
func (ts Troncons) Get(name string) *Troncon {
	tr, found := ts[name]
	if !found {
		tr = NewTroncon(name)
		ts[name] = tr
	}
	return tr
}

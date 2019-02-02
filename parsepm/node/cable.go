package node

import (
	"fmt"
	"sort"
	"strings"
)

type Cable struct {
	Name      string
	Capa      int
	Operation map[string]int
}

func NewCable(name string) *Cable {
	return &Cable{Name: name,
		Operation: make(map[string]int),
	}
}

func (c Cable) Operations() []string {
	res := []string{}
	for ope, _ := range c.Operation {
		res = append(res, ope)
	}
	sort.Strings(res)
	return res
}

func (c Cable) String(co Cables) string {
	res := ""
	res += fmt.Sprintf("Cable '%s' (%dFO)", c.Name, c.Capa)
	for _, ope := range c.Operations() {
		if !strings.Contains(ope, "->") {
			res += fmt.Sprintf("\n\t%s : %d", ope, c.Operation[ope])
			continue
		}
		cname := strings.Split(ope, "->")[1]
		res += fmt.Sprintf("\n\t%s (%dFO): %d", ope, co[cname].Capa, c.Operation[ope])
	}
	return res
}

type Cables map[string]*Cable

func NewCables() Cables {
	return make(map[string]*Cable)
}

func (cs Cables) Add(name, ope, fo, dest string) {
	cable, found := cs[name]
	if !found {
		cable = NewCable(name)
		cs[name] = cable
	}
	cable.Capa++
	if ope == "Love" || ope == "" {
		return
	}
	key := ope
	if fo != "" {
		key += "->" + dest
	}
	cable.Operation[key]++
}

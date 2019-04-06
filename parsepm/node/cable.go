package node

import "github.com/tealeg/xlsx"

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

const (
	nbColTirage int = 10
)

func (c *Cable) WriteTirageHeader(xs *xlsx.Sheet) {
	cols := []col{
		{"Type Cable", 15},
		{"Tronçon", 12},
		{"PT Départ", 15},
		{"Adr. Départ", 40},
		{"PT Arrivée", 15},
		{"Adr. Arrivée", 40},
		{"Distance Tot", 15},
		{"Souterrain", 15},
		{"Aérien", 15},
		{"Façade", 15},

		{"Statut", 15},
		{"Acteur(s)", 15},
		{"N° Déplacement", 15},
		{"Début", 15},
		{"Fin", 15},
	}
	addHeaderRow(xs, cols)
}

func (c *Cable) WriteTirageXLS(xs *xlsx.Sheet) {
	r := xs.AddRow()
	nBeg := c.Troncons[0].NodeSource
	nEnd := c.LastTroncon().NodeDest

	r.AddCell().SetString(c.Type)
	r.AddCell().SetString(c.Troncons[0].Name)
	r.AddCell().SetString(nBeg.PtName)
	r.AddCell().SetString(nBeg.Address)
	r.AddCell().SetString(nEnd.PtName)
	r.AddCell().SetString(nEnd.Address)
	r.AddCell().SetInt(c.Length)
	r.AddCell().SetString("-")
	r.AddCell().SetString("-")
	r.AddCell().SetString("-")

	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", colPM, "00000000")
	st.ApplyFill = true
	addStyleOnRow(r, st, nbColTirage)

	for _, tr := range c.Troncons {
		r := xs.AddRow()
		r.AddCell()
		r.AddCell().SetString(tr.Name)
		r.AddCell().SetString(tr.NodeSource.PtName)
		r.AddCell().SetString(tr.NodeSource.Address)
		r.AddCell().SetString(tr.NodeDest.PtName)
		r.AddCell().SetString(tr.NodeDest.Address)
		r.AddCell().SetInt(tr.NodeDest.DistFromPM - tr.NodeSource.DistFromPM)
		r.AddCell().SetString("-")
		r.AddCell().SetString("-")
		r.AddCell().SetString("-")

		st := xlsx.NewStyle()
		st.Font = *xlsx.NewFont(10, "Calibri")
		st.Font.Color = "FF6F6F6F"
		st.ApplyFont = true
		addStyleOnRow(r, st, nbColTirage)
	}

}

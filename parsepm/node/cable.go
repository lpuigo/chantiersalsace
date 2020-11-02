package node

import "github.com/tealeg/xlsx"

type Cable struct {
	Capa     int
	Length   int
	Troncons []*Troncon
}

func NewCable(tr *Troncon) *Cable {
	nc := &Cable{Capa: tr.Capa}
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

func (c *Cable) GetLenths() (lov, under, aer, fac int) {
	for _, tr := range c.Troncons {
		lov += tr.LoveLength
		under += tr.UndergroundLength
		aer += tr.AerialLength
		fac += tr.FacadeLength
	}
	return
}

const (
	nbColTirage int = 11
)

func (c *Cable) WriteTirageHeader(xs *xlsx.Sheet) {
	cols := []col{
		{"Type Cable", 36},
		{"Tronçon", 12},
		{"PT Départ", 15},
		{"Adr. Départ", 40},
		{"PT Arrivée", 15},
		{"Adr. Arrivée", 40},
		{"Distance Tot", 15},
		{"Love", 10},
		{"Souterrain", 10},
		{"Aérien", 10},
		{"Façade", 10},

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
	lov, under, aer, fac := c.GetLenths()
	r.AddCell().SetString(c.Troncons[0].CableType)
	r.AddCell().SetString(c.Troncons[0].Name)
	r.AddCell().SetString(nBeg.PtName)
	r.AddCell().SetString(nBeg.Address)
	r.AddCell().SetString(nEnd.PtName)
	r.AddCell().SetString(nEnd.Address)
	r.AddCell().SetInt(lov + under + aer + fac)
	r.AddCell().SetInt(lov)
	r.AddCell().SetInt(under)
	r.AddCell().SetInt(aer)
	r.AddCell().SetInt(fac)

	color := colSouterrain
	if (aer + fac) > 0 {
		color = colAerien
	}
	st := xlsx.NewStyle()
	st.Fill = *xlsx.NewFill("solid", color, "00000000")
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
		r.AddCell().SetInt(tr.LoveLength)
		r.AddCell().SetInt(tr.UndergroundLength)
		r.AddCell().SetInt(tr.AerialLength)
		r.AddCell().SetInt(tr.FacadeLength)

		st := xlsx.NewStyle()
		st.Font = *xlsx.NewFont(10, "Calibri")
		st.Font.Color = "FF6F6F6F"
		st.ApplyFont = true
		addStyleOnRow(r, st, nbColTirage)
	}

}

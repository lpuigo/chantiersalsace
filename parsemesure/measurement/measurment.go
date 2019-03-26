package measurement

import "fmt"

type Measurement struct {
	Name         string
	Wavelength   string
	TotLoss      float64
	Distance     float64
	Evt          int
	MaxSplice    float64
	TotORL       float64
	MaxConnector float64
	LenMaxSplice float64
}

func (m Measurement) String() string {
	res := "Measurement "
	res += fmt.Sprintf("Name: %s\n", m.Name)
	res += fmt.Sprintf("\tWavelength: %s\n", m.Wavelength)
	res += fmt.Sprintf("\tTotLoss: %6.3f\n", m.TotLoss)
	res += fmt.Sprintf("\tDistance: %6.3f\n", m.Distance)
	res += fmt.Sprintf("\tEvt: %d\n", m.Evt)
	res += fmt.Sprintf("\tMaxSplice: %6.3f\n", m.MaxSplice)
	res += fmt.Sprintf("\tTotORL: %6.3f\n", m.TotORL)
	res += fmt.Sprintf("\tMaxConnector: %6.3f\n", m.MaxConnector)
	res += fmt.Sprintf("\tLenMaxSplice: %6.3f\n", m.LenMaxSplice)

	return res
}

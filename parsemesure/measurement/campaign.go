package measurement

import (
	"fmt"
	"io"
)

type Campaign struct {
	Name   string
	PtName string

	Measurements []Measurement
}

func (c Campaign) Write(w io.Writer) error {
	_, _ = fmt.Fprintf(w, "Nb File	%d\r\n\r\n", len(c.Measurements))
	_, _ = fmt.Fprintf(w, "[Results]\r\n")
	_, _ = fmt.Fprintf(w, "Alarms 	Fib #	Dir.	Laser 	Tot loss	 Distance	Evt	Max Splice	Tot ORL	Max Connector 	Len. Max Splice	\r\n")
	for i, m := range c.Measurements {
		_, _ = fmt.Fprintf(w, "       \t%-5d\tO->E\t%-6s\t%-8.3f\t%-9.1f\t%-3d\t%-10s\t%-7s\t%-14.3f\t%-15s\t\r\n",
			i+1, //m.Name,
			m.Wavelength,
			m.TotLoss,
			m.Distance,
			m.Evt,
			formatFloat(".2", m.MaxSplice),
			formatFloat(".2", m.TotORL),
			m.MaxConnector,
			formatFloat(".1", m.LenMaxSplice),
		)
	}
	return nil
}

func formatFloat(format string, val float64) string {
	if val == 0 {
		return "-"
	}
	fm := "%" + format + "f"
	return fmt.Sprintf(fm, val)
}

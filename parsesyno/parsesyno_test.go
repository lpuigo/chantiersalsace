package parsesyno

import (
	"testing"
)

const (
	synotest       = `C:\Users\Laurent\Golang\src\github.com\lpuig\ewin\chantiersalsace\test\test1.xlsx`
	synoresulttest = `C:\Users\Laurent\Golang\src\github.com\lpuig\ewin\chantiersalsace\test\test1result.xlsx`
)

func TestSyno_Parse(t *testing.T) {
	syno := Syno{File: synotest}

	err := syno.Parse()
	if err != nil {
		t.Fatal(err)
	}
	t.Log("Nb Sites found :", syno.nbSites)

	err = syno.WriteXLS(synoresulttest)
	if err != nil {
		t.Fatal("could not write to XLSx :", err)
	}
}

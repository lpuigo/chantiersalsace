package parsesyno

import (
	"testing"
)

const (
	synotest = `C:\Users\Laurent\Golang\src\github.com\lpuig\ewin\chantiersalsace\test\DXA_SYNO.xlsx`
)

func TestSyno_Parse(t *testing.T) {
	syno := Syno{File: synotest}

	err := syno.Parse()
	if err != nil {
		t.Fatal(err)
	}
}

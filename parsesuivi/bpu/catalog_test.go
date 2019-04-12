package bpu

import "testing"

const testBpuXls string = `C:\Users\Laurent\Google Drive (laurent.puig.ewin@gmail.com)\Axians\Axians Moselle\Chantiers\JTestevuide - DESSELING PM3\Suivi PM3\BPU.xlsx`

func TestNewCatalogFromXLS(t *testing.T) {
	bpu, err := NewCatalogFromXLS(testBpuXls)
	if err != nil {
		t.Fatalf("NewCatalogFromXLS returned unexpected: %s", err.Error())
	}
	_ = bpu

}

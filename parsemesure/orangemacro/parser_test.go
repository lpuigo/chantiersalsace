package orangemacro

import (
	"github.com/lpuig/ewin/chantiersalsace/dirbrowser"
	"os"
	"testing"
)

const (
	testFile string = `C:\Users\Laurent\Desktop\Mesures PM3 TR182008\PT182064\PMZ-PB_CCPE_DES_PM03_FI CCPE_DES_PM03_TR182008_PT182064.xlsx`
	testDir  string = `C:\Users\Laurent\Desktop\Mesures PM3 TR182008 reprises - Copie\`
)

func TestParse(t *testing.T) {
	mc, err := Parse(testFile)
	if err != nil {
		t.Fatalf("Parser return unexpected: %s", err.Error())
	}

	mc.Write(os.Stdout)
}

func TestXlsToTxt(t *testing.T) {
	ext := ".xlsx"
	fn := func(file string) error {
		return XlsToTxt(ext, file)
	}
	err := dirbrowser.Process(testDir, ext, fn)
	if err != nil {
		t.Fatalf("Process returned unexpected: %v", err)
	}
}

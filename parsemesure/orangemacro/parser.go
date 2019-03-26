package orangemacro

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsemesure/measurement"
	"github.com/tealeg/xlsx"
	"os"
	"strings"
)

type Parser struct {
	file string
}

const (
	sheetName   string = "Données"
	campaignCol int    = 3
	campaignRow int    = 3
	ptNameCol   int    = 3
	ptNameRow   int    = 6

	measNameCol        int = 2
	measNameStartRow   int = 32
	measTitleRow       int = 28
	measWaveLengthCol  int = 3
	measEventColOffset int = 5
	eventBeginCol      int = 4
	eventResLossCol    int = 6
	reportResLossCol   int = 6
	reportTotLengthCol int = 7
)

func XlsToTxt(fileExt, file string) error {
	mc, err := Parse(file)
	if err != nil {
		return err
	}
	txtFile := strings.TrimPrefix(file, fileExt) + ".txt"
	f, err := os.Create(txtFile)
	if err != nil {
		return err
	}
	defer f.Close()
	return mc.Write(f)
}

func Parse(file string) (mc measurement.Campaign, err error) {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		return
	}

	xs := xf.Sheet[sheetName]
	if xs == nil {
		err = fmt.Errorf("could not find tab '%s' in file '%s'", sheetName, file)
		return
	}

	mc = measurement.Campaign{}

	mc.Name = xs.Cell(campaignRow, campaignCol).Value
	mc.PtName = xs.Cell(ptNameRow, ptNameCol).Value

	// calc number of event
	nbEvent := 0
	for {
		title := xs.Cell(measTitleRow, measEventColOffset*(nbEvent+1)+eventBeginCol).Value
		if title == "" {
			break
		}
		nbEvent++
	}

	row := measNameStartRow
	for {
		measName := xs.Cell(row, measNameCol).Value
		if measName == "" {
			break
		}
		meas := measurement.Measurement{
			Name:       measName,
			Wavelength: xs.Cell(row, measWaveLengthCol).Value,
			TotLoss:    parseFloat(xs.Cell(row, measEventColOffset*(nbEvent+1)+reportResLossCol)),
			Distance:   parseFloat(xs.Cell(row, measEventColOffset*(nbEvent+1)+reportTotLengthCol)),
			//Evt:          0,
			//MaxSplice:    0,
			TotORL:       0,
			MaxConnector: parseFloat(xs.Cell(row, eventResLossCol)),
			LenMaxSplice: 0,
		}
		maxSpliceLoss := 0.0
		maxEvent := 1
		for en := 1; en <= nbEvent; en++ {
			eventSpliceLoss := parseFloat(xs.Cell(row, measEventColOffset*en+eventResLossCol))
			if eventSpliceLoss > maxSpliceLoss {
				maxSpliceLoss = eventSpliceLoss
				maxEvent = en
			}
		}
		meas.Evt = maxEvent
		meas.MaxSplice = maxSpliceLoss
		mc.Measurements = append(mc.Measurements, meas)
		row++
	}

	return
}

func parseFloat(cell *xlsx.Cell) float64 {
	f, err := cell.Float()
	if err != nil {
		return 0
	}
	return f
}

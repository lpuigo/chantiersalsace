package node

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strconv"
	"strings"
)

type Node struct {
	Name         string // 3001 (??)
	PtName       string //  PT 182002 (1,8)
	BPEType      string // TENIO T1 (4, 1)
	LocationType string // Chambre Orange (??)
	Address      string // 0, FERME DU TOUPET AZOUDANGE (2,8)

	cIn   *Cable
	csOut Cables
}

func NewNode() *Node {
	n := &Node{
		csOut: NewCables(),
	}
	return n
}

const (
	rowPtName  = 1
	colPtName  = 8
	rowBPEType = 4
	colBPEType = 1
	rowAddress = 2
	colAddress = 8

	colFiberNumIn   = 11
	colFiberNumOut  = 19
	colCableNameIn  = 3
	colCableNameOut = 24
	colOperation    = 13
	colTubulure     = 17
	colCableDict    = 18
)

func (n *Node) ParseXLS(file string) error {
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}

	sheet := xls.Sheets[0]
	if !strings.HasPrefix(sheet.Name, "Plan ") {
		return fmt.Errorf("Unexpected Sheet name: '%s'", sheet.Name)
	}

	// n.Name
	n.PtName = sheet.Cell(rowPtName, colPtName).Value
	n.BPEType = sheet.Cell(rowBPEType, colBPEType).Value
	// n.LocationType
	n.Address = sheet.Cell(rowAddress, colAddress).Value

	var cableIn, cableOut string
	var CableDict bool
	cableIns := NewCables()
	// Scan all fiber info rows
	for row := 9; row < sheet.MaxRow; row++ {
		if CableDict {
			cableInfo := sheet.Cell(row, colCableDict).Value
			if cableInfo == "" {
				continue
			}
			infos := strings.Split(cableInfo, "-")
			if len(infos) < 2 {
				return fmt.Errorf("could not parse Cable Info line %d : '%s'", row+1, cableInfo)
			}
			nc := NewCable(infos[1])
			capa := strings.Split(infos[0], " ")[0]
			nbFo, err := strconv.ParseInt(capa, 10, 64)
			if err != nil {
				return fmt.Errorf("could not parse Cable Info line %d : %s", row+1, err.Error())
			}
			nc.Capa = int(nbFo)
			n.csOut[infos[1]] = nc
			continue
		}
		fiberIn := sheet.Cell(row, colFiberNumIn).Value
		fiberOut := sheet.Cell(row, colFiberNumOut).Value
		ope := sheet.Cell(row, colOperation).Value
		ncableIn := sheet.Cell(row, colCableNameIn).Value
		if ncableIn != "" {
			cableIn = ncableIn
		}
		ncableOut := sheet.Cell(row, colCableNameOut).Value
		if ncableOut != "" {
			cableOut = ncableOut
		}
		tube := sheet.Cell(row, colTubulure).Value

		if fiberIn != "" { // No Input Cable info available , skip to Output Cable parsing
			cableIns.Add(cableIn, ope, fiberOut, cableOut)
		}

		if strings.HasPrefix(tube, "Affectation des") {
			CableDict = true
			row += 2
		}
	}

	if len(cableIns) == 0 {
		return fmt.Errorf("could not find cable In info")
	}
	for _, cable := range cableIns {
		if len(cable.Operation) > 0 {
			if n.cIn != nil {
				return fmt.Errorf("could not define unique cable In info")
			}
			n.cIn = cable
		}
	}
	return nil
}

func (n *Node) String(co Cables) string {
	res := ""
	res += fmt.Sprintf("%s : cableIn=%s", n.PtName, n.cIn.String(n.csOut))
	return res
}

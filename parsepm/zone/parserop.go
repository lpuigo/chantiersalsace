package zone

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/tealeg/xlsx"
	"log"
	"strconv"
	"strings"
)

type Pos struct {
	row int
	col int
}

func (p Pos) Right(offset int) Pos {
	return Pos{p.row, p.col + offset}
}

type RopParser struct {
	sheet *xlsx.Sheet
	zone  *Zone
	pos   Pos
}

const (
	acolPmName     int = 0
	acolDrawer     int = 3
	acolDrawerLine int = 4
	acolDrawerCol  int = 5
	colTubulure    int = 0
	colCableIn     int = 2
	colName        int = 4
	colPtName      int = 5
	colDistFromPM  int = 6
	colOpe         int = 7
	colNextBlock   int = 8
)

func NewRopParser(sh *xlsx.Sheet, zone *Zone) *RopParser {
	rp := &RopParser{
		sheet: sh,
		zone:  zone,
	}
	return rp
}

func (rp *RopParser) debug(msg string) {
	log.Fatalf("pos %s : %s", xlsx.GetCellIDStringFromCoords(rp.pos.col, rp.pos.row), msg)
}

func (rp *RopParser) GetChildRopParser() *RopParser {
	nrp := *rp
	nrp.pos = rp.pos.Right(colNextBlock)
	return &nrp
}

func (rp *RopParser) GetPosValue(row, col int) string {
	return rp.sheet.Cell(row, col).Value
}

func (rp *RopParser) GetPosInt(row, col int) int {
	i, err := rp.sheet.Cell(row, col).Int()
	if err != nil {
		i = 0
	}
	return i
}

func (rp *RopParser) GetValue(colOffset int) string {
	return rp.sheet.Cell(rp.pos.row, rp.pos.col+colOffset).Value
}

// ChildExists returns true if Child block exists
func (rp *RopParser) ChildExists() bool {
	if rp.GetPosValue(0, rp.pos.col+colNextBlock) == "T" && rp.GetValue(colNextBlock) != "" {
		return true
	}
	return false
}

// SetNodeInfo sets node Name and check TronconIn consistency (creates It if not already defined)
func (rp *RopParser) SetNodeInfo(n *node.Node) {
	n.Name = rp.GetValue(colName)
	// check cable In consistency
	cableInName := rp.GetValue(colCableIn)
	if n.TronconIn != nil && cableInName != n.TronconIn.Name {
		rp.debug(fmt.Sprintf("SetNodeInfo: not matching cable In name '%s' ('%s' expected) for node '%s'", cableInName, n.TronconIn.Name, n.PtName))
	}
	if n.TronconIn == nil {
		trIn := rp.zone.Troncons[cableInName]
		if trIn == nil {
			rp.debug(fmt.Sprintf("SetNodeInfo: could not get troncon from cable '%s'", cableInName))
		}
		if trIn.NodeDest != nil && trIn.NodeDest.PtName != n.PtName {
			rp.debug(fmt.Sprintf("SetNodeInfo: troncon '%s' already has a destination node '%s' instead of '%s'", cableInName, trIn.NodeDest.PtName, n.PtName))
		}
		if trIn.NodeDest == nil {
			trIn.NodeDest = n
			n.TronconIn = trIn
			trIn.NodeSource.AddChild(n)
		}
	}
}

func (rp *RopParser) ParseRop() {
	rp.pos = Pos{1, 6}
	// Init root PM Node
	rp.zone.Sro.PtName = rp.GetPosValue(rp.pos.row, acolPmName)
	rp.zone.Sro.LocationType = "PM"
	rp.zone.Nodes.Add(rp.zone.Sro)

	// Start Parsing
	done := false
	for !done {
		if rp.GetValue(colTubulure) != "" {
			topnode := rp.Parse()
			rp.zone.Sro.AddChild(topnode)
			continue
		}
		if rp.GetValue(-1) == "" {
			done = true
			continue
		}
		rp.pos.row++
	}
	rp.zone.Sro.SetOperationFromChildren()
	rp.zone.Sro.SetSplicePTs()
	rp.zone.DetectCables(rp.zone.Sro)
}

// Parse returns current block Node (populated with all its defined children) and move RopParser pos to the next child within same level
func (rp *RopParser) Parse() *node.Node {
	ptName := rp.GetValue(colPtName)
	currentNode := rp.zone.Nodes[ptName]
	//if currentNode == nil {
	//	currentNode = rp.zone.GetNodeByTronconIn(rp.GetValue(colCableIn))
	//}
	if currentNode == nil {
		if !strings.Contains(ptName, "_PM") { // unknown Pt is not a PM ... exit with debug
			rp.debug(fmt.Sprintf("Parse: could not get node from ptname '%s'", ptName))
		}
		// unknown Pt is a PM, let's create it
		currentNode = node.NewPMNode(nil)
		currentNode.PtName = ptName
		currentNode.Name = rp.GetValue(colName)
		rp.zone.Nodes.Add(currentNode)

		// check if cableIn is already defined (false when PM directly behind NRO => creates it)

		cableInName := rp.GetValue(colCableIn)
		trIn := rp.zone.Troncons[cableInName]
		if trIn == nil {
			newTr := node.NewTroncon(cableInName)
			newTr.NodeSource = rp.zone.Sro
			rp.zone.Troncons.Add(newTr)
		}
	}
	distString := rp.GetValue(colDistFromPM)
	dist, err := strconv.ParseInt(distString, 10, 64)
	if err != nil {
		rp.debug(fmt.Sprintf("Parse: could not get distance from '%s'", distString))
	}
	currentNode.DistFromPM = int(dist)
	rp.SetNodeInfo(currentNode)
	inNode := true
	for inNode {
		if rp.ChildExists() {
			crp := rp.GetChildRopParser()
			childNode := crp.Parse()
			if childNode != nil {
				currentNode.AddChild(childNode)
			}
			rp.pos.row = crp.pos.row
		} else {
			switch rp.GetValue(colOpe) {
			case "ATTENTE":
				drawer := rp.GetPosValue(rp.pos.row, acolDrawer)
				drawerSuffix := drawer
				parts := strings.Split(drawer, "_")
				if len(parts) >= 2 {
					drawerSuffix = strings.Join(parts[len(parts)-2:], "_")
				}
				drawerInfo := fmt.Sprintf("%s/%s/%02d",
					drawerSuffix,
					rp.GetPosValue(rp.pos.row, acolDrawerLine),
					rp.GetPosInt(rp.pos.row, acolDrawerCol),
				)
				currentNode.AddDrawerInfo(drawerInfo)
				if currentNode.LocationType == "PM" {
					currentNode.AddOperation(rp.GetValue(colCableIn), "ATTENTE", "", "")
					if currentNode.TronconIn.NodeSource.LocationType == "PM" {
						currentNode.TronconIn.Capa++
					}
				}
			case "EPISSURE":
				if currentNode.LocationType == "PM" {
					currentNode.AddOperation(rp.GetValue(colCableIn), "EPISSURE", "", "")
					if currentNode.TronconIn.NodeSource.LocationType == "PM" {
						currentNode.TronconIn.Capa++
					}
				}
			}
			rp.pos.row++
		}
		// test if pos is still in same Node
		if rp.GetValue(colPtName) != ptName {
			inNode = false
		}
	}
	return currentNode
}

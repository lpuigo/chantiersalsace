package zone

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/tealeg/xlsx"
)

type Pos struct {
	row int
	col int
}

func (p Pos) Right(offset int) Pos {
	return Pos{p.row, p.col + offset}
}

type RopParser struct {
	sheet          *xlsx.Sheet
	zone           *Zone
	pos            Pos
	serviceCol     int
	CountReserveOR bool
}

const (
	acolPmName     int = 0
	acolDrawer     int = 3
	acolDrawerLine int = 4
	acolDrawerCol  int = 5
	acolFirstChild int = 6
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
	//log.Printf("pos %s : %s", xlsx.GetCellIDStringFromCoords(rp.pos.col, rp.pos.row), msg)
}

func (rp *RopParser) CloneRopParser() *RopParser {
	nrp := *rp
	return &nrp
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

func (rp *RopParser) FindServiceColumn() bool {
	colNum := acolFirstChild
	for rp.GetPosValue(0, colNum) == "T" {
		colNum += colNextBlock
	}
	rp.serviceCol = colNum + 2
	return rp.GetPosValue(0, rp.serviceCol) == "SERVICE"
}

func (rp *RopParser) IsCurrentRouteForClient() bool {
	serviceValue := strings.ToLower(strings.Trim(rp.GetPosValue(rp.pos.row, rp.serviceCol), " "))
	// return forClient || forReserve
	forClient := strings.HasPrefix(serviceValue, "client ftth")
	if rp.CountReserveOR && !forClient {
		return strings.HasPrefix(serviceValue, "reserve")
	}
	return forClient
}

// GetParentPtName return parent PT (or PM) name
func (rp *RopParser) GetParentPtName() string {
	col := rp.pos.col + colPtName - colNextBlock
	if col < acolFirstChild {
		col = acolPmName
	}
	return rp.sheet.Cell(rp.pos.row, col).Value
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
	// set DistFromPM
	distString := rp.GetValue(colDistFromPM)
	dist, err := strconv.ParseInt(distString, 10, 64)
	if err != nil {
		rp.debug(fmt.Sprintf("Parse: could not get distance from '%s'", distString))
	}
	n.DistFromPM = int(dist)

	n.Name = rp.GetValue(colName)
	// check cable In consistency
	cableInName := rp.GetValue(colCableIn)

	// if TronconIn already defined, check its consistency
	if n.TronconIn != nil && cableInName != n.TronconIn.Name {
		if strings.ReplaceAll(cableInName, " ", "") != strings.ReplaceAll(n.TronconIn.Name, " ", "") {
			rp.debug(fmt.Sprintf("SetNodeInfo: not matching cable In name '%s' ('%s' expected) for node '%s'", cableInName, n.TronconIn.Name, n.PtName))
		}
		cableInName = n.TronconIn.Name
	}
	if n.TronconIn == nil {
		// if TronconIn not defined, create it
		trIn := rp.zone.Troncons[cableInName]
		if trIn == nil {
			if !rp.zone.CreateNodeFromRop {
				rp.debug(fmt.Sprintf("SetNodeInfo: could not get troncon from cable '%s'", cableInName))
			}
			trIn = node.NewTroncon(cableInName)
			rp.zone.Troncons.Add(trIn)

			parentNodeName := rp.GetParentPtName()
			parentNode, found := rp.zone.Nodes[parentNodeName]
			if !found {
				rp.debug(fmt.Sprintf("SetNodeInfo: could not get parentNode '%s'", parentNodeName))
			}
			trIn.NodeSource = parentNode
		}
		if trIn.NodeDest == nil {
			trIn.NodeDest = n
			n.TronconIn = trIn
			trIn.NodeSource.AddChild(n)
		}
		if trIn.NodeDest != nil && trIn.NodeDest.PtName != n.PtName {
			rp.debug(fmt.Sprintf("SetNodeInfo: troncon '%s' already has a destination node '%s' instead of '%s'", cableInName, trIn.NodeDest.PtName, n.PtName))
		}
	}
}

func (rp *RopParser) ParseRop() {
	// check for Service column
	if !rp.FindServiceColumn() {
		rp.pos = Pos{0, rp.serviceCol}
		rp.debug("could not find service column")
	}
	rp.pos = Pos{1, acolFirstChild}
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
	if currentNode == nil {
		// the Node is not already defined after processing BPE Directory
		if !strings.Contains(ptName, "_PM") {
			if !rp.zone.CreateNodeFromRop {
				// unknown Pt is not a PM ... exit with debug
				rp.debug(fmt.Sprintf("Parse: could not get node from ptname '%s'", ptName))
			}
			// Unknown Node, creating it from RopFile data
			fmt.Printf("\tCreate node '%s' from Rop File\n", ptName)
			currentNode = node.NewNode()
			currentNode.PtName = ptName
			currentNode.Name = rp.GetValue(colName)
			rp.zone.Nodes.Add(currentNode)
		} else {
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
	}
	rp.SetNodeInfo(currentNode)
	if rp.zone.CreateNodeFromRop && !rp.zone.DefineNodeOperation[ptName] {
		// Operation are to be defined from Rop data
		// reset Operation
		currentNode.Operation = make(map[string]int)
		// mark currentNode operation as reseted
		rp.zone.DefineNodeOperation[ptName] = true
	}
	inNode := true
	for inNode {
		if rp.ChildExists() {
			crp := rp.GetChildRopParser()
			childNode := crp.Parse()
			if childNode != nil {
				currentNode.AddChild(childNode)
			}
			if rp.zone.CreateNodeFromRop {
				// define currentNode Operation for childNode
				nbOpe := crp.pos.row - rp.pos.row
				currentNode.AddNbOperations(rp.GetValue(colOpe), childNode.TronconIn.Name, nbOpe)
			}
			rp.pos.row = crp.pos.row
		} else {
			switch rp.GetValue(colOpe) {
			case "ATTENTE":
				// Drawer management
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
				// Operation management
				if currentNode.LocationType == "PM" {
					currentNode.AddOperation(rp.GetValue(colCableIn), "ATTENTE", "", "")
					if currentNode.TronconIn.NodeSource.LocationType == "PM" {
						currentNode.TronconIn.Capa++
					}
				} else if rp.zone.CreateNodeFromRop {
					// check if current Opt. Route is to bu used
					if rp.IsCurrentRouteForClient() {
						currentNode.AddOperation("", "ATTENTE", "", "")
					}
					// else its operation = love => No Op.
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

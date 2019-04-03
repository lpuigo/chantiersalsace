package zone

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/tealeg/xlsx"
	"gopkg.in/src-d/go-vitess.v1/vt/log"
	"strconv"
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
	colCableIn    int = 2
	colName       int = 4
	colPtName     int = 5
	colDistFromPM int = 6
	colNextBlock  int = 8
)

func NewRopParser(sh *xlsx.Sheet, zone *Zone) *RopParser {
	rp := &RopParser{
		sheet: sh,
		zone:  zone,
	}
	return rp
}

func (rp *RopParser) debug(msg string) {
	log.Fatalf("pos %d, %d : %s", rp.pos.row+1, rp.pos.col+1, msg)
}

func (rp *RopParser) GetChildRopParser() *RopParser {
	nrp := *rp
	nrp.pos = rp.pos.Right(colNextBlock)
	return &nrp
}

func (rp *RopParser) GetPosValue(row, col int) string {
	return rp.sheet.Cell(row, col).Value
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

// SetNodeInfo sets node Name and check TronconIn consistency
func (rp *RopParser) SetNodeInfo(n *node.Node) {
	n.Name = rp.GetValue(colName)
	// check cable In consistency
	cableInName := rp.GetValue(colCableIn)
	if cableInName != n.TronconIn.Name {
		rp.debug(fmt.Sprintf("not matching cable In name '%s' ('%s' expected) for node '%s'", cableInName, n.TronconIn.Name, n.PtName))
	}
}

func (rp *RopParser) ParseRop() {
	done := false
	for !done {
		if rp.GetValue(0) != "" {
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
}

// Parse returns current block Node (populated with all its defined children) and move RopParser pos to the next child within same level
func (rp *RopParser) Parse() *node.Node {
	ptName := rp.GetValue(colPtName)
	currentNode := rp.zone.GetNodeByPtName(ptName)
	if currentNode == nil {
		rp.debug(fmt.Sprintf("could not get node from ptname '%s'", ptName))
	}
	distString := rp.GetValue(colDistFromPM)
	dist, err := strconv.ParseInt(distString, 10, 64)
	if err != nil {
		rp.debug(fmt.Sprintf("could not get distance from '%s'", distString))
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
			rp.pos.row++
		}
		// test if pos is still in same Node
		if rp.GetValue(colPtName) != ptName {
			inNode = false
		}
	}
	return currentNode
}

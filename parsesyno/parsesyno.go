package parsesyno

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/site"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"strings"
)

const (
	SynoSheetName = "Syno"

	SROMaxColNum = 4
)

type Position struct {
	row, col int
}

func (p Position) IsValid() bool {
	return p.row > 0 && p.col > 0
}

type Syno struct {
	File      string
	synoSheet *xlsx.Sheet
	nbCols    int
	nbRows    int

	nbSites int
	Sro     site.SRO
}

func (s *Syno) Parse() error {
	xf, err := xlsx.OpenFile(s.File)
	if err != nil {
		return err
	}
	synoSheet, found := xf.Sheet[SynoSheetName]
	if !found {
		return fmt.Errorf("could not find XLSx Sheet '%s' in '%s'", SynoSheetName, filepath.Base(s.File))
	}
	s.synoSheet = synoSheet

	s.nbCols = len(synoSheet.Cols)
	s.nbRows = len(synoSheet.Rows)

	//for i:= 59; i < 85; i++ {
	//	s.printBorderInfo(Position{i,26})
	//
	//}

	sroPos, sroName := s.GetSROInfo()
	if !sroPos.IsValid() {
		return fmt.Errorf("could not find SRO info")
	}
	fmt.Printf("SRO %s found pos %v\n", sroName, sroPos)
	s.Sro = site.NewSRO(sroName)

	nextPos := s.findFirstChild(Position{sroPos.row, sroPos.col + 2})
	if !nextPos.IsValid() {
		return fmt.Errorf("could not find first Site Position's")
	}
	var nextSite *site.Site
	for {
		nextPos, nextSite = s.GetSite(nextPos, s.Sro)
		s.Sro.Children = append(s.Sro.Children, nextSite)
		nextPos = s.GetSiblingSitePos(nextPos)
		if !nextPos.IsValid() {
			break
		}
	}

	return nil
}

func (s Syno) WriteXLS(file string) error {
	of, err := os.Create(file)
	if err != nil {
		return err
	}
	defer of.Close()

	xlsx.SetDefaultFont(11, "Calibri")
	oxf := xlsx.NewFile()
	oxs, err := oxf.AddSheet(s.Sro.Name())
	if err != nil {
		return err
	}
	hs := &site.Site{}
	hs.WriteXLSHeader(oxs)
	for _, psite := range s.Sro.Children {
		psite.WriteXLSRow(oxs)
	}

	return oxf.Write(of)
}

// GetSROInfo returns rowNumber and Name of SRO (position invalid if not found)
func (s Syno) GetSROInfo() (pos Position, name string) {
	for rn, r := range s.synoSheet.Rows {
		if len(r.Cells) <= SROMaxColNum {
			continue
		}
		for cn, c := range r.Cells {
			val := c.Value
			if !strings.HasPrefix(val, "SRO-") {
				continue
			}
			name = val
			pos = Position{rn, cn}
			return
		}
	}
	return Position{}, ""
}

// GetSiblingSitePos returns Sibing Site Position (or invalid Position if no sibling site found)
func (s Syno) GetSiblingSitePos(curPos Position) Position {
	for {
		left, bottom := s.HasBorder(curPos)
		if !left {
			return Position{} // no other sibling here, exit with invalid pos
		}
		if bottom {
			return curPos // found something
		}
		curPos.row += 1
		if curPos.row >= s.nbRows {
			return Position{} // end of sheet reached, exit with invalid pos
		}
	}
}

func (s Syno) HasBorder(pos Position) (left, bottom bool) {
	cell := s.synoSheet.Cell(pos.row, pos.col)
	border := cell.GetStyle().Border
	left = border.Left != "" && border.Left != "none"
	if !left {
		cell2 := s.synoSheet.Cell(pos.row, pos.col-1)
		border2 := cell2.GetStyle().Border
		left = border2.Right != "" && border2.Right != "none"
	}
	bottom = border.Bottom != "" && border.Bottom != "none"
	if bottom {
		return
	}
	cell2 := s.synoSheet.Cell(pos.row+1, pos.col)
	border2 := cell2.GetStyle().Border
	bottom = border2.Top != "" && border2.Top != "none"
	return
}

func (s Syno) printBorderInfo(pos Position) {
	cell := s.synoSheet.Cell(pos.row, pos.col)
	style := cell.GetStyle()
	border := style.Border
	color := style.Fill.FgColor
	fmt.Printf("Pos %v: l:'%s' r:'%s' t:'%s' b:'%s' val:%s (%s)\n", pos, border.Left, border.Right, border.Top, border.Bottom, cell.Value, color)
}

// GetSite returns the site found from current Position and starting sibling position (under the given curPos)
func (s *Syno) GetSite(startPos Position, parent site.Ascendent) (nextSibling Position, newSite *site.Site) {
	pos := startPos
	// seek (to the right) the site ref
	for s.synoSheet.Cell(pos.row, pos.col).Value == "" && pos.col < s.nbCols {
		pos.col++
	}
	newSite = &site.Site{
		FiberIn:  s.synoSheet.Cell(pos.row+1, pos.col).Value,
		Lenght:   s.synoSheet.Cell(pos.row+2, pos.col).Value,
		Parent:   parent,
		Children: []*site.Site{},
	}
	// seek (to the right) the block site info
	spos := Position{pos.row, pos.col + 1}
	//TODO detect also left border
	for s.synoSheet.Cell(spos.row, spos.col).Value == "" && spos.col < s.nbCols {
		spos.col++
	}
	childPos := s.getSiteBlock(spos, newSite)
	if childPos.IsValid() {
		for {
			nextSiblingPos, childSite := s.GetSite(childPos, newSite)
			newSite.Children = append(newSite.Children, childSite)
			childPos = s.GetSiblingSitePos(nextSiblingPos)
			if !childPos.IsValid() {
				break
			}
		}
	}
	s.nbSites++
	fmt.Printf("found Site %3v <%3v> Id:%-12s #Site:%3d Type:%6s (%s)\t%s\n", pos, startPos, newSite.Id, newSite.NbSite(), newSite.Type, newSite.Color, newSite.Hierarchy())
	nextSibling = Position{startPos.row + 2, startPos.col}
	return
}

// getSiteBlock populates curSite with Block info, and returns first Child Position (invalid position if not found)
func (s Syno) getSiteBlock(pos Position, curSite *site.Site) (nextChild Position) {
	// seek for first row of site block
	for {
		if pos.row == 0 {
			break
		}
		//_, bottom := s.HasBorder(Position{pos.row - 1, pos.col})
		//if bottom {
		//	break
		//}
		color := s.synoSheet.Cell(pos.row-1, pos.col).GetStyle().Fill.FgColor
		if color == "" || color == "FFFFFFFF" {
			break
		}
		pos.row--
	}
	curSite.Type = s.synoSheet.Cell(pos.row, pos.col).Value
	curSite.Id = s.synoSheet.Cell(pos.row+1, pos.col).Value
	curSite.BPEType = s.synoSheet.Cell(pos.row+2, pos.col).Value
	curSite.Operation = s.synoSheet.Cell(pos.row+3, pos.col).Value
	curSite.Ref = s.synoSheet.Cell(pos.row+4, pos.col).Value
	curSite.Ref2 = s.synoSheet.Cell(pos.row+5, pos.col).Value
	curSite.FiberOut = s.synoSheet.Cell(pos.row+6, pos.col).Value
	curSite.Color = s.synoSheet.Cell(pos.row+3, pos.col).GetStyle().Fill.FgColor

	// seek for upwards child
	return s.findFirstChild(Position{pos.row, pos.col + 1})
}

func siteCoords(p Position) Position {
	return Position{}
}

// findFirstChild return highest child position (invalid position if not found)
func (s Syno) findFirstChild(pos Position) Position {
	// seek for upwards child
	if pos.row > 0 {
		upos := pos
		for {
			left, bottom := s.HasBorder(upos)
			/////////////////////////////////////////////////////////////////////////////////////////////////
			//if upos.row > 56 && upos.row < 83 && upos.col == 26 {
			//	fmt.Printf("\t\t%v : left:%v bottom:%v\n", upos, left, bottom)
			//}
			/////////////////////////////////////////////////////////////////////////////////////////////////
			if !left && bottom {
				return upos // found first child pos
			}
			if !left && !bottom {
				break // no child up there
			}
			if upos.row == 0 {
				break
			}
			upos.row--
		}
	}

	// seek for child on the right
	rpos := pos
	for {
		left, bottom := s.HasBorder(rpos)
		if bottom {
			return rpos // found first child pos
		}
		if !left || rpos.row > pos.row+6 {
			break // no child there
		}
		rpos.row++
	}
	return Position{}
}

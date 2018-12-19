package parsesyno

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/site"
	"github.com/tealeg/xlsx"
	"path/filepath"
)

const (
	SynoSheetName = "Syno"

	SROColNum       = 4
	SiteStartColNum = 7
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

	Sro site.SRO
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

	sroPos, sroName := s.GetSROInfo()
	if sroPos == -1 {
		return fmt.Errorf("could not find SRO info")
	}
	s.Sro = site.NewSRO(sroName)

	nextPos := s.GetFirstSitePos()
	if !nextPos.IsValid() {
		return fmt.Errorf("could not find first Site Position's")
	}
	var nextSite *site.Site
	for {
		nextPos, nextSite = s.GetSite(nextPos)
		s.Sro.Children = append(s.Sro.Children, nextSite)
		nextPos = s.GetSiblingSitePos(nextPos)
		if !nextPos.IsValid() {
			break
		}
	}

	return nil
}

// GetSROInfo returns rowNumber and Name of SRO (pos = -1 if not found)
func (s Syno) GetSROInfo() (pos int, name string) {
	for rn, r := range s.synoSheet.Rows {
		if len(r.Cells) <= SROColNum {
			continue
		}
		val := r.Cells[SROColNum].Value
		if val != "" {
			name = val
			pos = rn
			return
		}
	}
	return -1, ""
}

func (s Syno) GetFirstSitePos() Position {
	for rn, r := range s.synoSheet.Rows {
		if len(r.Cells) <= SiteStartColNum {
			continue
		}
		cell := r.Cells[SiteStartColNum]
		pos := Position{rn, SiteStartColNum}
		_, bottom := s.HasBorder(pos)
		if cell.Value != "" || bottom {
			return pos
		}
	}
	return Position{}
}

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
	border := cell.GetStyle().Border
	fmt.Printf("Pos %v: l:'%s' r:'%s' t:'%s' b:'%s' (%s)\n", pos, border.Left, border.Right, border.Top, border.Bottom, cell.Value)
}

// GetSite returns the site found from current Position, starting sibling position (under the given curPos), and starting Child Position
func (s Syno) GetSite(curpos Position) (nextSibling Position, newSite *site.Site) {
	pos := curpos
	// seek (to the right) the site ref
	for s.synoSheet.Cell(pos.row, pos.col).Value == "" && pos.col < s.nbCols {
		pos.col++
	}
	newSite = &site.Site{
		FiberIn:  s.synoSheet.Cell(pos.row+1, pos.col).Value,
		Lenght:   s.synoSheet.Cell(pos.row+2, pos.col).Value,
		Parent:   nil,
		Children: []*site.Site{},
	}
	// seek (to the right) the site info
	spos := Position{pos.row, pos.col + 1}
	for s.synoSheet.Cell(spos.row, spos.col).Value == "" && spos.col < s.nbCols {
		spos.col++
	}
	childPos := s.getSiteBlock(spos, newSite)

	fmt.Printf("found pos %v %s %s (ref %s, color %s), child pos %v\n", pos, newSite.Id, newSite.Type, newSite.Ref, newSite.Color, childPos)
	//TODO Seek for children
	nextSibling = Position{curpos.row + 6, curpos.col}
	return
}

func (s Syno) getSiteBlock(pos Position, curSite *site.Site) (nextChild Position) {
	// seek for first row of site block
	for {
		left, bottom := s.HasBorder(Position{pos.row - 1, pos.col})
		if !left && bottom {
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
	cpos := Position{pos.row - 1, pos.col + 1}
	for {
		left, bottom := s.HasBorder(cpos)
		//fmt.Printf("\t %v l:%v b:%v\n", cpos, left, bottom)
		if !left && bottom {
			return cpos // found first child pos
		}
		if !left && !bottom {
			break // no child up there
		}
		cpos.row--
	}

	// seek for child on the right
	cpos = Position{pos.row + 2, pos.col + 1}
	for {
		left, bottom := s.HasBorder(cpos)
		if left && bottom {
			return cpos // found first child pos
		}
		if !left {
			break // no child up there
		}
		cpos.row++
	}

	return Position{}
}

func siteCoords(p Position) Position {
	return Position{}
}

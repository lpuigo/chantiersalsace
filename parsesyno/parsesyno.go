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
	fmt.Printf("next pos : %v\n", nextPos)
	var nextSite *site.Site
	for {
		nextPos, nextSite = s.GetSite(nextPos)
		fmt.Printf("next pos : %v\n", nextPos)
		s.Sro.Children = append(s.Sro.Children, nextSite)
		nextPos = s.GetSiblingSitePos(nextPos)
		fmt.Printf("sibling pos : %v\n", nextPos)
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
		border := cell.GetStyle().Border
		if cell.Value != "" || (border.Bottom != "" && border.Bottom != "none") {
			return Position{
				row: rn,
				col: SiteStartColNum,
			}
		}
	}
	return Position{}
}

func (s Syno) GetSiblingSitePos(curPos Position) Position {
	for {
		cell := s.synoSheet.Cell(curPos.row, curPos.col)
		border := cell.GetStyle().Border
		fmt.Printf("border at %v : %v\n", curPos, border)
		if border.Left == "none" || border.Left == "" {
			return Position{} // no other sibling here, exit with invalid pos
		}
		if border.Bottom != "" {
			return curPos // found something
		}
		curPos.row += 8
		if curPos.row >= s.nbRows {
			return Position{} // end of sheet reached, exit with invalid pos
		}
	}
}

// GetFirstSite returns the first site found starting from given row
func (s Syno) GetSite(curpos Position) (next Position, newSite *site.Site) {
	pos := curpos
	// seek (to the right) the site ref
	for s.synoSheet.Cell(pos.row, pos.col).Value == "" && pos.col < s.nbCols {
		pos.col++
	}
	// seek (to the right) the site info
	spos := Position{pos.row, pos.col + 1}
	for s.synoSheet.Cell(spos.row, spos.col).Value == "" && spos.col < s.nbCols {
		spos.col++
	}
	newSite = &site.Site{
		Type:      s.synoSheet.Cell(pos.row-3, spos.col).Value,
		Id:        s.synoSheet.Cell(pos.row-2, spos.col).Value,
		BPEType:   s.synoSheet.Cell(pos.row-1, spos.col).Value,
		Operation: s.synoSheet.Cell(pos.row, spos.col).Value,
		Ref:       s.synoSheet.Cell(pos.row+1, spos.col).Value,
		Ref2:      s.synoSheet.Cell(pos.row+2, spos.col).Value,
		FiberOut:  s.synoSheet.Cell(pos.row+3, spos.col).Value,
		FiberIn:   s.synoSheet.Cell(pos.row+1, pos.col).Value,
		Lenght:    s.synoSheet.Cell(pos.row+2, pos.col).Value,
		Color:     s.synoSheet.Cell(pos.row-3, spos.col).GetStyle().Fill.FgColor,
		Parent:    nil,
		Children:  []*site.Site{},
	}

	fmt.Printf("found %s %s\n", newSite.Id, newSite.Type)
	//TODO Seek for children
	next = Position{curpos.row + 8, curpos.col}
	return
}

func siteCoords(p Position) Position {
	return Position{}
}

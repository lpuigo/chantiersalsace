package suivi

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/tealeg/xlsx"
)

type BlockParser interface {
	Activity() string
	ParseBlock(sh *xlsx.Sheet, catalog *bpu.Catalog, row int) (items []*bpu.Item, err ParsingError, nextRow int)
}

func ParseTab(sh *xlsx.Sheet, blockParser BlockParser, s *Suivi) (err ParsingError) {
	// Check if Catalog has related Activity Chapters
	if s.Catalog.GetCategoryChapters(blockParser.Activity()) == nil {
		err.Add(fmt.Errorf("no %s activity declared in BPU catalog. Skipping related Items.", blockParser.Activity()), true)
		return
	}

	// Parse Tab
	row := 1
	//var bpe *Bpe
	for {
		nItems, nParsingError, nRow := blockParser.ParseBlock(sh, s.Catalog, row)
		if nParsingError.HasError() {
			err.Append(nParsingError)
		}
		s.Add(nItems...)
		if nRow == 0 {
			break
		}
		row = nRow
	}
	return
}

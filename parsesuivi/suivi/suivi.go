package suivi

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/tealeg/xlsx"
)

type Suivi struct {
	Bpes []*Bpe
}

func (s *Suivi) String() string {
	res := ""
	for id, bpe := range s.Bpes {
		res += fmt.Sprintf("%3d:%s\n", id, bpe.String())
	}
	return res
}

const (
	colBpeName   int = 1
	colBpeSize   int = 6
	colBpeOpe    int = 7
	colBpeFiber  int = 8
	colBpeSplice int = 9
	colBpeStatus int = 10
	colBpeDate   int = 14
)

func NewSuiviFromXLS(file string, pc *bpu.Bpu) (s *Suivi, err error) {
	xf, err := xlsx.OpenFile(file)
	if err != nil {
		return
	}
	if len(xf.Sheets) < 2 {
		err = fmt.Errorf("could not find sheet with Bpe details in '%s'", file)
		return
	}
	bsh := xf.Sheets[1]

	s = &Suivi{}

	row := 0
	inProgress := true
	var nbFiber, nbSplice, bpeRow int
	var bpe *Bpe
	for inProgress {
		row++
		bpeName := bsh.Cell(row, colBpeName).Value
		bpeOpe := bsh.Cell(row, colBpeOpe).Value
		if bpeOpe == "" && bpeName == "" {
			inProgress = false
			continue
		}
		if bpeName != "" { // This Row contains BPE definition, parse and process it
			if bpe != nil {
				if !bpe.CheckFiber(nbFiber) {
					err = fmt.Errorf("invalid Nb Fiber for Bpe on line %d", bpeRow+1)
					return
				}
				if !bpe.CheckSplice(nbSplice) {
					err = fmt.Errorf("invalid Nb Splice for Bpe on line %d", bpeRow+1)
					return
				}
			}
			bpeRow = row
			bpe, err = NewBpeFromXLSRow(bsh, row)
			if err != nil {
				return
			}
			bpe.SetValues(pc)
			s.Bpes = append(s.Bpes, bpe)
			nbFiber = 0
			nbSplice = 0
			continue
		}
		// this row contains BPE detail, check for fiber and splice number

		sf := bsh.Cell(row, colBpeFiber).Value
		if sf != "" {
			nbf, e := bsh.Cell(row, colBpeFiber).Int()
			if e != nil {
				err = fmt.Errorf("could not parse Nb Fiber from '%s' line %d", sf, row+1)
				return
			}
			nbFiber += nbf
		}
		ss := bsh.Cell(row, colBpeSplice).Value != ""
		if ss {
			nbs, e := bsh.Cell(row, colBpeSplice).Int()
			if e != nil {
				err = fmt.Errorf("could not parse Nb Splice from '%s' line %d", ss, row+1)
				return
			}
			nbSplice += nbs
		}
	}

	return
}

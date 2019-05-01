package suivi

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/xls"
	"github.com/tealeg/xlsx"
	"strings"
)

func parseStatus(sh *xlsx.Sheet, row, col int) (todo, done bool, err error) {
	isDone := sh.Cell(row, col).Value
	switch strings.ToLower(isDone) {
	case "ok":
		done = true
		todo = true
	case "na", "annule", "supprime", "suprime":
		todo = false
	case "", "nok", "ko", "blocage", "en cours":
		todo = true
	default:
		err = fmt.Errorf(
			"unknown Status '%s' in cell %s!%s",
			isDone,
			sh.Name,
			xls.RcToAxis(row, col),
		)
	}
	return
}

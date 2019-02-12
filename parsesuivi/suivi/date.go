package suivi

import "time"

func GetMonday(d time.Time) time.Time {
	wd := int(d.Weekday())
	if wd == 0 {
		wd = 7
	}
	wd--
	return d.AddDate(0, 0, -wd)
}

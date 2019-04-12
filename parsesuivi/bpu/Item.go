package bpu

import (
	"fmt"
	"time"
)

type Item struct {
	Activity string // Racco, Tirage, ...
	Name     string // PTxxx, Cablezzz, ...
	Info     string // BoxType + nbFO ...
	Date     time.Time
	Chapter  *Chapter
	Quantity int
	Todo     bool
	Done     bool
}

func NewItem(activity, name, info string, date time.Time, chapter *Chapter, quantity int, todo, done bool) *Item {
	return &Item{
		Activity: activity,
		Name:     name,
		Info:     info,
		Date:     date,
		Chapter:  chapter,
		Quantity: quantity,
		Todo:     todo,
		Done:     done,
	}
}

func (i *Item) String() string {
	return fmt.Sprintf(`Activity: %s Name: %s
	Info: %s
	Date: %s
	Chapter: %s
	Quantity: %d
	Todo: %t
	Done: %t
`, i.Activity, i.Name, i.Info, i.Date.Format("2006-01-02"), i.Chapter.Name, i.Quantity, i.Todo, i.Done)
}

// Price returns the price for the given item
func (i *Item) Price() float64 {
	return i.Chapter.Price * float64(i.Quantity)
}

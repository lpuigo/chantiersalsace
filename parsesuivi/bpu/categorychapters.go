package bpu

import (
	"fmt"
	"sort"
	"strings"
)

type CategoryChapters map[string][]*Article //map[Category][]*Article

func NewCategoryChapters() CategoryChapters {
	return make(CategoryChapters)
}

// SortChapters sort CategoryChapters by ascending size order
func (cc CategoryChapters) SortChapters() {
	for cat, chapters := range cc {
		sort.Slice(chapters, func(i, j int) bool {
			return chapters[i].Size < chapters[j].Size
		})
		cc[cat] = chapters
	}
}

func (cc CategoryChapters) GetChapterForSize(category string, size int) (*Article, error) {
	chapters := cc[strings.ToUpper(category)]
	if len(chapters) == 0 {
		return nil, fmt.Errorf("unknown category '%s'", category)
	}
	for _, p := range chapters {
		if size <= p.Size {
			return p, nil
		}
	}
	return chapters[len(chapters)-1], nil
}

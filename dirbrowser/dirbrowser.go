package dirbrowser

import (
	"os"
	"path/filepath"
	"strings"
)

type ProcessFileFunc func(path string) error

// Process walks trhough path directory,
// and apply fn function on all found files having fileExt extension
func Process(path, fileExt string, fn ProcessFileFunc) error {
	err := filepath.Walk(path, processFn(fileExt, fn))
	if err != nil {
		return err
	}
	return nil
}

func processFn(ext string, pfunc ProcessFileFunc) filepath.WalkFunc {
	return func(path string, info os.FileInfo, err error) error {
		if err != nil {
			// silently skip file with error
			return nil
		}
		if info.IsDir() {
			// skip directory
			return nil
		}
		name := info.Name()
		fileExt := strings.ToLower(filepath.Ext(name))
		if fileExt != ext {
			return nil
		}
		return pfunc(path)
	}
}

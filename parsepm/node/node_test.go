package node

import (
	"path/filepath"
	"testing"
)

const (
	testDir     string = `C:\Users\Laurent\Desktop\CCPE_DES_PM3_BPE`
	testXLS     string = ``
	blobpattern string = `*.xlsx`
)

func TestNode_ParseXLS(t *testing.T) {
	parseBlobPattern := filepath.Join(testDir, blobpattern)
	files, err := filepath.Glob(parseBlobPattern)
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range files {
		n := NewNode()
		err := n.ParseBPEXLS(f)
		if err != nil {
			t.Errorf("'%s' returned unexpected : %s\n", filepath.Base(f), err.Error())
		}
	}
}

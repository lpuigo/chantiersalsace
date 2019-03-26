package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/dirbrowser"
	"github.com/lpuig/ewin/chantiersalsace/xlsxconvert"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

const (
	blobpattern string = `*.xls`
)

//go:generate bash build_debug.sh
//go:generate bash build.sh

func main() {
	gc := GuiContext{}
	err := MainWindow{
		AssignTo: &gc.MainWindow,
		Title:    "EWIN Services XLS to XLSx",
		MinSize:  Size{640, 480},
		OnDropFiles: func(files []string) {
			go gc.GoConvert(files)
		},
		Layout: VBox{},
		Children: []Widget{
			//Label{Text: "Glisser des fichiers XLS ici ..."},
			CheckBox{
				Alignment:           AlignHNearVCenter,
				AssignTo:            &gc.eraseFileCB,
				Text:                "Supprimer les fichiers XLS originaux après conversion",
				OnCheckStateChanged: func() { gc.SwitchEraseState() },
			},
			TextEdit{
				Text:      "Glisser des fichiers XLS ici ...\r\n",
				AssignTo:  &gc.textEdit,
				ReadOnly:  true,
				VScroll:   true,
				MaxLength: 100 * 1024,
			},
		},
	}.Create()
	if err != nil {
		log.Fatal(err)
	}
	gc.eraseFileCB.SetChecked(true)
	gc.Run()
}

type GuiContext struct {
	*walk.MainWindow
	textEdit    *walk.TextEdit
	eraseFileCB *walk.CheckBox
	msg         chan string
}

func (gc *GuiContext) SwitchEraseState() {
	gc.textEdit.AppendText(fmt.Sprintf("Effacer les fichiers originaux : %t\r\n", gc.eraseFileCB.Checked()))
}

func (gc *GuiContext) GoConvert(files []string) {
	// check if another goroutine is running
	if gc.msg != nil {
		return
	}
	wg := sync.WaitGroup{}
	gc.msg = make(chan string)
	wg.Add(1)
	go func() {
		for msg := range gc.msg {
			gc.textEdit.AppendText(msg)
		}
		wg.Done()
	}()

	gc.Convert(files)
	close(gc.msg)
	wg.Wait()
	gc.msg = nil
}

func (gc GuiContext) Logln(text string) {
	gc.msg <- text + "\r\n"
}

func (gc GuiContext) Logf(format string, arg ...interface{}) {
	gc.msg <- fmt.Sprintf(format, arg...)
}

func (gc *GuiContext) Convert(filenames []string) {
	for _, filename := range filenames {
		fi, err := os.Stat(filename)
		if err != nil {
			gc.Logf("Error : %v\r\n", err)
			continue
		}
		if fi.IsDir() {
			gc.ConvertDirToXlsx(filename)
			continue
		}
		if !strings.HasSuffix(strings.ToLower(filename), ".xls") {
			gc.Logf("Ignoring file : %s\r\n", filename)
			continue
		}
		gc.ConvertToXlsx(filename)
	}
	gc.Logln("Done")
}

func (gc *GuiContext) ConvertDirToXlsx(dir string) {
	gc.Logf("Processing dir %s :\r\n", dir)
	//parseBlobPattern := filepath.Join(dir, blobpattern)
	//files, err := filepath.Glob(parseBlobPattern)
	//if err != nil {
	//	gc.Logf("\tfailed : %s\r\n", err.Error())
	//	return
	//}
	//for _, f := range files {
	//	gc.ConvertToXlsx(f)
	//}
	fn := func(file string) error {
		gc.ConvertToXlsx(file)
		return nil
	}
	err := dirbrowser.Process(dir, ".xls", fn)
	if err != nil {
		gc.Logf("\tError :%s\r\n", err.Error())
	}
}

func (gc *GuiContext) ConvertToXlsx(filename string) {
	gc.Logf("Converting file %s :\r\n", filepath.Base(filename))
	errs := xlsxconvert.OleXlsToXlsx(filename)
	if len(errs) > 1 && errs[0] != nil {
		gc.Logf("\tfailed : %s\r\n", errs[0].Error())
		return
	}
	if !gc.eraseFileCB.Checked() {
		return
	}
	if err := os.Remove(filename); err != nil {
		gc.Logf("\tcould not delete file '%s' : %v\r\n", filename, err.Error())
	}
}

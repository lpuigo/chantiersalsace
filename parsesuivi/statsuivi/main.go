package main

import (
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/bpu"
	"github.com/lpuig/ewin/chantiersalsace/parsesuivi/suivi"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
	"log"
	"path/filepath"
	"strings"
	"sync"
)

//go_:generate bash build_debug.sh
//go:generate bash build.sh

const (
	bpuFile                 string = `BPU.xlsx`
	suiviOutFileSuffix      string = `_Suivi.xlsx`
	attachmentOutFileSuffix string = `_Attachement.xlsx`
	separator               string = "_suivi_"
)

func main() {
	gc := GuiContext{}
	err := MainWindow{
		AssignTo: &gc.MainWindow,
		Title:    "EWIN Services Mise à jour du Suivi Chantier",
		MinSize:  Size{800, 480},
		OnDropFiles: func(files []string) {
			gc.SetAndProcess(files[0])
		},
		Layout: VBox{},
		Children: []Widget{
			Composite{
				Layout: Grid{Columns: 10, MarginsZero: true},
				Children: []Widget{
					PushButton{
						Text:               "Choisir...",
						AlwaysConsumeSpace: true,
						OnClicked: func() {
							gc.BrowseXLS()
						},
					},
					Label{
						Text:       "",
						AssignTo:   &gc.suiviLbl,
						ColumnSpan: 8,
					},
					PushButton{
						Text:               "Traitement",
						AssignTo:           &gc.processPB,
						AlwaysConsumeSpace: true,
						Enabled:            false,
						OnClicked: func() {
							gc.Process()
						},
					},
				},
			},
			TextEdit{
				Text:      "Glisser un fichier de suivi XLS ici ...\r\n",
				AssignTo:  &gc.msgTE,
				ReadOnly:  true,
				VScroll:   true,
				MaxLength: 100 * 1024,
			},
		},
	}.Create()
	if err != nil {
		log.Fatal(err)
	}
	// Init actions here

	// Show mainwindow
	gc.Run()
}

type GuiContext struct {
	*walk.MainWindow
	msgTE *walk.TextEdit
	msg   chan string

	suiviLbl  *walk.Label
	processPB *walk.PushButton
}

func (gc *GuiContext) GoProcess(process func()) {
	// exit if another process is running
	if gc.msg != nil {
		return
	}
	wg := sync.WaitGroup{}
	gc.msg = make(chan string)
	wg.Add(1)
	go func() {
		for msg := range gc.msg {
			gc.msgTE.AppendText(msg)
		}
		wg.Done()
	}()

	process()

	close(gc.msg)
	wg.Wait()
	gc.msg = nil
}

func (gc GuiContext) AddMsgLn(msg string) {
	gc.msgTE.AppendText(msg + "\r\n")
}

func (gc GuiContext) Logln(text string) {
	gc.msg <- text + "\r\n"
}

func (gc GuiContext) Logf(format string, arg ...interface{}) {
	gc.msg <- fmt.Sprintf(format, arg...)
}

func (gc GuiContext) BrowseXLS() {
	dlg := new(walk.FileDialog)
	if gc.suiviLbl.Text() != "" {
		dlg.FilePath = filepath.Dir(gc.suiviLbl.Text())
	}
	dlg.Filter = "XLSx Files (*.xslx)"
	dlg.Title = "Choisir un fichier de suivi"

	if ok, err := dlg.ShowOpen(gc.MainWindow); err != nil {
		gc.AddMsgLn(err.Error())
	} else if !ok {
		return
	}

	gc.SetAndProcess(dlg.FilePath)
}

func (gc *GuiContext) SetSuiviFile(file string) error {
	var err error
	baseFile := filepath.Base(file)
	if !strings.Contains(strings.ToLower(baseFile), separator) {
		err = fmt.Errorf("'%s' n'est pas un fichier XLSX de suivi (le nom du fichier doit contenir %s)", baseFile, separator)
	}
	if !strings.HasSuffix(baseFile, ".xlsx") {
		err = fmt.Errorf("'%s' n'est pas un fichier XLSX", file)
	}

	if err != nil {
		gc.suiviLbl.SetText("")
		gc.processPB.SetEnabled(false)
		return err
	}
	gc.suiviLbl.SetText(file)
	gc.processPB.SetEnabled(true)
	return nil
}

func (gc *GuiContext) SetAndProcess(file string) {
	err := gc.SetSuiviFile(file)
	if err != nil {
		gc.AddMsgLn(fmt.Sprintf("Erreur : %s", err.Error()))
		return
	}
	gc.Process()
}

func (gc *GuiContext) Process() {
	file := gc.suiviLbl.Text()
	print("Traitement du fichier " + file)
	process := func() {
		gc.processPB.SetEnabled(false)
		gc.AddMsgLn("Traitement du fichier " + file)
		gc.ProcessSuivi()
		gc.AddMsgLn("Traitement terminé")
		gc.processPB.SetEnabled(true)
	}
	go gc.GoProcess(process)
}

func (gc *GuiContext) ProcessSuivi() {
	file := gc.suiviLbl.Text()
	dir := filepath.Dir(file)

	fpart := strings.Split(strings.ToLower(filepath.Base(file)), separator)
	prefix := strings.ToUpper(fpart[0])

	bpuFile := filepath.Join(dir, bpuFile)
	suiviOutFile := filepath.Join(dir, prefix+suiviOutFileSuffix)
	//attachementOutFile := filepath.Join(dir, prefix+attachmentOutFileSuffix)

	currentBpu, err := bpu.NewCatalogFromXLS(bpuFile)
	if err != nil {
		gc.Logf("Erreur : impossible de traiter le fichier '%'\r\n\t%s\r\n", bpuFile, err.Error())
		return
	}

	progress, err := suivi.NewSuiviFromXLS(file, currentBpu)
	if err != nil {
		gc.Logf("Erreur lors du traitement du fichier de suivi\r\n%s", err.Error())
		return
	}

	err = progress.WriteSuiviXLS(suiviOutFile)
	if err != nil {
		gc.Logf("Erreur : impossible de mettre à jour le fichier de suivi '%s'\r\n\t%s\r\n", suiviOutFile, err.Error())
		return
	}

	//err = progress.WriteAttachmentXLS(attachementOutFile)
	//if err != nil {
	//	gc.Logf("Erreur : impossible de mettre à jour le fichier d'attachement '%s'\r\n\t%s\r\n", attachementOutFile, err.Error())
	//}
}

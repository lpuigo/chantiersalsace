package zone

import (
	"encoding/json"
	"fmt"
	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/lpuig/ewin/doe/website/backend/model/ripsites"
	"github.com/lpuig/ewin/doe/website/frontend/model/ripsite/ripconst"
	"github.com/tealeg/xlsx"
	"os"
	"path/filepath"
	"strings"
)

type Zone struct {
	Nodes     node.Nodes
	Troncons  node.Troncons
	Cables    node.Cables
	Sro       *node.Node
	NodeRoots []*node.Node
}

func New() *Zone {
	z := &Zone{
		Nodes:     node.NewNodes(),
		Troncons:  node.NewTroncons(),
		NodeRoots: []*node.Node{},
		Sro:       node.NewNode(),
	}
	z.Sro.Name = "SRO"
	z.Sro.PtName = "SRO"
	z.Sro.BPEType = "SRO"
	z.Sro.TronconIn = node.NewTroncon("Aduction")
	return z
}

const (
	blobpattern string = `*PT*.xlsx`
)

func (z *Zone) ParseBPEDir(dir string) error {
	fs, err := os.Stat(dir)
	if err != nil {
		return err
	}
	if !fs.IsDir() {
		return fmt.Errorf("'%s' is not a directory", dir)
	}
	parseBlobPattern := filepath.Join(dir, blobpattern)
	files, err := filepath.Glob(parseBlobPattern)
	if err != nil {
		return err
	}
	for _, f := range files {
		// skip XLS temp files
		if strings.HasPrefix(filepath.Base(f), "~") {
			continue
		}
		n := node.NewNode()
		err := n.ParseBPEXLS(f, z.Troncons)
		if err != nil {
			return fmt.Errorf("parsing '%s' returned error : %s\n", filepath.Base(f), err.Error())
		}
		fmt.Printf("'%s' parsed\n", n.PtName)
		newNode := z.Nodes.Add(n)
		if !newNode {
			return fmt.Errorf("node %s was already defined", n.PtName)
		}
	}

	return nil
}

func (z *Zone) WriteXLS(dir, name string) error {
	if len(z.Nodes) == 0 {
		return fmt.Errorf("zone is empty, nothing to write to XLSx")
	}
	file := filepath.Join(dir, name+"_suivi.xlsx")

	xlsx.SetDefaultFont(11, "Calibri")
	xls := xlsx.NewFile()

	if len(z.Cables) > 0 && z.Cables[0].Troncons[0].CableType != "" {
		err := z.addTirageSheet(xls)
		if err != nil {
			return fmt.Errorf("Tirage : %s", err.Error())
		}
	}

	err := z.addRaccoSheet(xls)
	if err != nil {
		return fmt.Errorf("Racco : %s", err.Error())
	}

	err = z.addMesuresSheet(xls)
	if err != nil {
		return fmt.Errorf("Mesures : %s", err.Error())
	}

	of, err := os.Create(file)
	if err != nil {
		return err
	}
	defer of.Close()

	return xls.Write(of)
}

func (z *Zone) addTirageSheet(xls *xlsx.File) error {
	sheet, err := xls.AddSheet("Tirage")
	if err != nil {
		return err
	}

	z.Cables[0].WriteTirageHeader(sheet)

	for _, cable := range z.Cables {
		cable.WriteTirageXLS(sheet)
	}
	return nil
}

func (z *Zone) addRaccoSheet(xls *xlsx.File) error {
	sheet, err := xls.AddSheet("Racco")
	if err != nil {
		return err
	}

	node.NewNode().WriteRaccoHeader(sheet)

	if len(z.Sro.Children) > 0 {
		z.Sro.WriteRaccoXLS(sheet)
	} else {
		for _, rootnode := range z.NodeRoots {
			rootnode.WriteRaccoXLS(sheet)
		}
	}
	return nil
}

func (z *Zone) addMesuresSheet(xls *xlsx.File) error {
	sheet, err := xls.AddSheet("Mesures")
	if err != nil {
		return err
	}

	node.NewNode().WriteMesuresHeader(sheet)
	z.Sro.WriteMesuresXLS(sheet, z.Nodes)
	return nil
}

func (z *Zone) ParseROPXLS(file string) error {
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}

	var sheet *xlsx.Sheet
	for _, sh := range xls.Sheets {
		if strings.HasPrefix(sh.Name, "TAB") {
			sheet = sh
			break
		}
	}
	if sheet == nil {
		return fmt.Errorf("could not find Tab sheet")
	}

	//parse sheet
	rp := NewRopParser(sheet, z)

	rp.ParseRop()

	return nil
}

func (z *Zone) CreateBPETree() {
	// Create Troncon list with source / dest node
	cables := map[string]Link{}
	for _, nod := range z.Nodes {
		if nod.TronconIn != nil && nod.TronconIn.Name != "" {
			link := cables[nod.TronconIn.Name]
			link.Dest = nod
			cables[nod.TronconIn.Name] = link
		}
		for cableName, cable := range nod.TronconsOut {
			if nod.TronconIn != nil && nod.TronconIn.Name == cable.Name {
				continue
			}
			link := cables[cableName]
			link.Source = nod
			cables[cableName] = link
		}
	}
	// Populate Nodes Children
	for cableName, link := range cables {
		if link.Source != nil {
			if link.Dest != nil {
				link.Source.Children = append(link.Source.Children, link.Dest)
				link.Dest.IsChild = true
			} else {
				link.Source.AddPMChild(link.Source.TronconsOut[cableName])
			}
		}
	}
	// Detect and attach root nodes to new PM (TODO : detect PM using PM Splice file)
	for _, nod := range z.Nodes {
		if nod.IsChild {
			continue
		}
		if nod.TronconIn != nil && nod.TronconIn.Name != "" {
			z.NodeRoots = append(z.NodeRoots, node.NewPMNode(nod))
		} else {
			z.NodeRoots = append(z.NodeRoots, nod)
		}
	}
}

func (z *Zone) DetectCables(node *node.Node) {
	for _, tr := range node.SpliceTRs() {
		z.AddNewCable(tr)
	}
}

func (z *Zone) AddNewCable(tr *node.Troncon) {
	nc := node.NewCable(tr)
	z.Cables.Add(nc)
	for tr != nil {
		nc.AddTroncon(tr, 20)

		// detect new cable starting in dest node
		nextNode := tr.NodeDest
		z.DetectCables(nextNode)

		// check for passage
		tr = nextNode.GetTronconPassage()
	}
}

const (
	rowQCStart      int = 5
	colQCTroncon    int = 1
	colQCCableType  int = 2
	colQCLoveLength int = 4
	colQCCapa       int = 5
	colQCLength     int = 10
	colQCTirageType int = 12
)

func (z *Zone) ParseQuantiteCableXLS(file string) error {
	baseFile := filepath.Base(file)
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}
	sheet := xls.Sheets[0]
	if sheet.Cell(rowQCStart-2, colQCTroncon).Value != "etiquette" {
		return fmt.Errorf("could not find 'etiquette' label on line %d, col %d", rowQCStart-1, colQCTroncon+1)
	}
	for row := rowQCStart; row < sheet.MaxRow; row++ {
		trName := sheet.Cell(row, colQCTroncon).Value
		if trName == "" {
			continue
		}
		tr := z.Troncons[trName]
		if tr == nil {
			return fmt.Errorf("unknown Troncon '%s' found on line %d", trName, row+1)
		}
		tr.CableType = sheet.Cell(row, colQCCableType).Value
		tr.LoveLength, err = sheet.Cell(row, colQCLoveLength).Int()
		if err != nil {
			fmt.Printf("\t%s: could not read Love length '%s' on line %d, col %d (use default 20m instead)\n", baseFile, sheet.Cell(row, colQCLoveLength).Value, row+1, colQCLoveLength+1)
			tr.LoveLength = 20
		}
		// Capa to be check with already existing value
		//tr.Capa, err = sheet.Cell(row, colQCCapa).Int()
		tirageType := strings.ToUpper(sheet.Cell(row, colQCTirageType).Value)
		tirageLength, err := sheet.Cell(row, colQCLength).Int()
		switch {
		case err == nil && strings.Contains(tirageType, "AERIEN"):
			tr.AerialLength += tirageLength
		case err == nil && strings.Contains(tirageType, "FACADE"):
			tr.FacadeLength += tirageLength
		case err == nil && strings.Contains(tirageType, "INFRA"):
			tr.UndergroundLength += tirageLength
		default:
			if tirageType == "" {
				continue
			}
			if err == nil && tirageType != "" {
				fmt.Printf("\t%s: Unknown tirage type '%s' on line %d, col %d\n", baseFile, tirageType, row+1, colQCTirageType+1)
				continue
			}
			if err != nil {
				return fmt.Errorf("could not read tirage length '%s' on line %d, col %d", sheet.Cell(row, colQCLength).Value, row+1, colQCLoveLength+1)
			}
		}
	}
	return nil
}

func (z *Zone) WriteJSON(dir, name string) error {
	if len(z.Nodes) == 0 {
		return fmt.Errorf("zone is empty, nothing to write to Json")
	}
	file := filepath.Join(dir, name+".json")

	site := ripsites.Site{
		Id:           0,
		Client:       "UNDEFINED",
		Ref:          name,
		Manager:      "UNDEFINED",
		OrderDate:    "YYYY-MM-DD",
		Status:       ripconst.RsStatusInProgress,
		Comment:      "TO BE DEFINED",
		Nodes:        make(map[string]*ripsites.Node),
		Troncons:     make(map[string]*ripsites.Troncon),
		Pullings:     nil,
		Junctions:    nil,
		Measurements: nil,
	}

	z.addSiteNodes(&site)
	z.addSiteTroncon(&site)

	z.addSitePullings(&site)
	z.addSiteJunctions(&site)
	z.addSiteMeasurements(&site)

	f, err := os.Create(file)
	if err != nil {
		return fmt.Errorf("could not create file:%s\n", err.Error())
	}
	return json.NewEncoder(f).Encode(&site)
}

func (z *Zone) addSiteNodes(site *ripsites.Site) {
	for _, node := range z.Nodes {
		siteNode := &ripsites.Node{
			Name:          node.PtName,
			Address:       node.Address,
			Type:          node.LocationType,
			BoxType:       node.BPEType,
			Ref:           node.Name,
			TronconInName: node.TronconIn.Name,
			DistFromPm:    node.DistFromPM,
		}
		site.Nodes[siteNode.Name] = siteNode
	}
}

func (z *Zone) addSiteTroncon(site *ripsites.Site) {
	for _, troncon := range z.Troncons {
		siteTroncon := &ripsites.Troncon{
			Name: troncon.Name,
			Size: troncon.Capa,
		}
		site.Troncons[siteTroncon.Name] = siteTroncon
	}
}

func (z *Zone) addSitePullings(site *ripsites.Site) {
	site.Pullings = []*ripsites.Pulling{}
	state := ripsites.State{
		Status: ripconst.StateToDo,
	}

	for _, cable := range z.Cables {
		if cable.Troncons[0].CableType == "" {
			continue
		}
		sitePulling := &ripsites.Pulling{
			CableName: cable.Troncons[0].CableType,
			Chuncks:   nil,
			State:     state,
		}
		for _, tr := range cable.Troncons {
			chunck := ripsites.PullingChunk{
				TronconName:      tr.Name,
				StartingNodeName: tr.NodeSource.PtName,
				EndingNodeName:   tr.NodeDest.PtName,
				LoveDist:         tr.LoveLength,
				UndergroundDist:  tr.UndergroundLength,
				AerialDist:       tr.AerialLength,
				BuildingDist:     tr.FacadeLength,
				State:            state,
			}
			if chunck.LoveDist+chunck.UndergroundDist+chunck.AerialDist+chunck.BuildingDist == 0 {
				chunck.LoveDist = 20
				chunck.UndergroundDist = tr.NodeDest.DistFromPM - tr.NodeSource.DistFromPM
			}
			sitePulling.Chuncks = append(sitePulling.Chuncks, chunck)
		}
		site.Pullings = append(site.Pullings, sitePulling)
	}
}

func (z *Zone) addSiteJunctions(site *ripsites.Site) {
	if len(z.Sro.Children) > 0 {
		addJunction(z.Sro, site)
	} else {
		for _, rootnode := range z.NodeRoots {
			addJunction(rootnode, site)
		}
	}
}

func addJunction(n *node.Node, site *ripsites.Site) {
	state := ripsites.State{
		Status: ripconst.StateToDo,
	}
	junction := &ripsites.Junction{
		NodeName:   n.PtName,
		Operations: nil,
		State:      state,
	}

	for _, opname := range n.Operations() {
		opeType, trName := "", ""
		lOpName := strings.ToLower(opname)
		switch {
		case strings.HasPrefix(lOpName, "att"):
			opeType = "Attente"
			trName = ""
		case strings.HasPrefix(lOpName, "epi"):
			opeType = "Epissure"
			if strings.Contains(opname, "->") {
				trName = strings.Split(opname, "->")[1]
			}
		case strings.HasPrefix(lOpName, "pas"):
			opeType = "Passage"
			if strings.Contains(opname, "->") {
				trName = strings.Split(opname, "->")[1]
			}
		}
		e, o := n.GetOperationNumbers(opname)
		operation := ripsites.Operation{
			Type:        opeType,
			TronconName: trName,
			NbFiber:     o + e,
			NbSplice:    e,
			State:       state,
		}
		junction.Operations = append(junction.Operations, operation)
	}

	site.Junctions = append(site.Junctions, junction)

	for _, cnode := range n.GetChildren() {
		addJunction(cnode, site)
	}
}

func (z *Zone) addSiteMeasurements(site *ripsites.Site) {
	z.addMeasurement(z.Sro, site)
}

func (z *Zone) addMeasurement(n *node.Node, site *ripsites.Site) {
	wf := n.GetToBeMeasuredFiber()
	if wf > 0 {
		state := ripsites.State{
			Status: ripconst.StateToDo,
		}
		measurement := &ripsites.Measurement{
			DestNodeName: n.PtName,
			NbFiber:      wf,
			Dist:         n.DistFromPM,
			NodeNames:    n.SplicePT,
			State:        state,
		}
		site.Measurements = append(site.Measurements, measurement)
	}

	for _, cnode := range n.GetChildren() {
		z.addMeasurement(cnode, site)
	}
}

func (z *Zone) EnableCables(enableDestBPECable map[string]string) {
	for _, cable := range z.Cables {
		destBPEType := cable.Troncons[len(cable.Troncons)-1].NodeDest.BPEType
		cableType := enableDestBPECable[destBPEType]
		if cableType == "" {
			continue
		}
		cable.Troncons[0].CableType = fmt.Sprintf(cableType, cable.Troncons[0].Capa)
	}
}

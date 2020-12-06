package zone

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/lpuig/ewin/chantiersalsace/parsepm/node"
	"github.com/lpuig/ewin/doe/website/backend/model/date"
	"github.com/lpuig/ewin/doe/website/backend/model/ripsites"
	"github.com/lpuig/ewin/doe/website/frontend/model/ripsite/ripconst"
	"github.com/tealeg/xlsx"
)

type Zone struct {
	Nodes               node.Nodes
	Troncons            node.Troncons
	Cables              node.Cables
	Sro                 *node.Node
	NodeRoots           []*node.Node
	DoPulling           bool
	DoJunctions         bool
	DoEline             bool
	DoOtherThanEline    bool
	DoMeasurement       bool
	CreateNodeFromRop   bool
	DefineNodeOperation map[string]bool
	BlobPattern         string
}

func New() *Zone {
	z := &Zone{
		Nodes:               node.NewNodes(),
		Troncons:            node.NewTroncons(),
		NodeRoots:           []*node.Node{},
		Sro:                 node.NewNode(),
		CreateNodeFromRop:   true,
		DefineNodeOperation: make(map[string]bool),
		BlobPattern:         blobpattern_EasyFibre,
	}
	z.Sro.Name = "SRO"
	z.Sro.PtName = "SRO"
	z.Sro.BPEType = "SRO"
	z.Sro.TronconIn = node.NewTroncon("Aduction")
	return z
}

const (
	blobpattern_EasyFibre string = `*PT*.xlsx`
	Blobpattern_Sogetrel  string = `*/_*.xlsx`
)

func (z *Zone) ParseBPEDir(dir string) error {
	fs, err := os.Stat(dir)
	if err != nil {
		return err
	}
	if !fs.IsDir() {
		return fmt.Errorf("'%s' is not a directory", dir)
	}
	parseBlobPattern := filepath.Join(dir, z.BlobPattern)
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
			fmt.Printf("skip unknown Troncon '%s' found on line %d", trName, row+1)
			continue
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

const (
	rowQC2Start       int = 1
	colQC2PullingType int = 1
	colQC2Length      int = 2
	colQC2Capa        int = 3
	colQC2Orig        int = 4
	colQC2Dest        int = 5
)

func (z *Zone) ParseQuantiteCableOptiqueC2Xlsx(file string) error {
	baseFile := filepath.Base(file)
	xls, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}
	sheet := xls.Sheets[0]
	if sheet.Cell(rowQC2Start-1, colQC2Orig).Value != "ORIGINE" {
		return fmt.Errorf("could not find 'ORIGINE' label on cell '%s'", xlsx.GetCellIDStringFromCoords(colQC2Orig, rowQC2Start))
	}

	type c2Data struct {
		orig        string
		dest        string
		pullingType string
		lenght      float64
		capa        int
	}

	c2DataDict := make(map[string]c2Data)

	for row := rowQC2Start; row < sheet.MaxRow; row++ {
		length, err := sheet.Cell(row, colQC2Length).Float()
		if err != nil {
			fmt.Printf("\t%s: could not read length '%s' on line %d, col %d (use default 20m instead)\n", baseFile, sheet.Cell(row, colQC2Length).Value, row+1, colQC2Length+1)
			length = 20
		}
		capa, err := sheet.Cell(row, colQC2Capa).Int()
		if err != nil {
			fmt.Printf("\t%s: could not read capa '%s' on line %d, col %d (use default 0 instead)\n", baseFile, sheet.Cell(row, colQC2Capa).Value, row+1, colQC2Length+1)
			capa = 0
		}
		c := c2Data{
			orig:        sheet.Cell(row, colQC2Orig).Value,
			dest:        sheet.Cell(row, colQC2Dest).Value,
			pullingType: sheet.Cell(row, colQC2PullingType).Value,
			lenght:      length,
			capa:        capa,
		}
		c2DataDict[c.dest] = c
	}

	for _, tr := range z.Troncons {
		if tr.NodeDest == nil {
			fmt.Printf("\t%s has no destination node. Skipping\n", tr.Name)
			continue
		}
		c2, found := c2DataDict[tr.NodeDest.PtName]
		if !found {
			fmt.Printf("\t%s has unknown destination node %s. Skipping\n", tr.Name, tr.NodeDest.PtName)
			continue
		}
		tr.LoveLength = 0
		if tr.Capa != c2.capa {
			if z.CreateNodeFromRop {
				tr.Capa = c2.capa
			} else {
				fmt.Printf("\t%s has unexpected capacity %d vs %d\n", tr.Name, c2.capa, tr.Capa)
			}
		}
		tr.CableType = fmt.Sprintf("CABLE_%dFO", tr.Capa)
		tirageType := c2.pullingType
		pullingLength := int(c2.lenght)
		switch {
		case strings.Contains(tirageType, "AERIEN"):
			tr.AerialLength += pullingLength
		case strings.Contains(tirageType, "FACADE"):
			tr.FacadeLength += pullingLength
		case strings.Contains(tirageType, "CONDUITE"):
			tr.UndergroundLength += pullingLength
		default:
			fmt.Printf("\t%s has unknown pulling type '%s'\n", tr.Name, tirageType)
		}
	}
	return nil
}

const (
	rowQBOStart     int = 1
	colQBOName      int = 0
	colQBOReference int = 1
	colQBOType      int = 2
	colQBOFunction  int = 5
)

func (z *Zone) ParseQuantiteBoiteOptiqueD2Xlsx(file string) error {
	baseFile := filepath.Base(file)
	xlsFile, err := xlsx.OpenFile(file)
	if err != nil {
		return err
	}
	sheet := xlsFile.Sheets[0]
	if sheet.Cell(rowQBOStart-1, colQBOName).Value != "NOM" {
		return fmt.Errorf("could not find 'NOM' label on cell '%s'", xlsx.GetCellIDStringFromCoords(colQBOName, rowQBOStart-1))
	}

	type boData struct {
		name     string
		ref      string
		boxType  string
		function string
	}

	boDataDict := make(map[string]boData)

	for row := rowQBOStart; row < sheet.MaxRow; row++ {
		c := boData{
			name:     sheet.Cell(row, colQBOName).Value,
			ref:      sheet.Cell(row, colQBOReference).Value,
			boxType:  sheet.Cell(row, colQBOType).Value,
			function: sheet.Cell(row, colQBOFunction).Value,
		}
		boDataDict[c.name] = c
	}

	for _, node := range z.Nodes {
		bo, found := boDataDict[node.PtName]
		if !found {
			fmt.Printf("\t%s node is not declared in XLSx file '%s'. Skipping\n", node.PtName, baseFile)
			continue
		}
		if node.BPEType != bo.ref {
			if z.CreateNodeFromRop {
				node.BPEType = bo.ref
			} else {
				fmt.Printf("\t%s has unexpected box model '%s' instead of '%s'\n", node.PtName, node.BPEType, bo.ref)
			}
		}
		if bo.function != "PBO" {
			bo.function = "BPE"
		}
		if node.LocationType != bo.function {
			if z.CreateNodeFromRop {
				node.LocationType = bo.function
			} else {
				fmt.Printf("\t%s has unexpected usage '%s' instead of '%s'\n", node.PtName, node.LocationType, bo.function)
			}
		}
	}
	return nil
}

func (z *Zone) WriteJSON(dir, name, client, manager string, siteId int) error {
	if len(z.Nodes) == 0 {
		return fmt.Errorf("zone is empty, nothing to write to Json")
	}
	file := filepath.Join(dir, fmt.Sprintf("%06d.json", siteId))

	site := &ripsites.Site{
		Id:           siteId,
		Client:       client,
		Ref:          name,
		Manager:      manager,
		OrderDate:    date.Today().String(),
		Status:       ripconst.RsStatusInProgress,
		Comment:      "",
		Nodes:        make(map[string]*ripsites.Node),
		Troncons:     make(map[string]*ripsites.Troncon),
		Pullings:     nil,
		Junctions:    nil,
		Measurements: nil,
	}

	z.addSiteNodes(site)
	z.addSiteTroncon(site)

	z.addSitePullings(site)
	z.addSiteJunctions(site)
	z.addSiteMeasurements(site)

	f, err := os.Create(file)
	if err != nil {
		return fmt.Errorf("could not create file:%s\n", err.Error())
	}
	defer f.Close()
	return json.NewEncoder(f).Encode(site)
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
	if !z.DoPulling {
		return
	}

	state := ripsites.MakeState(ripconst.StateToDo)

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
	site.Junctions = []*ripsites.Junction{}
	if !z.DoJunctions {
		return
	}
	if len(z.Sro.Children) > 0 {
		z.addJunction(z.Sro, site)
	} else {
		for _, rootnode := range z.NodeRoots {
			z.addJunction(rootnode, site)
		}
	}
}

func (z *Zone) addJunction(n *node.Node, site *ripsites.Site) {
	state := ripsites.MakeState(ripconst.StateToDo)
	if !z.DoEline && n.BPEType == "ELINE" {
		state.Status = ripconst.StateCanceled
		state.Comment = "A ne pas faire"
	}
	if !z.DoOtherThanEline && n.BPEType != "ELINE" {
		state.Status = ripconst.StateCanceled
		state.Comment = "A ne pas faire"
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
		z.addJunction(cnode, site)
	}
}

func (z *Zone) addSiteMeasurements(site *ripsites.Site) {
	site.Measurements = []*ripsites.Measurement{}
	if !z.DoMeasurement {
		return
	}
	z.addMeasurement(z.Sro, site)
}

func (z *Zone) addMeasurement(n *node.Node, site *ripsites.Site) {
	wf := n.GetToBeMeasuredFiber()
	if wf > 0 {
		state := ripsites.MakeState(ripconst.StateToDo)

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

func (z *Zone) CheckConsistency() {
	for _, node := range z.Nodes {
		if node.BPEType == "" {
			node.BPEType = "ELINE"
		}
	}
}

func (z *Zone) GetNodeByTronconIn(troncon string) *node.Node {
	for _, node := range z.Nodes {
		if node.TronconIn.Name == troncon {
			return node
		}
	}
	return nil
}

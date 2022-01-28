package workbook

import (
    "errors"
    "fmt"
    "log"
    "sync"

    "github.com/xuri/excelize/v2"
)

var (
    ErrOpeningExcelFile       = errors.New("Error opening Excel file")
    ErrClosingExcelFile       = errors.New("Error closing Excel file")
    ErrNoHeaderFound          = errors.New("Error no valid header found")
    ErrNotEnoughHeader        = errors.New("Error not enough headers")
    ErrHeaderNotAligned       = errors.New("Error header not aligned")
    ErrUnknownCellType        = errors.New("Error unknown cell type")
    ErrSearchingHeader        = errors.New("Error searching header")
    ErrReadingCellValue       = errors.New("Error reading cell value")
    ErrReadingColumns         = errors.New("Error reading columns")
    ErrReadingRows            = errors.New("Error reading rows")
    HeaderAligned             = "Header aligned"
    HeaderNotFound            = "Header not found"
    HeaderFound               = "Header found"
    HeaderMultipleCoordinates = "Header multiple coordinates found"
    ValueParsedVerified       = "Value parsed/verified"
)

type Workbook struct {
    File       *excelize.File
    Worksheets []*Worksheet
}

type Worksheet struct {
    SheetIndex int
    SheetName  string
    Metadata   *NodeData
    Analysis   *NodeData
}

type NodeData struct {
    NodeData      [][]interface{}
    CellNames     [][]string
    NodeHeader    []string
    HazopElements []Element
    HeaderAligned bool
    DataLogger    *Logger
    HeaderLogger  *Logger
}

type Logger struct {
    Warnings []string
    Errors   []string
    Info     []string
}

func ReadVerifyWorkbook(fpath string, wg *sync.WaitGroup) (*Workbook, error) {
    wb, err := newWorkbook(fpath)
    if err != nil {
        return nil, err
    }

    for i, name := range wb.File.GetSheetMap() {
        wg.Add(1)

        go func(i int, name string) {
            defer wg.Done()

            ws := wb.newWorksheet(i, name)

            ncols, err := wb.getNCols(name)
            if err != nil {
                log.Println(err)
                return
            }

            nrows, err := wb.getNRows(name)
            if err != nil {
                log.Println(err)
                return
            }

            metadataElements, err := Hazop.Elements(Hazop.DataType.Metadata)
            if err != nil {
                log.Println(err)
                return
            }

            analysisElements, err := Hazop.Elements(Hazop.DataType.Analysis)
            if err != nil {
                log.Println(err)
                return
            }

            metadataReader := reader{
                varDimension: readXCoordinate,
                fixDimension: readYCoordinate,
                cellNames:    readXCellNames,
            }

            analysisReader := reader{
                varDimension: readYCoordinate,
                fixDimension: readXCoordinate,
                cellNames:    readYCellNames,
            }

            metadataReadVerifier := &readVerifier{
                nsize:    ncols,
                elements: metadataElements,
                sname:    ws.SheetName,
                node:     ws.Metadata,
                reader:   metadataReader,
            }

            analysisReadVerifier := &readVerifier{
                nsize:    nrows,
                elements: analysisElements,
                sname:    ws.SheetName,
                node:     ws.Analysis,
                reader:   analysisReader,
            }

            wg.Add(2)
            go wb.readVerifyNodeData(metadataReadVerifier, wg)
            go wb.readVerifyNodeData(analysisReadVerifier, wg)

            wb.Worksheets = append(wb.Worksheets, ws)
        }(i, name)
    }

    if err := wb.File.Close(); err != nil {
        return nil, fmt.Errorf("%v %v", ErrClosingExcelFile, err)
    }

    return wb, nil
}

func newWorkbook(fpath string) (*Workbook, error) {
    f, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, fmt.Errorf("%v: %v", ErrOpeningExcelFile, err)
    }

    return &Workbook{File: f}, nil
}

func (wb *Workbook) newWorksheet(sheetIndex int, sheetName string) *Worksheet {
    return &Worksheet{
        SheetIndex: sheetIndex,
        SheetName:  sheetName,
        Metadata: &NodeData{
            DataLogger:   &Logger{},
            HeaderLogger: &Logger{},
        },
        Analysis: &NodeData{
            DataLogger:   &Logger{},
            HeaderLogger: &Logger{},
        },
    }
}

func (wb *Workbook) getNCols(name string) (int, error) {
    cols, err := wb.File.GetCols(name)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrReadingColumns, err)
    }

    return len(cols), nil
}

func (wb *Workbook) getNRows(name string) (int, error) {
    rows, err := wb.File.GetRows(name)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrReadingRows, err)
    }
    return len(rows), nil
}

type readVerifier struct {
    nsize    int
    elements []Element
    sname    string
    node     *NodeData
    reader   reader
}

func (wb *Workbook) readVerifyNodeData(rv *readVerifier, wg *sync.WaitGroup) {
    defer wg.Done()

    if err := wb.readHazopElements(rv); err != nil {
        log.Println(err)
        return
    }

    if err := rv.node.verifyElementsAlignment(rv.reader); err != nil {
        log.Println(err)
        return
    }

    if err := wb.readNodeData(rv); err != nil {
        log.Println(err)
        return
    }

    if err := rv.node.verifyNodeData(); err != nil {
        log.Println(err)
        return
    }
}

func (wb *Workbook) readHazopElements(rv *readVerifier) error {
    for _, element := range rv.elements {
        coords, err := wb.File.SearchSheet(rv.sname, element.Regex, true)
        if err != nil {
            return fmt.Errorf("%v: %v", ErrSearchingHeader, err)
        }

        filterHazopElements(coords, element, rv.node)
    }
    return nil
}

func filterHazopElements(coords []string, element Element, node *NodeData) {
    switch len(coords) {
    case 0:
        msg := fmt.Sprintf("%s: `%s`", HeaderNotFound, element.Name)
        node.HeaderLogger.newWarning(msg)
    case 1:
        node.NodeHeader = append(node.NodeHeader, coords[0])
        node.HazopElements = append(node.HazopElements, element)
        name, coord := element.Name, coords[0]
        msg := fmt.Sprintf("%s: `%s` `%s`", HeaderFound, name, coord)
        node.HeaderLogger.newInfo(msg)
    default:
        msg := fmt.Sprintf("%v: `%s` %v", HeaderMultipleCoordinates, element.Name, coords)
        node.HeaderLogger.newWarning(msg)
    }
}

func (node *NodeData) verifyElementsAlignment(r reader) error {
    if len(node.NodeHeader) == 0 {
        node.HeaderAligned = false
        node.HeaderLogger.newError(ErrNoHeaderFound.Error())
        return nil
    }

    if len(node.NodeHeader) == 1 {
        node.HeaderAligned = false
        msg := fmt.Sprintf("%v: %v", ErrNotEnoughHeader, node.NodeHeader)
        node.HeaderLogger.newError(msg)
        return nil
    }

    aligned, err := node.elementsAligned(r)
    if err != nil {
        return err
    }

    if !aligned {
        node.HeaderAligned = false
        msg := fmt.Sprintf("%v: %v", ErrHeaderNotAligned, node.NodeHeader)
        node.HeaderLogger.newError(msg)
    }

    node.HeaderAligned = true
    msg := fmt.Sprintf("%s: %v", HeaderAligned, node.NodeHeader)
    node.HeaderLogger.newInfo(msg)

    return nil
}

func (node *NodeData) elementsAligned(r reader) (bool, error) {
    ref, err := r.varDimension(node.NodeHeader[0])
    if err != nil {
        return false, err
    }

    for i := 1; i < len(node.NodeHeader); i++ {
        v, err := r.varDimension(node.NodeHeader[i])
        if err != nil {
            return false, err
        }

        if ref != v {
            return false, nil
        }
    }

    return true, nil
}

func (wb *Workbook) readNodeData(rv *readVerifier) error {
    rv.node.NodeData = make([][]interface{}, len(rv.node.NodeHeader))
    rv.node.CellNames = make([][]string, len(rv.node.NodeHeader))

    for i, cname := range rv.node.NodeHeader {
        cnames, err := rv.reader.readCellNames(cname, rv.nsize)
        if err != nil {
            return err
        }

        cols := make([]interface{}, len(cnames))
        for k := 0; k < len(cnames); k++ {
            value, err := wb.File.GetCellValue(rv.sname, cnames[k])
            if err != nil {
                return fmt.Errorf("%s: %v", ErrReadingCellValue, err)
            }

            cols[k] = value
        }

        rv.node.NodeData[i] = cols
        rv.node.CellNames[i] = cnames
    }

    return nil
}

func (node *NodeData) verifyNodeData() error {
    for i := 0; i < len(node.NodeHeader); i++ {
        verifier, err := newCellVerifier(node.HazopElements[i].CellType)
        if err != nil {
            return err
        }

        for k, v := range node.NodeData[i] {
            cell, err := verifier.checkCellType(v)
            if err != nil {
                msg := fmt.Sprintf("%v: `%s`", err, node.CellNames[i][k])
                node.DataLogger.newError(msg)
                continue
            }

            min := node.HazopElements[i].MinLen
            max := node.HazopElements[i].MaxLen

            if err := verifier.checkCellLength(cell, min, max); err != nil {
                msg := fmt.Sprintf("%v: `%s`", err, node.CellNames[i][k])
                node.DataLogger.newError(msg)
                continue
            }

            cname := node.CellNames[i][k]
            msg := fmt.Sprintf("%s: `%s`", ValueParsedVerified, cname)
            node.DataLogger.newInfo(msg)
        }
    }

    return nil
}

func (l *Logger) newWarning(msg string) {
    l.Warnings = append(l.Warnings, msg)
}

func (l *Logger) newError(msg string) {
    l.Errors = append(l.Errors, msg)
}

func (l *Logger) newInfo(msg string) {
    l.Info = append(l.Info, msg)
}

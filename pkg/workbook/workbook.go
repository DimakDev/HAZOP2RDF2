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
    f, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, fmt.Errorf("%v: %v", ErrOpeningExcelFile, err)
    }

    var wb = &Workbook{File: f}
    for i, name := range wb.File.GetSheetMap() {
        wg.Add(1)

        go func(i int, name string) {
            defer wg.Done()

            ws := &Worksheet{
                SheetIndex: i,
                SheetName:  name,
                Metadata: &NodeData{
                    DataLogger:   &Logger{},
                    HeaderLogger: &Logger{},
                },
                Analysis: &NodeData{
                    DataLogger:   &Logger{},
                    HeaderLogger: &Logger{},
                },
            }

            cols, err := wb.File.GetCols(name)
            if err != nil {
                log.Println(fmt.Errorf("%v: %v", ErrReadingColumns, err))
                return
            }

            rows, err := wb.File.GetRows(name)
            if err != nil {
                log.Println(fmt.Errorf("%v: %v", ErrReadingRows, err))
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
                nsize:    len(cols),
                elements: metadataElements,
                sname:    ws.SheetName,
                node:     ws.Metadata,
                reader:   metadataReader,
            }

            analysisReadVerifier := &readVerifier{
                nsize:    len(rows),
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

type readVerifier struct {
    nsize    int
    elements []Element
    sname    string
    node     *NodeData
    reader   reader
}

func (wb *Workbook) readVerifyNodeData(rv *readVerifier, wg *sync.WaitGroup) {
    defer wg.Done()

    for _, element := range rv.elements {
        coords, err := wb.File.SearchSheet(rv.sname, element.Regex, true)
        if err != nil {
            log.Println(fmt.Errorf("%v: %v", ErrSearchingHeader, err))
            return
        }

        rv.node.verifyElementCoords(coords, element)
    }

    if err := rv.node.verifyHeaderAlignment(rv.reader); err != nil {
        log.Println(err)
        return
    }

    rv.node.NodeData = make([][]interface{}, len(rv.node.NodeHeader))
    rv.node.CellNames = make([][]string, len(rv.node.NodeHeader))

    for i := 0; i < len(rv.node.NodeHeader); i++ {
        cnames, err := rv.reader.readCellNames(rv.node.NodeHeader[i], rv.nsize)
        if err != nil {
            log.Println(err)
            return
        }

        verifier, err := newCellVerifier(rv.node.HazopElements[i].CellType)
        if err != nil {
            log.Println(err)
            return
        }

        cols := make([]interface{}, len(cnames))
        for k := 0; k < len(cnames); k++ {
            value, err := wb.File.GetCellValue(rv.sname, cnames[k])
            if err != nil {
                log.Println(fmt.Errorf("%s: %v", ErrReadingCellValue, err))
                return
            }

            cell, err := verifier.checkCellType(value)
            if err != nil {
                rv.node.DataLogger.newError(
                    fmt.Sprintf("%v: `%s`",
                        err,
                        cnames[k],
                    ),
                )
                continue
            }

            if err := verifier.checkCellLength(
                cell,
                rv.node.HazopElements[i].MinLen,
                rv.node.HazopElements[i].MaxLen,
            ); err != nil {
                rv.node.DataLogger.newError(
                    fmt.Sprintf("%v: `%s`",
                        err,
                        cnames[k],
                    ),
                )
                continue
            }

            rv.node.DataLogger.newInfo(
                fmt.Sprintf("%s: `%s`",
                    ValueParsedVerified,
                    cnames[k],
                ),
            )

            cols[k] = value
        }

        rv.node.NodeData[i] = cols
        rv.node.CellNames[i] = cnames
    }
}

func (node *NodeData) verifyElementCoords(coords []string, element Element) {
    switch len(coords) {
    case 0:
        node.HeaderLogger.newWarning(
            fmt.Sprintf("%s: `%s`",
                HeaderNotFound,
                element.Name,
            ),
        )
    case 1:
        node.NodeHeader = append(node.NodeHeader, coords[0])
        node.HazopElements = append(node.HazopElements, element)
        node.HeaderLogger.newInfo(
            fmt.Sprintf("%s: `%s` `%s`",
                HeaderFound,
                element.Name,
                coords[0],
            ),
        )
    default:
        node.HeaderLogger.newWarning(
            fmt.Sprintf("%v: `%s` %v",
                HeaderMultipleCoordinates,
                element.Name,
                coords,
            ),
        )
    }
}

func (node *NodeData) verifyHeaderAlignment(r reader) error {
    if len(node.NodeHeader) == 0 {
        node.HeaderAligned = false
        node.HeaderLogger.newError(ErrNoHeaderFound.Error())
        return nil
    }

    if len(node.NodeHeader) == 1 {
        node.HeaderAligned = false
        node.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrNotEnoughHeader,
                node.NodeHeader,
            ),
        )
        return nil
    }

    ref, err := r.varDimension(node.NodeHeader[0])
    if err != nil {
        return err
    }

    for i := 1; i < len(node.NodeHeader); i++ {
        v, err := r.varDimension(node.NodeHeader[i])
        if err != nil {
            return err
        }

        if ref != v {
            node.HeaderAligned = false
            node.HeaderLogger.newError(
                fmt.Sprintf("%v: %v",
                    ErrHeaderNotAligned,
                    node.NodeHeader,
                ),
            )
        }
    }

    node.HeaderAligned = true
    node.HeaderLogger.newInfo(
        fmt.Sprintf("%s: %v",
            HeaderAligned,
            node.NodeHeader,
        ),
    )

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

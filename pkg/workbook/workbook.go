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
    file, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, fmt.Errorf("%v: %v", ErrOpeningExcelFile, err)
    }

    var wb = &Workbook{File: file}

    for sindex, sname := range wb.File.GetSheetMap() {
        wg.Add(1)
        go func(sindex int, sname string) {
            defer wg.Done()

            cols, err := wb.File.GetCols(sname)
            if err != nil {
                log.Println(fmt.Errorf("%v: %v", ErrReadingColumns, err))
                return
            }

            rows, err := wb.File.GetRows(sname)
            if err != nil {
                log.Println(fmt.Errorf("%v: %v", ErrReadingRows, err))
                return
            }

            ncols := len(cols)
            nrows := len(rows)

            metadata := &NodeData{
                DataLogger:   &Logger{},
                HeaderLogger: &Logger{},
            }

            analysis := &NodeData{
                DataLogger:   &Logger{},
                HeaderLogger: &Logger{},
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

            metadataArgs := &args{
                sname: sname,
                nsize: ncols,
                node:  metadata,
                reader: &reader{
                    varDimension: readXCoordinate,
                    fixDimension: readYCoordinate,
                    cellNames:    readXCellNames,
                },
                elements: metadataElements,
            }

            analysisArgs := &args{
                sname: sname,
                nsize: nrows,
                node:  analysis,
                reader: &reader{
                    varDimension: readYCoordinate,
                    fixDimension: readXCoordinate,
                    cellNames:    readYCellNames,
                },
                elements: analysisElements,
            }

            wg.Add(2)
            go wb.readVerifyNodeData(metadataArgs, wg)
            go wb.readVerifyNodeData(analysisArgs, wg)

            wb.Worksheets = append(wb.Worksheets, &Worksheet{
                SheetIndex: sindex,
                SheetName:  sname,
                Metadata:   metadata,
                Analysis:   analysis,
            })
        }(sindex, sname)
    }

    if err := wb.File.Close(); err != nil {
        return nil, fmt.Errorf("%v %v", ErrClosingExcelFile, err)
    }

    return wb, nil
}

type args struct {
    sname    string
    nsize    int
    node     *NodeData
    reader   *reader
    elements []Element
}

func (wb *Workbook) readVerifyNodeData(args *args, wg *sync.WaitGroup) {
    defer wg.Done()
    if err := wb.readNodeHeader(args); err != nil {
        log.Println(err)
        return
    }

    if err := verifyNodeHeader(args); err != nil {
        log.Println(err)
        return
    }

    if err := wb.readNodeData(args); err != nil {
        log.Println(err)
        return
    }

    if err := verifyNodeData(args); err != nil {
        log.Println(err)
        return
    }
}

func (wb *Workbook) readNodeHeader(args *args) error {
    for _, element := range args.elements {
        coords, err := wb.File.SearchSheet(args.sname, element.Regex, true)
        if err != nil {
            return fmt.Errorf("%v: %v", ErrSearchingHeader, err)
        }

        switch len(coords) {
        case 0:
            args.node.HeaderLogger.newWarning(
                fmt.Sprintf("%s: `%s`",
                    HeaderNotFound,
                    element.Name,
                ),
            )
        case 1:
            args.node.NodeHeader = append(args.node.NodeHeader, coords[0])
            args.node.HazopElements = append(args.node.HazopElements, element)

            args.node.HeaderLogger.newInfo(
                fmt.Sprintf("%s: `%s` `%s`",
                    HeaderFound,
                    element.Name,
                    coords[0],
                ),
            )
        default:
            args.node.HeaderLogger.newWarning(
                fmt.Sprintf("%v: `%s` %v",
                    HeaderMultipleCoordinates,
                    element.Name,
                    coords,
                ),
            )
        }
    }

    return nil
}

func verifyNodeHeader(args *args) error {
    if len(args.node.NodeHeader) == 0 {
        args.node.HeaderAligned = false
        args.node.HeaderLogger.newError(ErrNoHeaderFound.Error())
        return nil
    }

    if len(args.node.NodeHeader) == 1 {
        args.node.HeaderAligned = false
        args.node.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrNotEnoughHeader,
                args.node.NodeHeader,
            ),
        )
        return nil
    }

    ref, err := args.reader.varDimension(args.node.NodeHeader[0])
    if err != nil {
        return err
    }

    for i := 1; i < len(args.node.NodeHeader); i++ {
        v, err := args.reader.varDimension(args.node.NodeHeader[i])
        if err != nil {
            return err
        }

        if ref != v {
            args.node.HeaderAligned = false
            args.node.HeaderLogger.newError(
                fmt.Sprintf("%v: %v",
                    ErrHeaderNotAligned,
                    args.node.NodeHeader,
                ),
            )
            return nil
        }
    }

    args.node.HeaderAligned = true
    args.node.HeaderLogger.newInfo(
        fmt.Sprintf("%s: %v",
            HeaderAligned,
            args.node.NodeHeader,
        ),
    )

    return nil
}

func (wb *Workbook) readNodeData(args *args) error {
    args.node.NodeData = make([][]interface{}, len(args.node.NodeHeader))
    args.node.CellNames = make([][]string, len(args.node.NodeHeader))

    for i, cname := range args.node.NodeHeader {
        cnames, err := args.reader.readCellNames(cname, args.nsize)
        if err != nil {
            return err
        }

        data := make([]interface{}, len(cnames))
        for k := 0; k < len(cnames); k++ {
            cell, err := wb.File.GetCellValue(args.sname, cnames[k])
            if err != nil {
                return fmt.Errorf("%s: %v", ErrReadingCellValue, err)
            }

            data[k] = cell
        }

        args.node.NodeData[i] = data
        args.node.CellNames[i] = cnames
    }

    return nil
}

func verifyNodeData(args *args) error {
    for i := 0; i < len(args.node.NodeHeader); i++ {
        verifier, err := newCellVerifier(args.node.HazopElements[i].CellType)
        if err != nil {
            return err
        }

        for k, c := range args.node.NodeData[i] {
            cell, err := verifier.checkCellType(c)
            if err != nil {
                args.node.DataLogger.newError(
                    fmt.Sprintf("%v: `%s`",
                        err,
                        args.node.CellNames[i][k],
                    ),
                )
                continue
            }

            if err := verifier.checkCellLength(
                cell,
                args.node.HazopElements[i].MinLen,
                args.node.HazopElements[i].MaxLen,
            ); err != nil {
                args.node.DataLogger.newError(
                    fmt.Sprintf("%v: `%s`",
                        err,
                        args.node.CellNames[i][k],
                    ),
                )
                continue
            }

            args.node.DataLogger.newInfo(
                fmt.Sprintf("%s: `%s`",
                    ValueParsedVerified,
                    args.node.CellNames[i][k],
                ),
            )
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

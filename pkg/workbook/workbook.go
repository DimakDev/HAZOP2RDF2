package workbook

import (
    "errors"
    "fmt"
    "log"
    "path/filepath"
    "strings"
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
    Name       string
    Worksheets []*Worksheet
}

type Worksheet struct {
    SheetIndex int
    SheetName  string
    Metadata   *NodeData
    Analysis   *NodeData
}

type NodeData struct {
    Data          map[int][]interface{}
    Header        map[int]string
    Element       map[int]Element
    HeaderAligned bool
    DataLogger    *Logger
    HeaderLogger  *Logger
}

type Logger struct {
    Warnings []string
    Errors   []string
    Info     []string
}

func NewWorkbook(datapath string) (*Workbook, error) {
    f, err := excelize.OpenFile(datapath)
    if err != nil {
        return nil, fmt.Errorf("%v: %v", ErrOpeningExcelFile, err)
    }

    _, filename := filepath.Split(datapath)
    wbname := strings.TrimSuffix(filename, filepath.Ext(filename))

    return &Workbook{
        File: f,
        Name: wbname,
    }, nil
}

func (wb *Workbook) ReadVerifyWorkbook(wg *sync.WaitGroup) error {
    for sindex, sname := range wb.File.GetSheetMap() {
        wg.Add(1)
        go func(sindex int, sname string) {
            defer wg.Done()
            metadata := &NodeData{
                Data:         map[int][]interface{}{},
                Header:       map[int]string{},
                Element:      map[int]Element{},
                DataLogger:   &Logger{},
                HeaderLogger: &Logger{},
            }

            analysis := &NodeData{
                Data:         map[int][]interface{}{},
                Header:       map[int]string{},
                Element:      map[int]Element{},
                DataLogger:   &Logger{},
                HeaderLogger: &Logger{},
            }

            if err := wb.readHazopElements(
                sname,
                Hazop.Elements(Hazop.DataType.Metadata),
                metadata,
            ); err != nil {
                log.Println(err)
                return
            }

            if err := wb.readHazopElements(
                sname,
                Hazop.Elements(Hazop.DataType.Analysis),
                analysis,
            ); err != nil {
                log.Println(err)
                return
            }

            metadataCoords, err := cellNamesToCoordinates(metadata.Header)
            if err != nil {
                log.Println(err)
                return
            }

            analysisCoords, err := cellNamesToCoordinates(analysis.Header)
            if err != nil {
                log.Println(err)
                return
            }

            metadata.verifyHeaderAlignment(
                metadataCoords.coordsX,
                metadataCoords.cnames,
            )

            analysis.verifyHeaderAlignment(
                analysisCoords.coordsY,
                analysisCoords.cnames,
            )

            numCols, err := wb.getNCols(sname)
            if err != nil {
                log.Println(err)
                return
            }

            numRows, err := wb.getNRows(sname)
            if err != nil {
                log.Println(err)
                return
            }

            metadataReader := &reader{
                varDimension: readXCoordinate,
                fixDimension: readYCoordinate,
                cellNames:    readXCellNames,
            }

            analysisReader := &reader{
                varDimension: readYCoordinate,
                fixDimension: readXCoordinate,
                cellNames:    readYCellNames,
            }

            if err := wb.readVerifyHazopData(
                sname,
                numCols,
                metadataReader,
                metadata,
            ); err != nil {
                log.Println(err)
                return
            }

            if err := wb.readVerifyHazopData(
                sname,
                numRows,
                analysisReader,
                analysis,
            ); err != nil {
                log.Println(err)
                return
            }

            wb.Worksheets = append(wb.Worksheets, &Worksheet{
                SheetIndex: sindex,
                SheetName:  sname,
                Metadata:   metadata,
                Analysis:   analysis,
            })
        }(sindex, sname)
    }

    if err := wb.File.Close(); err != nil {
        return fmt.Errorf("%v %v", ErrClosingExcelFile, err)
    }

    return nil
}

func (wb *Workbook) readVerifyHazopData(sname string, size int, reader *reader, node *NodeData) error {
    for hindex, hname := range node.Header {
        element := node.Element[hindex]

        d1, err := reader.varDimension(hname)
        if err != nil {
            return err
        }

        d2, err := reader.fixDimension(hname)
        if err != nil {
            return err
        }

        cnames, err := reader.cellNames(d1, d2, size-d1)
        if err != nil {
            return err
        }

        var verifier cellVerifier

        switch element.CellType {
        case Hazop.CellType.String:
            verifier = verifyString{}
        case Hazop.CellType.Integer:
            verifier = verifyInteger{}
        case Hazop.CellType.Float:
            verifier = verifyFloat{}
        default:
            return fmt.Errorf("%v: %d", ErrUnknownCellType, element.CellType)
        }

        data := make([]interface{}, len(cnames))
        data[0] = element.Name

        for i := 1; i < len(cnames); i++ {
            value, err := wb.File.GetCellValue(sname, cnames[i])
            if err != nil {
                return fmt.Errorf("%s: %v", ErrReadingCellValue, err)
            }

            cell, err := verifier.checkCellType(value)
            if err != nil {
                node.DataLogger.newError(
                    fmt.Sprintf("%v: `%s`",
                        err,
                        cnames[i],
                    ),
                )
                continue
            }

            err = verifier.checkCellLength(
                cell,
                element.MinLen,
                element.MaxLen,
            )
            if err != nil {
                node.DataLogger.newError(
                    fmt.Sprintf("%v: `%s`",
                        err,
                        cnames[i],
                    ),
                )
                continue
            }

            node.DataLogger.newInfo(
                fmt.Sprintf("%s: `%s`",
                    ValueParsedVerified,
                    cnames[i],
                ),
            )

            data[i] = cell
        }

        node.Data[hindex] = data
    }

    return nil
}

func (wb *Workbook) readHazopElements(sname string, elements map[int]Element, node *NodeData) error {
    for i, element := range elements {
        coords, err := wb.File.SearchSheet(sname, element.Regex, true)
        if err != nil {
            return fmt.Errorf("%v: %v", ErrSearchingHeader, err)
        }

        switch len(coords) {
        case 0:
            node.HeaderLogger.newWarning(
                fmt.Sprintf("%s: `%s`",
                    HeaderNotFound,
                    element.Name,
                ),
            )
        case 1:
            node.Header[i], node.Element[i] = coords[0], element
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

    return nil
}

func (node *NodeData) verifyHeaderAlignment(coords []int, cnames []string) {
    if len(coords) == 0 {
        node.HeaderAligned = false
        node.HeaderLogger.newError(ErrNoHeaderFound.Error())
        return
    }

    if len(coords) == 1 {
        node.HeaderAligned = false
        node.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrNotEnoughHeader,
                cnames,
            ),
        )
        return
    }

    if !checkHeaderAlignment(coords) {
        node.HeaderAligned = false
        node.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrHeaderNotAligned,
                cnames,
            ),
        )
        return
    }

    node.HeaderAligned = true
    node.HeaderLogger.newInfo(fmt.Sprintf("%s: %v", HeaderAligned, cnames))
}

func (wb *Workbook) getNCols(sname string) (int, error) {
    cols, err := wb.File.GetCols(sname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrReadingColumns, err)
    }

    return len(cols), nil
}

func (wb *Workbook) getNRows(sname string) (int, error) {
    rows, err := wb.File.GetRows(sname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrReadingRows, err)
    }

    return len(rows), nil
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

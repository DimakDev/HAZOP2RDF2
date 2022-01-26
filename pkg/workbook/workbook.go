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
    ErrOpeningExcelFile = errors.New("Error opening Excel file")
    ErrClosingExcelFile = errors.New("Error closing Excel file")
    ErrUnknownCellType  = errors.New("Unknown cell type")
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
    for i, sname := range wb.File.GetSheetMap() {
        wg.Add(1)
        go func(i int, sname string) {
            defer wg.Done()
            nodeM := newNodeData()
            nodeA := newNodeData()

            elementsM := Hazop.groupHazopElements(Hazop.DataType.Metadata)
            elementsA := Hazop.groupHazopElements(Hazop.DataType.Analysis)

            if err := wb.readHazopElements(
                sname,
                elementsM,
                nodeM,
            ); err != nil {
                log.Println(err)
                return
            }

            if err := wb.readHazopElements(
                sname,
                elementsA,
                nodeA,
            ); err != nil {
                log.Println(err)
                return
            }

            coordsM, err := cellNamesToCoordinates(nodeM.Header)
            if err != nil {
                log.Println(err)
                return
            }

            coordsA, err := cellNamesToCoordinates(nodeA.Header)
            if err != nil {
                log.Println(err)
                return
            }
            verifyHeaderAlignment(coordsM.coordX, coordsM.cnames, nodeM)
            verifyHeaderAlignment(coordsA.coordY, coordsA.cnames, nodeA)

            nCols, err := wb.getNCols(sname)
            if err != nil {
                log.Println(err)
                return
            }

            nRows, err := wb.getNRows(sname)
            if err != nil {
                log.Println(err)
                return
            }

            readerM := &reader{
                varDimension: readXCoordinates,
                fixDimension: readYCoordinates,
                cellNames:    readXCellNames,
            }

            readerA := &reader{
                varDimension: readYCoordinates,
                fixDimension: readXCoordinates,
                cellNames:    readYCellNames,
            }

            if err := wb.readVerifyHazopData(
                sname,
                nCols,
                readerM,
                nodeM,
            ); err != nil {
                log.Println(err)
                return
            }

            if err := wb.readVerifyHazopData(
                sname,
                nRows,
                readerA,
                nodeA,
            ); err != nil {
                log.Println(err)
                return
            }

            wb.Worksheets = append(wb.Worksheets, &Worksheet{
                SheetIndex: i,
                SheetName:  sname,
                Metadata:   nodeM,
                Analysis:   nodeA,
            })
        }(i, sname)
    }

    if err := wb.File.Close(); err != nil {
        return fmt.Errorf("%v %v", ErrClosingExcelFile, err)
    }

    return nil
}

func (wb *Workbook) readVerifyHazopData(
    sname string,
    nmax int,
    r *reader,
    n *NodeData,
) error {
    for k, cname := range n.Header {
        e := n.Element[k]

        d1, err := r.varDimension(cname)
        if err != nil {
            return err
        }

        d2, err := r.fixDimension(cname)
        if err != nil {
            return err
        }

        cnames, err := r.cellNames(d1, d2, nmax-d1)
        if err != nil {
            return err
        }

        var v cellVerifier

        switch e.CellType {
        case Hazop.CellType.String:
            v = verifyString{}
        case Hazop.CellType.Integer:
            v = verifyInteger{}
        case Hazop.CellType.Float:
            v = verifyFloat{}
        default:
            return fmt.Errorf("%v: %d", ErrUnknownCellType, e.CellType)
        }

        vec := make([]interface{}, len(cnames))
        vec[0] = e.Name

        for i := 1; i < len(cnames); i++ {
            cell, err := wb.File.GetCellValue(sname, cnames[i])
            if err != nil {
                return fmt.Errorf("%s: %v", ErrReadingCellValue, err)
            }

            c, err := v.checkCellType(cell)
            if err != nil {
                n.DataLogger.newError(fmt.Sprintf("%v: `%s`", err, cnames[i]))
                continue
            }

            if err := v.checkCellLength(c, e.MinLen, e.MaxLen); err != nil {
                n.DataLogger.newError(fmt.Sprintf("%v: `%s`", err, cnames[i]))
                continue
            }

            n.DataLogger.newInfo(
                fmt.Sprintf("%s: `%s`",
                    ValueParsedVerified,
                    cnames[i],
                ),
            )

            vec[i] = c
        }

        n.Data[k] = vec
    }

    return nil
}

func newNodeData() *NodeData {
    return &NodeData{
        Data:         map[int][]interface{}{},
        Header:       map[int]string{},
        Element:      map[int]Element{},
        DataLogger:   &Logger{},
        HeaderLogger: &Logger{},
    }
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

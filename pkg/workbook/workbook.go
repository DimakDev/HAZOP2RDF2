package workbook

import (
    "errors"
    "fmt"
    "path/filepath"
    "strings"

    "github.com/xuri/excelize/v2"
)

var ErrOpeningFile = errors.New("Error opening file")

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
        return nil, fmt.Errorf("%v: %v", ErrOpeningFile, err)
    }

    _, filename := filepath.Split(datapath)
    wbname := strings.TrimSuffix(filename, filepath.Ext(filename))

    wb := &Workbook{
        File: f,
        Name: wbname,
    }

    if err := wb.newWorkbook(); err != nil {
        return nil, err
    }

    return wb, nil
}

func (wb *Workbook) newWorkbook() error {
    // Abbr: M — Metadata, A — Analysis
    elementsM := groupElements(Hazop.DataType.Metadata)
    elementsA := groupElements(Hazop.DataType.Analysis)

    for i, sname := range wb.File.GetSheetMap() {
        ws := &Worksheet{
            SheetIndex: i,
            SheetName:  sname,
            Metadata: &NodeData{
                Data:         map[int][]interface{}{},
                Header:       map[int]string{},
                Element:      map[int]Element{},
                DataLogger:   &Logger{},
                HeaderLogger: &Logger{},
            },
            Analysis: &NodeData{
                Data:         map[int][]interface{}{},
                Header:       map[int]string{},
                Element:      map[int]Element{},
                DataLogger:   &Logger{},
                HeaderLogger: &Logger{},
            },
        }

        nodeM := ws.Metadata
        nodeA := ws.Analysis

        if err := wb.readHazopHeader(sname, elementsM, nodeM); err != nil {
            return err
        }

        if err := wb.readHazopHeader(sname, elementsA, nodeA); err != nil {
            return err
        }

        headerM, err := splitHeader(nodeM.Header)
        if err != nil {
            return err
        }

        headerA, err := splitHeader(nodeA.Header)
        if err != nil {
            return err
        }

        verifyHeaderAlignment(headerM.coordX, headerM.coords, nodeM)
        verifyHeaderAlignment(headerA.coordY, headerA.coords, nodeA)

        ncols, err := wb.getNCols(sname)
        if err != nil {
            return err
        }

        nrows, err := wb.getNRows(sname)
        if err != nil {
            return err
        }

        readerM := &reader{
            runner: readXCoord,
            fixer:  readYCoord,
            cnames: readXCnames,
        }

        readerA := &reader{
            runner: readYCoord,
            fixer:  readXCoord,
            cnames: readYCnames,
        }

        if err := wb.readHazopData(sname, ncols, readerM, nodeM); err != nil {
            return err
        }

        if err := wb.readHazopData(sname, nrows, readerA, nodeA); err != nil {
            return err
        }

        wb.Worksheets = append(wb.Worksheets, ws)
    }

    return nil
}

func (r *Logger) newWarning(msg string) {
    r.Warnings = append(r.Warnings, msg)
}

func (r *Logger) newError(msg string) {
    r.Errors = append(r.Errors, msg)
}

func (r *Logger) newInfo(msg string) {
    r.Info = append(r.Info, msg)
}

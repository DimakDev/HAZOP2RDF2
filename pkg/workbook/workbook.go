package workbook

import (
    "errors"
    "fmt"
    "log"
    "path/filepath"
    "strings"

    "github.com/xuri/excelize/v2"
)

var ErrOpeningFile = errors.New("Error opening file")

type Workbook struct {
    File       *excelize.File
    PlantName  string
    PlantHazop map[int]*NodeHazop
}

type NodeHazop struct {
    Metadata *NodeData
    Analysis *NodeData
}

type NodeData struct {
    Data          map[int][]interface{}
    Header        map[int]string
    HeaderAligned bool
    DataReport    *Report
    HeaderReport  *Report
}

type Report struct {
    Warnings []string
    Errors   []error
    Info     []string
}

func NewWorkbook(datapath string) (*Workbook, error) {
    f, err := excelize.OpenFile(datapath)
    if err != nil {
        return nil, fmt.Errorf("%v: %v", ErrOpeningFile, err)
    }

    _, filename := filepath.Split(datapath)
    plantname := strings.TrimSuffix(filename, filepath.Ext(filename))

    wb := &Workbook{
        File:       f,
        PlantName:  plantname,
        PlantHazop: map[int]*NodeHazop{},
    }

    if err := wb.readWorkbook(); err != nil {
        return nil, err
    }

    return wb, nil
}

func (wb *Workbook) readWorkbook() error {
    // Abbr: M — Metadata, A — Analysis
    typeM := settings.Hazop.DataType.Metadata
    typeA := settings.Hazop.DataType.Analysis

    elementsM := groupElements(typeM)
    elementsA := groupElements(typeA)

    for i, sname := range wb.File.GetSheetMap() {
        wb.PlantHazop[i] = &NodeHazop{
            Metadata: &NodeData{
                Data:         map[int][]interface{}{},
                Header:       map[int]string{},
                DataReport:   &Report{},
                HeaderReport: &Report{},
            },
            Analysis: &NodeData{
                Data:         map[int][]interface{}{},
                Header:       map[int]string{},
                DataReport:   &Report{},
                HeaderReport: &Report{},
            },
        }

        nodeM := wb.PlantHazop[i].Metadata
        nodeA := wb.PlantHazop[i].Analysis

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

        nrows, err := wb.getNCols(sname)
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
            return nil
        }

        if err := wb.readHazopData(sname, nrows, readerA, nodeA); err != nil {
            return nil
        }

        log.Println(nodeM.Data)
        log.Println(nodeM.DataReport)
        log.Println(nodeM.Header)
        log.Println(nodeM.HeaderReport)
        log.Println(nodeM.HeaderAligned)
        log.Println(nodeA.Data)
        log.Println(nodeA.DataReport)
        log.Println(nodeA.Header)
        log.Println(nodeA.HeaderReport)
        log.Println(nodeA.HeaderAligned)
    }

    return nil
}

func (r *Report) newWarning(msg string) {
    r.Warnings = append(r.Warnings, msg)
}

func (r *Report) newError(err error) {
    r.Errors = append(r.Errors, err)
}

func (r *Report) newInfo(msg string) {
    r.Info = append(r.Info, msg)
}

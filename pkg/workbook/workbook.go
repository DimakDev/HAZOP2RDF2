package workbook

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

type Workbook struct {
    File       *excelize.File
    Worksheets []*Worksheet
}

type Worksheet struct {
    SheetIndex    int
    SheetName     string
    NumOfCols     int
    NumOfRows     int
    HazopData     *HazopData
    HazopHeader   *HazopHeader
    HazopValidity *HazopValidity
}

type HazopData struct {
    Raw          [][]string
    NodeMetadata [][]interface{}
    NodeHazop    [][]interface{}
    Report       *Report
}

type HazopHeader struct {
    Raw          []string
    NodeMetadata []string
    NodeHazop    []string
    Report       *Report
}

type HazopValidity struct {
    NodeMetadata bool
    NodeHazop    bool
}

type Report struct {
    Warnings []string
    Errors   []error
    Info     []string
}

func NewWorkbook(fpath string) (*Workbook, error) {
    f, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, fmt.Errorf("Error opening file: %v", err)
    }
    wb := &Workbook{File: f}
    wb.initWorksheets()
    if err := wb.readWorkbook(); err != nil {
        return nil, err
    }
    if err := wb.verifyWorkbook(); err != nil {
        return nil, err
    }
    return wb, nil
}

func (wb *Workbook) initWorksheets() {
    for i, sname := range wb.File.GetSheetMap() {
        wb.Worksheets = append(wb.Worksheets, &Worksheet{
            SheetIndex: i,
            SheetName:  sname,
            HazopHeader: &HazopHeader{
                Raw:          make([]string, len(settings.Hazop.Element)),
                NodeMetadata: make([]string, len(settings.Hazop.Element)),
                NodeHazop:    make([]string, len(settings.Hazop.Element)),
                Report:       &Report{},
            },
            HazopData: &HazopData{
                Report: &Report{},
            },
            HazopValidity: &HazopValidity{},
        })
    }
}

func (hd *HazopData) newWarning(msg string) {
    hd.Report.Warnings = append(hd.Report.Warnings, msg)
}

func (hd *HazopData) newError(err error) {
    hd.Report.Errors = append(hd.Report.Errors, err)
}

func (hd *HazopData) newInfo(msg string) {
    hd.Report.Info = append(hd.Report.Info, msg)
}

func (hh *HazopHeader) newWarning(msg string) {
    hh.Report.Warnings = append(hh.Report.Warnings, msg)
}

func (hh *HazopHeader) newError(err error) {
    hh.Report.Errors = append(hh.Report.Errors, err)
}

func (hh *HazopHeader) newInfo(msg string) {
    hh.Report.Info = append(hh.Report.Info, msg)
}

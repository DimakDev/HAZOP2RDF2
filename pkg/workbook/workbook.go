package workbook

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

type Workbook struct {
    File              *excelize.File
    SheetMap          map[int]string
    HazopData         map[int]*HazopData
    HazopHeader       map[int]*HazopHeader
    HazopValidity     map[int]*HazopValidity
    HazopDataReport   map[int]*Report
    HazopHeaderReport map[int]*Report
}

type HazopData struct {
    NodeMetadata [][]interface{}
    NodeHazop    [][]interface{}
}

type HazopHeader struct {
    NodeMetadata map[int]string
    NodeHazop    map[int]string
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

    wb := &Workbook{
        File:              f,
        SheetMap:          map[int]string{},
        HazopData:         map[int]*HazopData{},
        HazopHeader:       map[int]*HazopHeader{},
        HazopValidity:     map[int]*HazopValidity{},
        HazopDataReport:   map[int]*Report{},
        HazopHeaderReport: map[int]*Report{},
    }

    if err := wb.readWorkbook(); err != nil {
        return nil, err
    }

    return wb, nil
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

func (h *HazopHeader) newHeader(i, dtype int, coord string) error {
    switch dtype {
    case settings.Hazop.DataType.NodeMetadata:
        h.NodeMetadata[i] = coord
    case settings.Hazop.DataType.NodeHazop:
        h.NodeHazop[i] = coord
    default:
        return fmt.Errorf("Unknown data type %d", dtype)
    }
    return nil
}

package workbook

import (
    "fmt"
    "path/filepath"
    "strings"

    "github.com/xuri/excelize/v2"
)

type Workbook struct {
    File       *excelize.File
    Name       string
    Path       string
    SheetMap   map[int]string
    Worksheets map[int]*Worksheet
    Errors     map[int][]error
}

type Worksheet struct {
    Header  []interface{}
    Columns [][]interface{}
    Errors  map[int][]error
}

func New(datapath string) *Workbook {
    _, fname := filepath.Split(datapath)
    return &Workbook{
        Name:       strings.TrimSuffix(fname, filepath.Ext(fname)),
        Path:       datapath,
        SheetMap:   map[int]string{},
        Worksheets: map[int]*Worksheet{},
        Errors:     map[int][]error{},
    }
}

func (wb *Workbook) ReadWorkbook() error {
    if f, err := excelize.OpenFile(wb.Path); err == nil {
        wb.File = f
        return nil
    } else {
        return fmt.Errorf("Error opening `%s`: %v", wb.Path, err)
    }
}

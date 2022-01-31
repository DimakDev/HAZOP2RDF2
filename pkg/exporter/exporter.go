package exporter

import (
    "errors"
    "fmt"
    "os"
    "text/template"

    "github.com/dimakdev/hazop-formula/pkg/workbook"
)

type Exporter struct {
    ReportPath     string
    GraphPath      string
    ProgramName    string
    ProgramVersion string
    DateTime       string
    BaseUri        string
    Workbook       string
    Worksheets     []*workbook.Worksheet
}

var (
    ErrCreatingFile    = errors.New("Error creating report file")
    ErrReadingTemplate = errors.New("Error reading report template")
    ErrWritingReport   = errors.New("Error writing report")
)

func (e *Exporter) WriteToFile(fpath, tpath string) error {
    f, err := os.Create(fpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrCreatingFile, fpath, err)
    }
    defer f.Close()

    t, err := template.ParseFiles(tpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrReadingTemplate, tpath, err)
    }

    if err := t.Execute(f, e); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, tpath, err)
    }

    return nil
}

func (e *Exporter) WriteToStdout(tpath string) error {
    t, err := template.ParseFiles(tpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrReadingTemplate, tpath, err)
    }

    if err := t.Execute(os.Stdout, e); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, tpath, err)
    }

    return nil
}

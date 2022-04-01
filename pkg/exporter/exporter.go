package exporter

import (
    "errors"
    "fmt"
    "os"
    "text/template"

    "github.com/dimakdev/HAZOP2RDF2/pkg/importer"
)

type Exporter struct {
    ReportPath string
    GraphPath  string
    AppName    string
    AppVersion string
    DateTime   string
    BaseUri    string
    Workbook   string
    Worksheets []*importer.Worksheet
}

var (
    ErrCreatingOutputFile  = errors.New("Error creating output file")
    ErrReadingTemplateFile = errors.New("Error reading template file")
    ErrWritingTemplateFile = errors.New("Error writing template file")
)

func (e *Exporter) ExportToFile(fpath, tpath string) error {
    f, err := os.Create(fpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrCreatingOutputFile, fpath, err)
    }
    defer f.Close()

    t, err := template.ParseFiles(tpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrReadingTemplateFile, tpath, err)
    }

    if err := t.Execute(f, e); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingTemplateFile, tpath, err)
    }

    return nil
}

func (e *Exporter) ExportToStdout(tpath string) error {
    t, err := template.ParseFiles(tpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrReadingTemplateFile, tpath, err)
    }

    if err := t.Execute(os.Stdout, e); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingTemplateFile, tpath, err)
    }

    return nil
}

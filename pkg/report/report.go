package report

import (
    "errors"
    "fmt"
    "html/template"
    "os"

    "github.com/dimakdev/hazop-formula/pkg/workbook"
)

type Report struct {
    ReportPath     string
    ProgramName    string
    ProgramVersion string
    DateTime       string
    Workbook       string
    Worksheets     []*workbook.Worksheet
}

var (
    ErrCreatingReportFile  = errors.New("Error creating report file")
    ErrParsingTemplateFile = errors.New("Error parsing template file")
    ErrWritingReport       = errors.New("Error writing report")
)

func (r *Report) ReportToFile(path, temp string) error {
    f, err := os.Create(path)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrCreatingReportFile, path, err)
    }
    defer f.Close()

    t, err := template.ParseFiles(temp)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrParsingTemplateFile, temp, err)
    }

    if err := t.Execute(f, r); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, temp, err)
    }

    return nil
}

func (r *Report) ReportToStdout(temp string) error {
    t, err := template.ParseFiles(temp)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrParsingTemplateFile, temp, err)
    }

    if err := t.Execute(os.Stdout, r); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, temp, err)
    }

    return nil
}

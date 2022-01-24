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

func (r *Report) ReportToFile(fpath, ftemp string) error {
    f, err := os.Create(fpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrCreatingReportFile, fpath, err)
    }
    defer f.Close()

    temp, err := template.ParseFiles(ftemp)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrParsingTemplateFile, ftemp, err)
    }

    if err := temp.Execute(f, r); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, ftemp, err)
    }

    return nil
}

func (r *Report) ReportToStdout(stemp string) error {
    temp, err := template.ParseFiles(stemp)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrParsingTemplateFile, stemp, err)
    }

    if err := temp.Execute(os.Stdout, r); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, stemp, err)
    }

    return nil
}

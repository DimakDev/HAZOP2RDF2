package workbook

import (
    "errors"
    "fmt"
    "html/template"
    "os"
)

type ReportData struct {
    Package     string
    Version     string
    DateAndTime string
    FullReport  string
    Workbook    string
    Worksheets  []*Worksheet
}

var (
    ErrCreatingFile    = errors.New("Error creating file")
    ErrParsingTemplate = errors.New("Error parsing template")
    ErrWritingReport   = errors.New("Error writing report")
)

func NewReport(fpath, ftemp, stemp string, r *ReportData) error {
    if err := reportToFile(fpath, ftemp, r); err != nil {
        return err
    }

    if err := reportToStdout(stemp, r); err != nil {
        return err
    }

    return nil
}

func reportToFile(fpath, ftemp string, r *ReportData) error {
    f, err := os.Create(fpath)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrCreatingFile, fpath, err)
    }
    defer f.Close()

    temp, err := template.ParseFiles(ftemp)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrParsingTemplate, ftemp, err)
    }

    if err := temp.Execute(f, r); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, ftemp, err)
    }

    return nil
}

func reportToStdout(stemp string, r *ReportData) error {
    temp, err := template.ParseFiles(stemp)
    if err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrParsingTemplate, stemp, err)
    }

    if err := temp.Execute(os.Stdout, r); err != nil {
        return fmt.Errorf("%v `%s`: %v", ErrWritingReport, stemp, err)
    }

    return nil
}

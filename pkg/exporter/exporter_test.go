package exporter

import (
    "os"
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestReportToFile(t *testing.T) {
    assert := assert.New(t)

    rpath := "report.txt"
    tpath := "template_file.txt"

    report := new(Report)

    var err error

    err = report.ReportToFile("", "")
    assert.Error(err)

    err = report.ReportToFile(rpath, "")
    assert.Error(err)

    err = report.ReportToFile(rpath, tpath)
    assert.Empty(err)

    err = os.Remove(rpath)
    assert.Empty(err)
}

func TestReportToStdout(t *testing.T) {
    assert := assert.New(t)

    tpath := "template_stdout.txt"

    report := new(Report)

    var err error

    err = report.ReportToStdout("")
    assert.Error(err)

    err = report.ReportToStdout(tpath)
    assert.Empty(err)
}

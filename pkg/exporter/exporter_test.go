package exporter

import (
    "os"
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestExportToFile(t *testing.T) {
    assert := assert.New(t)

    rpath := "report_file.txt"
    tpath := "report_template_long.txt"

    exp := new(Exporter)

    var err error

    err = exp.ExportToFile("", "")
    assert.Error(err)

    err = exp.ExportToFile(rpath, "")
    assert.Error(err)

    err = exp.ExportToFile(rpath, tpath)
    assert.Empty(err)

    err = os.Remove(rpath)
    assert.Empty(err)
}

func TestExportToStdout(t *testing.T) {
    assert := assert.New(t)

    tpath := "report_template_short.txt"

    exp := new(Exporter)

    var err error

    err = exp.ExportToStdout("")
    assert.Error(err)

    err = exp.ExportToStdout(tpath)
    assert.Empty(err)
}

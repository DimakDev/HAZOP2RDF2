package exporter

import (
    "os"
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestWriteToFile(t *testing.T) {
    assert := assert.New(t)

    rpath := "report_to_file.txt"
    tpath := "report_template_long.txt"

    e := new(Exporter)

    var err error

    err = e.WriteToFile("", "")
    assert.Error(err)

    err = e.WriteToFile(rpath, "")
    assert.Error(err)

    err = e.WriteToFile(rpath, tpath)
    assert.Empty(err)

    err = os.Remove(rpath)
    assert.Empty(err)
}

func TestWriteToStdout(t *testing.T) {
    assert := assert.New(t)

    tpath := "report_template_short.txt"

    e := new(Exporter)

    var err error

    err = e.WriteToStdout("")
    assert.Error(err)

    err = e.WriteToStdout(tpath)
    assert.Empty(err)
}

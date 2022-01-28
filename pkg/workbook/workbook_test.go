package workbook

import (
    "fmt"
    "io/ioutil"
    "log"
    "path/filepath"
    "strings"
    "sync"
    "testing"

    "github.com/stretchr/testify/assert"
)

func BenchmarkReadVerifyWorkbook(b *testing.B) {
    files, err := ioutil.ReadDir("data")
    if err != nil {
        log.Fatal(err)
    }

    var wg sync.WaitGroup
    for _, file := range files {
        if strings.HasSuffix(file.Name(), ".xlsx") {
            fpath := filepath.Join("data", file.Name())
            b.Run(fmt.Sprintf("Workbook: %s\n", fpath), func(b *testing.B) {
                for i := 0; i < b.N; i++ {
                    ReadVerifyWorkbook(fpath, &wg)
                }
            })
        }
    }

    wg.Wait()
}

func TestReadVerifyWorkbook(t *testing.T) {
    assert := assert.New(t)

    var wg sync.WaitGroup

    wb, err := ReadVerifyWorkbook("", &wg)
    assert.Error(err)
    assert.Empty(wb)

    files, err := ioutil.ReadDir("data")
    assert.Empty(err)
    assert.NotEmpty(files)

    for _, file := range files {
        if strings.HasSuffix(file.Name(), ".xlsx") {
            fpath := filepath.Join("data", file.Name())
            wg, err := ReadVerifyWorkbook(fpath, &wg)
            assert.Empty(err)
            assert.NotEmpty(wg)
        }
    }

    wg.Wait()
}

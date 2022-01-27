package workbook

import (
    "io/ioutil"
    "path/filepath"
    "strings"
    "sync"
    "testing"
)

func BenchmarkReadVerifyWorkbook(b *testing.B) {
    files, _ := ioutil.ReadDir("data")

    var wg sync.WaitGroup
    for i := 0; i < b.N; i++ {
        for _, file := range files {
            if strings.HasSuffix(file.Name(), ".xlsx") {
                fpath := filepath.Join("data", file.Name())
                ReadVerifyWorkbook(fpath, &wg)
            }
        }
    }

    wg.Wait()
}

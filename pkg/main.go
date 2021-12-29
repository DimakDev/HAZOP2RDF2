package main

import (
    "log"

    "github.com/dimakdev/hazop-formula-cli/pkg/workbook"
)

func main() {
    fpath := "data/HazopCrawleyGuideToBestPracticeNormalizedShort.xlsx"

    _, err := workbook.NewWorkbook(fpath)
    if err != nil {
        log.Fatal(err)
    }
}

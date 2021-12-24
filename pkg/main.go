package main

import (
    "log"

    "github.com/dimakdev/hazop-formula-cli/pkg/workbook"
)

func main() {
    datapath := "data/HazopCrawleyGuideToBestPracticeNormalized.xlsx"

    wb := workbook.New(datapath)

    if err := wb.ReadWorkbook(); err != nil {
        log.Fatal(err)
    }

    if err := wb.VerifyWorksheets(); err != nil {
        log.Fatal(err)
    }

    log.Printf("%+v", wb.Verification)
}

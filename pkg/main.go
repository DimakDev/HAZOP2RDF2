package main

import (
    "log"

    "github.com/dimakdev/hazop-formula-cli/pkg/workbook"
)

func main() {
    datapath := "data/HazopCrawleyGuideToBestPracticeNormalizedShort.xlsx"

    wb := workbook.New(datapath)

    if err := wb.ReadWorkbook(); err != nil {
        log.Fatal(err)
    }

    if err := wb.VerifyWorkbook(); err != nil {
        log.Fatal(err)
    }

    if err := wb.VerifyWorksheets(); err != nil {
        log.Fatal(err)
    }

    wb.Preview()

    // for _, ws := range wb.Worksheets {
    //     log.Printf("%+v", ws)
    // }
}

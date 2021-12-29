package main

import (
    "log"

    "github.com/dimakdev/hazop-formula-cli/pkg/workbook"
)

func main() {
    fpath := "data/HazopCrawleyGuideToBestPracticeNormalizedShort.xlsx"

    wb, err := workbook.NewWorkbook(fpath)
    if err != nil {
        log.Fatal(err)
    }

    for _, ws := range wb.Worksheets {
        log.Println(ws.HazopData.NodeMetadata)
        // log.Println(ws.HazopHeader.NodeMetadata)
        log.Println(ws.HazopData.Report.Errors)
        break
    }
}

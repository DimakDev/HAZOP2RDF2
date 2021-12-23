package main

import (
    "fmt"
    "log"
    "os"

    "github.com/jedib0t/go-pretty/v6/table"
    "github.com/jedib0t/go-pretty/v6/text"
    "github.com/xuri/excelize/v2"
)

const columnWidth = 20

func representExcelData(datapath string) error {
    f, err := excelize.OpenFile(datapath)
    if err != nil {
        return fmt.Errorf("Error opening `%s`: %v", datapath, err)
    }

    for _, name := range f.GetSheetMap() {
        log.Printf("Reading Workbook: `%s`, Worksheet `%s`\n", datapath, name)

        rows, err := f.GetRows(name)
        if err != nil {
            return fmt.Errorf("Error reading Workbook `%s`, Worksheet `%s`: %v", datapath, name, err)
        }

        if len(rows) == 0 {
            log.Printf("Workbook `%s`, Worksheet `%s` is empty\n", datapath, name)
            continue
        }

        trows := make([]table.Row, len(rows))
        tconfigs := make([]table.ColumnConfig, len(rows))
        for i, row := range rows {
            tinter := make([]interface{}, len(row))
            for j, r := range row {
                tinter[j] = r
            }
            trows[i] = tinter
            tconfigs[i] = table.ColumnConfig{Number: i + 1, WidthMax: columnWidth, WidthMaxEnforcer: text.WrapHard}
        }

        t := table.NewWriter()
        t.SetOutputMirror(os.Stdout)
        t.SetColumnConfigs(tconfigs)
        t.SetTitle(rows[0][0])
        t.AppendHeader(trows[1])
        t.AppendRows(trows[2:])
        t.AppendSeparator()
        t.SetStyle(table.StyleColoredBright)
        t.Render()
    }

    return nil
}

func main() {
    datapath := "data/HazopGuideToBestPracticeUno.xlsx"
    if err := representExcelData(datapath); err != nil {
        log.Fatal(err)
    }
}

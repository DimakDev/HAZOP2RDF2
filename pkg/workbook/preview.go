package workbook

// import (
//     "fmt"
//     "log"
//     "os"

//     "github.com/jedib0t/go-pretty/table"
//     "github.com/jedib0t/go-pretty/text"
// )

// const columnWidth = 20

// func representWorkbook(datapath string) error {

//     return nil
// }

// func main() {
//     datapath := "data/HazopGuideToBestPracticeUno.xlsx"
//     if err := representExcelData(datapath, false, true, 4); err != nil {
//         log.Fatal(err)
//     }

//     for _, name := range f.GetSheetMap() {
//         rows, err := f.GetRows(name)
//         if err != nil {
//             return fmt.Errorf("Error reading Workbook `%s`, Worksheet `%s`: %v", datapath, name, err)
//         }

//         if len(rows) == 0 {
//             log.Printf("Workbook `%s`, Worksheet `%s` is empty\n", datapath, name)
//             continue
//         }
//         log.Println(len(rows))
//         // if head {
//         //     rows = append(rows[:1], rows[1:nrows+2]...)
//         //     log.Printf("Head Workbook: `%s`, Worksheet `%s`, Rows: %d, Columns: %d\n", datapath, name, len(rows), len(rows[1]))
//         // } else if tail {
//         //     rows = append(rows[:1], rows[len(rows)-nrows:]...)
//         //     log.Printf("Tail Workbook: `%s`, Worksheet `%s`, Rows: %d, Columns: %d\n", datapath, name, len(rows), len(rows[1]))
//         // } else {
//         //     log.Printf("Reading Workbook: `%s`, Worksheet `%s`, Rows: %d, Columns: %d\n", datapath, name, len(rows), len(rows[1]))
//         // }

//         trows := make([]table.Row, len(rows))
//         for i, row := range rows {
//             tinter := make([]interface{}, len(row))
//             for j, r := range row {
//                 tinter[j] = r
//             }
//             trows[i] = tinter
//         }

//         tconfigs := make([]table.ColumnConfig, len(rows[0]))
//         for i := 0; i < len(rows[0]); i++ {
//             tconfigs[i] = table.ColumnConfig{Number: i + 1, WidthMax: columnWidth, WidthMaxEnforcer: text.WrapHard}
//         }

//         t := table.NewWriter()
//         t.SetOutputMirror(os.Stdout)
//         t.SetColumnConfigs(tconfigs)
//         t.AppendHeader(trows[0])
//         t.AppendRows(trows[1:])
//         t.AppendSeparator()
//         t.SetStyle(table.StyleColoredBright)
//         t.Render()
//     }
// }

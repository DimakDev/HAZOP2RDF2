// // package main
package workbook

// import (
//     "os"

//     "github.com/jedib0t/go-pretty/v6/table"
//     "github.com/jedib0t/go-pretty/v6/text"
// )

// const columnWidth = 20

// func newMatrix(d1, d2 int) []table.Row {
//     a := make([]interface{}, d1*d2)
//     m := make([]table.Row, d2)
//     lo, hi := 0, d1
//     for i := range m {
//         m[i] = a[lo:hi]
//         lo, hi = hi, hi+d1
//     }
//     return m
// }

// func transposeMatrix(a [][]interface{}) []table.Row {
//     b := newMatrix(len(a), len(a[0]))
//     for i := 0; i < len(b); i++ {
//         for j := 0; j < len(b[i]); j++ {
//             b[i][j] = a[j][i]
//         }
//     }
//     return b
// }

// func createConfigs(d1 int) []table.ColumnConfig {
//     configs := make([]table.ColumnConfig, d1)
//     for i := 0; i < d1; i++ {
//         configs[i] = table.ColumnConfig{
//             Number:           i + 1,
//             WidthMax:         columnWidth,
//             WidthMaxEnforcer: text.WrapHard,
//         }
//     }
//     return configs
// }

// func (wb *Workbook) Preview() {
//     for _, ws := range wb.Worksheets {
//         t := table.NewWriter()
//         t.SetOutputMirror(os.Stdout)
//         t.AppendHeader(ws.Header)
//         t.AppendRows(transposeMatrix(ws.Columns))
//         t.SetColumnConfigs(createConfigs(len(ws.Columns)))
//         // t.SetStyle(table.StyleColoredBright)
//         t.AppendSeparator()
//         t.Render()
//     }
// }

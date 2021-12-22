package data

import (
	"fmt"
	"log"
	"os"
	"reflect"

	"github.com/jedib0t/go-pretty/v6/table"
	"github.com/jedib0t/go-pretty/v6/text"
	"github.com/xuri/excelize/v2"
)

func ShowExcel() {
	f, err := excelize.OpenFile("data/HazopGuideToBestPracticeUno.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	sheetmap := f.GetSheetMap()
	// fmt.Println(sheetmap[1])

	// docprops, err := f.GetDocProps()
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }

	// fmt.Print(docprops)

	// cols, err := f.GetCols(sheetmap[1])
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }

	// for _, col := range cols {
	// 	for _, c := range col {
	// 		// fmt.Printf("%T", c)
	// 		fmt.Print(c, "\t\t")
	// 	}
	// 	fmt.Println()
	// }

	rows, err := f.GetRows(sheetmap[1])
	if err != nil {
		fmt.Println(err)
		return
	}

	// for _, row := range rows {
	// 	for _, colCell := range row {
	// 		fmt.Print(colCell, "\t")
	// 	}
	// 	fmt.Println()
	// }
	var sa interface{} = rows[1]
	log.Printf("%#v", sa)
	s := interfaceSlice(rows[1])
	// s := make([]interface{}, len(rows[1]))
	// for i, v := range rows[1] {
	// 	s[i] = v
	// }

	rr := make([]table.Row, len(rows[2:]))
	// for _, v := range rows[2:] {
	// 	for k, j := range v {
	// 		if len(j) > 20 {
	// 			v[k] = text.WrapText(j, 20)
	// 		}
	// 	}
	// }
	for i, v := range rows[2:] {
		r := interfaceSlice(v)
		rr[i] = r
	}

	t := table.NewWriter()
	t.SetOutputMirror(os.Stdout)
	t.SetColumnConfigs([]table.ColumnConfig{{Number: 1, WidthMax: 20, WidthMaxEnforcer: text.WrapHard}, {Number: 2, WidthMax: 20, WidthMaxEnforcer: text.WrapHard}, {Number: 3, WidthMax: 20, WidthMaxEnforcer: text.WrapHard}, {Number: 4, WidthMax: 20, WidthMaxEnforcer: text.WrapHard}, {Number: 5, WidthMax: 20, WidthMaxEnforcer: text.WrapHard}, {Number: 6, WidthMax: 20, WidthMaxEnforcer: text.WrapHard}, {Number: 7, WidthMax: 20, WidthMaxEnforcer: text.WrapHard}})
	// t.SetTitle(rows[0][0])
	t.AppendHeader(s)
	t.AppendRows(rr)
	t.AppendSeparator()
	// t.AppendRow([]interface{}{300, "Tyrion", "Lannister", 5000})
	// t.AppendFooter(table.Row{"", "", "Total", 10000})
	t.SetStyle(table.StyleColoredBright)
	// t.SetAllowedRowLength(100)
	t.Render()
}

func interfaceSlice(slice interface{}) []interface{} {
	s := reflect.ValueOf(slice)
	if s.Kind() != reflect.Slice {
		panic("InterfaceSlice() given a non-slice type")
	}

	// Keep the distinction between nil and empty slice input
	if s.IsNil() {
		return nil
	}

	ret := make([]interface{}, s.Len())

	for i := 0; i < s.Len(); i++ {
		ret[i] = s.Index(i).Interface()
	}

	return ret
}

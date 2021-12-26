package workbook

import (
    "errors"
    "fmt"
    "regexp"
    "strconv"
)

type verifyCellValue func(interface{}, int, int) error
type convertCellValue func(string) (interface{}, error)

type verifier struct {
    ver verifyCellValue
    con convertCellValue
    min int
    max int
    col []string
    lab string
}

func (wb *Workbook) VerifyWorkbook() error {
    for i, sname := range wb.File.GetSheetMap() {
        // Worksheet name convention: << Node Name >> - << Worksheet Type >>
        regName := regexp.MustCompile(settings.Hazop.Pattern).FindString(sname)

        if len(regName) == 0 {
            err := errors.New("Worksheet not found")
            wb.Errors[i] = append(wb.Errors[i], err)
            continue
        }

        wb.SheetMap[i] = sname
    }

    return nil
}

func (wb *Workbook) VerifyWorksheets() error {
    for i, sname := range wb.SheetMap {
        cols, err := wb.File.GetCols(sname)
        if err != nil {
            return fmt.Errorf("Error reading rows: %v", err)
        }

        if len(cols) == 0 {
            err := errors.New("Worksheet is empty")
            wb.Errors[i] = append(wb.Errors[i], err)
            continue
        }

        wb.Worksheets[i] = &Worksheet{
            // Header: make([]interface{}, len(cols)),
            // Columns: make([][]interface{}, len(cols)),
            Errors: make(map[int][]error),
        }

        if err := wb.Worksheets[i].verifyWorksheet(cols); err != nil {
            return err
        }
    }

    return nil
}

func (ws *Worksheet) verifyWorksheet(cols [][]string) error {
    for i, col := range cols {
        for _, el := range settings.Hazop.Elements {
            if regexp.MustCompile(el.Regex).MatchString(col[0]) {
                switch el.Type {
                case "string":
                    v := &verifier{
                        ver: verifyStringLength,
                        con: parseString,
                        min: el.Min,
                        max: el.Max,
                        col: col[1:],
                        lab: el.Name,
                    }
                    ws.verify(i, v)
                case "integer":
                    v := &verifier{
                        ver: verifyIntegerRange,
                        con: parseInteger,
                        min: el.Min,
                        max: el.Max,
                        lab: el.Name,
                        col: col[1:],
                    }
                    ws.verify(i, v)
                case "float":
                    v := &verifier{
                        ver: verifyFloatRange,
                        con: parseFloat,
                        min: el.Min,
                        max: el.Max,
                        lab: el.Name,
                        col: col[1:],
                    }
                    ws.verify(i, v)
                default:
                    return fmt.Errorf("Unknown element type `%s`", el.Type)
                }
                goto Next
            }
        }
        ws.Errors[i] = append(ws.Errors[i], errors.New("Column not found"))
    Next:
    }
    return nil
}

func (ws *Worksheet) verify(idx int, v *verifier) {
    if inter, err := v.run(); err != nil {
        ws.Errors[idx] = append(ws.Errors[idx], err)
    } else {
        ws.Header = append(ws.Header, v.lab)
        ws.Columns = append(ws.Columns, inter)
    }
}

func (v *verifier) run() ([]interface{}, error) {
    var inter = make([]interface{}, len(v.col))
    for i, cell := range v.col {
        val, err := v.con(cell)
        if err != nil {
            return nil, err
        }

        if err := v.ver(val, v.min, v.max); err != nil {
            return nil, err
        }

        inter[i] = val
    }
    return inter, nil
}

func parseString(val string) (interface{}, error) {
    return val, nil
}

func parseInteger(val string) (interface{}, error) {
    if v, err := strconv.Atoi(val); err != nil {
        return nil, err
    } else {
        return v, nil
    }
}

func parseFloat(val string) (interface{}, error) {
    if v, err := strconv.ParseFloat(val, 32); err == nil {
        return nil, err
    } else {
        return v, nil
    }
}

func verifyStringLength(val interface{}, min, max int) error {
    if len(val.(string)) <= min || len(val.(string)) >= max {
        return errors.New("String length is out of range")
    } else {
        return nil
    }
}

func verifyIntegerRange(val interface{}, min, max int) error {
    if val.(int) <= min || val.(int) >= max {
        return errors.New("Integer number is out of range")
    } else {
        return nil
    }
}

func verifyFloatRange(val interface{}, min, max int) error {
    if val.(float32) <= float32(min) || val.(float32) >= float32(max) {
        return errors.New("Float number is out of range")
    } else {
        return nil
    }
}

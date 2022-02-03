package workbook

import (
    "fmt"
    "log"
    "math"
    "sync"

    "github.com/xuri/excelize/v2"
)

var (
    ErrNoHeaderFound    = "Error no header found"
    ErrInvalidLength    = "Error invalid object length"
    ErrOnlyOneHeader    = "Error only one header found"
    ErrHeaderNotAligned = "Error header not aligned"
    InfoHeaderAligned   = "Info header aligned"
    ErrHeaderNotFound   = "Error header not found"
    InfoHeaderFound     = "Info header found"
    WarnMultipleCooeds  = "Warning header multiple coordinates"
    InfoValueIsValid    = "Info value parsed/verified"
)

type Workbook struct {
    File       *excelize.File
    Worksheets []*Worksheet
}

type Worksheet struct {
    Index     int
    Name      string
    NCols     int
    NRows     int
    TotalSize int
    ValidSize int
    ValidPart float64
    Data      map[int][]interface{}
    Header    map[int]string
    Elements  map[int]HazopElement
    IsValid   bool
    Logger    *Logger
}

func ReadVerifyWorkbook(fpath string, wg *sync.WaitGroup) (*Workbook, error) {
    f, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, err
    }

    var wb = &Workbook{File: f}
    sheetMap := wb.File.GetSheetMap()
    wb.Worksheets = make([]*Worksheet, len(sheetMap))

    for i, name := range sheetMap {
        wg.Add(1)

        go func(i int, name string) {
            defer wg.Done()

            cols, err := wb.File.GetCols(name)
            if err != nil {
                log.Println(err)
                return
            }

            rows, err := wb.File.GetRows(name)
            if err != nil {
                log.Println(err)
                return
            }

            ws := &Worksheet{
                Index:     i,
                Name:      name,
                NCols:     len(cols),
                NRows:     len(rows),
                TotalSize: len(cols) * len(rows),
                Data:      map[int][]interface{}{},
                Header:    map[int]string{},
                Elements:  map[int]HazopElement{},
                Logger:    &Logger{},
            }

            if err := wb.readVerifyWorksheet(ws); err != nil {
                log.Println(err)
                return
            }

            wb.Worksheets[i-1] = ws
        }(i, name)
    }

    if err := wb.File.Close(); err != nil {
        return nil, err
    }

    return wb, nil
}

func (wb *Workbook) readVerifyWorksheet(ws *Worksheet) error {
    if err := wb.searchHazopElements(ws); err != nil {
        return err
    }

    var header []string
    for _, v := range ws.Header {
        header = append(header, v)
    }
    aligned, msg, err := testHeaderAlignment(header)
    if err != nil {
        return err
    }

    if err := ws.Logger.newMessage(msg); err != nil {
        return err
    }
    if !aligned {
        ws.IsValid = false
        ws.Header = make(map[int]string)
        return nil
    }

    ws.IsValid = true

    for k, v := range ws.Header {
        coords, err := getDataCoordinates(v, ws.NRows)
        if err != nil {
            return err
        }

        tester, err := newTester(ws.Elements[k].Type)
        if err != nil {
            return err
        }

        col := make([]interface{}, len(coords))
        for i := 0; i < len(coords); i++ {
            value, err := wb.File.GetCellValue(ws.Name, coords[i])
            if err != nil {
                return err
            }

            v, err := tester.testValueType(value)
            if err != nil {
                msg := fmt.Sprintf("%v `%v`", err, coords[i])
                if err := ws.Logger.newMessage(msg); err != nil {
                    return err
                }
                continue
            }

            minLen, maxLen := ws.Elements[k].Min, ws.Elements[k].Max
            if err := tester.testValueLength(v, minLen, maxLen); err != nil {
                msg := fmt.Sprintf("%v `%v`", err, coords[i])
                if err := ws.Logger.newMessage(msg); err != nil {
                    return err
                }
                continue
            }

            msg := fmt.Sprintf("%s: `%s`", InfoValueIsValid, coords[i])
            if err := ws.Logger.newMessage(msg); err != nil {
                return err
            }

            ws.ValidSize += 1
            col[i] = v
        }

        ws.Data[k] = col
    }

    ws.ValidPart = math.Round((float64(ws.ValidSize)/float64(ws.TotalSize))*10000) / 100
    return nil
}

func (wb *Workbook) searchHazopElements(ws *Worksheet) error {
    for _, el := range Hazop.Elements {
        coords, err := wb.File.SearchSheet(ws.Name, el.Regex, true)
        if err != nil {
            return err
        }

        var msg string
        switch len(coords) {
        case 0:
            msg = fmt.Sprintf("%s `%s`", ErrHeaderNotFound, el.Name)
        case 1:
            ws.Header[el.Id] = coords[0]
            msg = fmt.Sprintf("%s `%s` %v", InfoHeaderFound, el.Name, coords)
        default:
            msg = fmt.Sprintf("%s `%s` %v", WarnMultipleCooeds, el.Name, coords)
        }
        if err := ws.Logger.newMessage(msg); err != nil {
            return err
        }
        ws.Elements[el.Id] = el
    }
    return nil
}

func getDataCoordinates(coord string, length int) ([]string, error) {
    x, y, err := excelize.CellNameToCoordinates(coord)
    if err != nil {
        return nil, err
    }

    coords := make([]string, length)
    for i := 0; i < length; i++ {
        coord, err := excelize.CoordinatesToCellName(x, y+i)
        if err != nil {
            return nil, err
        }
        coords[i] = coord
    }
    return coords, nil
}

func testHeaderAlignment(coords []string) (bool, string, error) {
    switch l := len(coords); {
    case l == 0:
        return false, ErrNoHeaderFound, nil
    case l == 1:
        return false, fmt.Sprintf("%s %v", ErrOnlyOneHeader, coords), nil
    case l > 1:
        _, yref, err := excelize.CellNameToCoordinates(coords[0])
        if err != nil {
            return false, "", err
        }
        for i := 1; i < len(coords); i++ {
            _, y, err := excelize.CellNameToCoordinates(coords[i])
            if err != nil {
                return false, "", err
            }
            if yref != y {
                return false, fmt.Sprintf("%s %v", ErrHeaderNotAligned, coords), nil
            }
        }
        return true, fmt.Sprintf("%s %v", InfoHeaderAligned, coords), nil
    default:
        return false, "", fmt.Errorf("%s %d", ErrInvalidLength, l)
    }
}

package workbook

import (
    "errors"
    "fmt"
    "log"
    "sync"

    "github.com/xuri/excelize/v2"
)

var (
    ErrOpeningExcelFile  = errors.New("Error opening Excel file")
    ErrClosingExcelFile  = errors.New("Error closing Excel file")
    ErrNoHeaderFound     = errors.New("Error no header found")
    ErrInvalidLength     = errors.New("Error invalid object length")
    ErrOnlyOneHeader     = errors.New("Error only one header found")
    ErrReadingCoords     = errors.New("Error reading coordniates")
    ErrHeaderNotAligned  = errors.New("Error header not aligned")
    ErrSearchingElements = errors.New("Error searching elements")
    ErrReadingValue      = errors.New("Error reading cell value")
    ErrReadingColumns    = errors.New("Error reading columns")
    ErrReadingRows       = errors.New("Error reading rows")
    WarnMultipleCoords   = errors.New("Warning multiple coords for one header")
    InfoHeaderAligned    = errors.New("Info header aligned")
    ErrHeaderNotFound    = errors.New("Error header not found")
    InfoHeaderFound      = errors.New("Info header found")
    InfoValueIsValid     = errors.New("Info value parsed/verified")
)

type Workbook struct {
    File       *excelize.File
    Worksheets []*Worksheet
}

type Worksheet struct {
    Index    int
    Name     string
    NCols    int
    NRows    int
    IsValid  bool
    Data     map[int][]interface{}
    Header   map[int]string
    Elements map[int]HazopElement
    Logger   *Logger
}

func ReadVerifyWorkbook(fpath string, wg *sync.WaitGroup) (*Workbook, error) {
    f, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, fmt.Errorf("%v: %v", ErrOpeningExcelFile, err)
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
                log.Println(fmt.Errorf("%v: %v", ErrReadingColumns, err))
                return
            }

            rows, err := wb.File.GetRows(name)
            if err != nil {
                log.Println(fmt.Errorf("%v: %v", ErrReadingRows, err))
                return
            }

            ws := &Worksheet{
                Index:    i,
                Name:     name,
                NCols:    len(cols),
                NRows:    len(rows),
                Data:     map[int][]interface{}{},
                Header:   map[int]string{},
                Elements: map[int]HazopElement{},
                Logger:   &Logger{},
            }

            if err := wb.readVerifyWorksheet(ws); err != nil {
                log.Println(err)
                return
            }

            wb.Worksheets[i-1] = ws
        }(i, name)
    }

    if err := wb.File.Close(); err != nil {
        return nil, fmt.Errorf("%v %v", ErrClosingExcelFile, err)
    }

    return wb, nil
}

func (wb *Workbook) readVerifyWorksheet(ws *Worksheet) error {
    if err := wb.searchHazopElements(ws); err != nil {
        return err
    }

    header := getSliceHeader(ws.Header)
    aligned, msg, err := testHeaderAlignment(header)
    if err != nil {
        return err
    }

    if err := ws.Logger.newMessage(msg.Error()); err != nil {
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
                return fmt.Errorf("%s: %v", ErrReadingValue, err)
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

            col[i] = v
        }

        ws.Data[k] = col
    }
    return nil
}

func getSliceHeader(coords map[int]string) (header []string) {
    for _, v := range coords {
        header = append(header, v)
    }
    return
}

func getDataCoordinates(coord string, length int) ([]string, error) {
    x, y, err := excelize.CellNameToCoordinates(coord)
    if err != nil {
        return nil, fmt.Errorf("%v %v", ErrReadingCoords, err)
    }

    coords := make([]string, length)
    for i := 0; i < length; i++ {
        coord, err := excelize.CoordinatesToCellName(x, y+i)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrReadingCoords, err)
        }
        coords[i] = coord
    }
    return coords, nil
}

func (wb *Workbook) searchHazopElements(ws *Worksheet) error {
    for _, el := range Hazop.Elements {
        coords, err := wb.File.SearchSheet(ws.Name, el.Regex, true)
        if err != nil {
            return fmt.Errorf("%v: %v", ErrSearchingElements, err)
        }
        coord, err := testCoordinats(coords)
        msg := fmt.Sprintf("%v `%v` %v", err, el.Name, coords)
        if err := ws.Logger.newMessage(msg); err != nil {
            return err
        }
        if coord != "" {
            ws.Header[el.Id] = coord
        }
        ws.Elements[el.Id] = el
    }
    return nil
}

func testCoordinats(coords []string) (string, error) {
    switch len(coords) {
    case 0:
        return "", ErrHeaderNotFound
    case 1:
        return coords[0], InfoHeaderFound
    default:
        return "", WarnMultipleCoords
    }
}

func testHeaderAlignment(coords []string) (bool, error, error) {
    switch l := len(coords); {
    case l == 0:
        return false, ErrNoHeaderFound, nil
    case l == 1:
        return false, fmt.Errorf("%v %v", ErrOnlyOneHeader, coords), nil
    case l > 1:
        _, yref, err := excelize.CellNameToCoordinates(coords[0])
        if err != nil {
            return false, nil, fmt.Errorf("%v %v", ErrReadingCoords, err)
        }
        for i := 1; i < len(coords); i++ {
            _, y, err := excelize.CellNameToCoordinates(coords[i])
            if err != nil {
                return false, nil, fmt.Errorf("%v %v", ErrReadingCoords, err)
            }
            if yref != y {
                return false, fmt.Errorf("%v %v", ErrHeaderNotAligned, coords), nil
            }
        }
        return true, fmt.Errorf("%v %v", InfoHeaderAligned, coords), nil
    default:
        return false, nil, fmt.Errorf("%v %d", ErrInvalidLength, l)
    }
}

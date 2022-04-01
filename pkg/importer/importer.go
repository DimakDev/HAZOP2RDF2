package importer

import (
    "fmt"
    "log"
    "math"
    "sync"

    "github.com/xuri/excelize/v2"
)

var (
    ErrNoHeaderFound    = "Error no header found"
    ErrHeaderNotAligned = "Error header not aligned"
    ErrHeaderNotFound   = "Error header not found"
    ErrHeaderMulCoords  = "Error header multiple coordinates"
    InfoHeaderAligned   = "Info header aligned"
    InfoHeaderFound     = "Info header found"
    InfoValueIsValid    = "Info value parsed/verified"
)

type Workbook struct {
    File          *excelize.File
    SheetMap      map[int]string
    Worksheets    []*Worksheet
    HazopElements map[int]HazopElement
}

type Worksheet struct {
    Index       int
    Name        string
    NCols       int
    NRows       int
    NCells      int
    NValidCells int
    PValidCells float64
    Graph       []map[string]interface{}
    GraphNRows  int
    GraphNCols  int
    Headers     map[int]string
    HeaderX     map[int]int
    HeaderY     map[int]int
    IsValid     bool
    Report      *Report
}

type Report struct {
    Warnings []string
    Errors   []string
    Info     []string
}

func (r *Report) NewWarning(msg string) {
    r.Warnings = append(r.Warnings, msg)
}

func (r *Report) NewError(msg string) {
    r.Errors = append(r.Errors, msg)
}

func (r *Report) NewInfo(msg string) {
    r.Info = append(r.Info, msg)
}

type HazopElement struct {
    Id       int    `mapstructure:"id"`
    Name     string `mapstructure:"name"`
    Regex    string `mapstructure:"regex"`
    DataType int    `mapstructure:"data_type"`
    MinLen   int    `mapstructure:"min_len"`
    MaxLen   int    `mapstructure:"max_len"`
}

type HazopElements struct {
    Elements []HazopElement `mapstructure:"elements"`
}

var Hazop HazopElements

func ImportWorkbook(fpath string) (*Workbook, error) {
    wb, err := initHazopWorkbook(fpath)
    if err != nil {
        return nil, err
    }

    if err := wb.readVerifyHazopWorkbook(); err != nil {
        return nil, err
    }

    return wb, nil
}

func initHazopWorkbook(fpath string) (*Workbook, error) {
    f, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, err
    }

    var hazopElements = make(map[int]HazopElement, len(Hazop.Elements))
    for _, e := range Hazop.Elements {
        hazopElements[e.Id] = e
    }

    var sheetMap = f.GetSheetMap()
    var wb = &Workbook{
        File:          f,
        HazopElements: hazopElements,
        SheetMap:      sheetMap,
        Worksheets:    make([]*Worksheet, len(sheetMap)),
    }

    return wb, nil
}

func (wb *Workbook) initWorksheet(i int, name string) (*Worksheet, error) {
    cols, err := wb.File.GetCols(name)
    if err != nil {
        return nil, err
    }

    rows, err := wb.File.GetRows(name)
    if err != nil {
        return nil, err
    }

    ws := &Worksheet{
        Index:  i,
        Name:   name,
        NCols:  len(cols),
        NRows:  len(rows),
        NCells: len(cols) * len(rows),
        Report: &Report{},
    }

    return ws, nil
}

func (wb *Workbook) readVerifyHazopWorkbook() error {
    var wg sync.WaitGroup

    for i, name := range wb.SheetMap {
        wg.Add(1)

        go func(i int, name string) {
            defer wg.Done()

            ws, err := wb.initWorksheet(i, name)
            if err != nil {
                log.Println(err)
                return
            }

            if err := wb.searchHazopHeaders(ws); err != nil {
                log.Println(err)
                return
            }

            if err := ws.testHeadersAlignment(); err != nil {
                log.Println(err)
                return
            }

            if ws.IsValid {
                if err := wb.readVerifyHazopData(ws); err != nil {
                    log.Println(err)
                    return
                }
            }

            wb.Worksheets[i-1] = ws
        }(i, name)
    }

    wg.Wait()

    if err := wb.File.Close(); err != nil {
        return err
    }

    return nil
}

func (wb *Workbook) searchHazopHeaders(ws *Worksheet) error {
    ws.Headers = make(map[int]string)
    for _, e := range wb.HazopElements {
        coords, err := wb.File.SearchSheet(ws.Name, e.Regex, true)
        if err != nil {
            return err
        }

        switch clen := len(coords); {
        case clen == 0:
            ws.Report.NewError(fmt.Sprintf("%s `%d:%s` %v",
                ErrHeaderNotFound,
                e.Id,
                e.Name,
                coords,
            ))
        case clen == 1:
            ws.Headers[e.Id] = coords[0]
            ws.Report.NewInfo(fmt.Sprintf("%s `%d:%s` %v",
                InfoHeaderFound,
                e.Id,
                e.Name,
                coords,
            ))
        default:
            ws.Report.NewError(fmt.Sprintf("%s `%d:%s` %v",
                ErrHeaderMulCoords,
                e.Id,
                e.Name,
                coords,
            ))
        }
    }

    return nil
}

func (ws *Worksheet) testHeadersAlignment() error {
    hLength := len(ws.Headers)
    if hLength < 2 {
        ws.IsValid = false
        ws.Report.NewError(ErrNoHeaderFound)
        return nil
    }

    var headerX = make(map[int]int, hLength)
    var headerY = make(map[int]int, hLength)
    for k, v := range ws.Headers {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return err
        }
        headerX[k] = x
        headerY[k] = y
    }

    var k0 int
    for k := range headerY {
        if k0 == 0 {
            k0 = k
            continue
        }

        if headerY[k0] != headerY[k] {
            ws.IsValid = false
            ws.Report.NewError(fmt.Sprintf("%s %v",
                ErrHeaderNotAligned,
                ws.Headers,
            ))
            return nil
        }
    }

    ws.IsValid = true
    ws.GraphNRows = ws.NRows - headerY[k0]
    ws.GraphNCols = hLength
    ws.HeaderX = headerX
    ws.HeaderY = headerY
    ws.Report.NewInfo(fmt.Sprintf("%s %v", InfoHeaderAligned, ws.Headers))

    return nil
}

func (wb *Workbook) readVerifyHazopData(ws *Worksheet) error {
    ws.Graph = make([]map[string]interface{}, ws.GraphNRows)
    for i := 0; i < ws.GraphNRows; i++ {
        ws.Graph[i] = make(map[string]interface{}, ws.GraphNCols)
    }

    // k (key) - hazop element id
    for k := range ws.Headers {
        tester, err := newTester(wb.HazopElements[k].DataType)
        if err != nil {
            return err
        }

        for i := 0; i < ws.GraphNRows; i++ {
            cellName, err := excelize.CoordinatesToCellName(
                ws.HeaderX[k],
                ws.HeaderY[k]+1+i,
            )
            if err != nil {
                return err
            }

            val, err := wb.File.GetCellValue(ws.Name, cellName)
            if err != nil {
                return err
            }

            parsed, err := tester.testCellType(val)
            if err != nil {
                ws.Report.NewError(fmt.Sprintf("%v `%v`", err, cellName))
                continue
            }

            err = tester.testCellLength(
                parsed,
                wb.HazopElements[k].MinLen,
                wb.HazopElements[k].MaxLen,
            )
            if err != nil {
                ws.Report.NewError(fmt.Sprintf("%v `%v`", err, cellName))
                continue
            }

            ws.Report.NewInfo(fmt.Sprintf("%s: `%s`",
                InfoValueIsValid,
                cellName,
            ))

            ws.NValidCells += 1
            ws.Graph[i][wb.HazopElements[k].Name] = parsed
        }
    }

    validCells := float64(ws.NValidCells) / float64(ws.NCells)
    ws.PValidCells = math.Round(validCells*10000) / 100

    return nil
}

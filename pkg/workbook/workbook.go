package workbook

import (
    "fmt"
    "log"
    "math"
    "sync"

    "github.com/xuri/excelize/v2"
)

var (
    ErrHeaderLength     = "Error invalid header length"
    ErrHeaderNotAligned = "Error header not aligned"
    ErrHeaderNotFound   = "Error header not initialized"
    InfoHeaderAligned   = "Info header aligned"
    InfoHeaderFound     = "Info header initialized"
    InfoValueIsValid    = "Info value parsed/verified"
)

type Workbook struct {
    File       *excelize.File
    Worksheets []*Worksheet
    Elements   map[int]HazopElement
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
    GraphSizeX  int
    GraphSizeY  int
    Header      map[int]string
    HeaderX     map[int]int
    HeaderY     map[int]int
    Valid       bool
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
    Id    int    `mapstructure:"id"`
    Name  string `mapstructure:"name"`
    Regex string `mapstructure:"regex"`
    Type  int    `mapstructure:"data_type"`
    Min   int    `mapstructure:"min_len"`
    Max   int    `mapstructure:"max_len"`
}

type HazopElements struct {
    Elements []HazopElement `mapstructure:"elements"`
}

var Hazop HazopElements

func ReadVerifyWorkbook(fpath string, wg *sync.WaitGroup) (*Workbook, error) {
    f, err := excelize.OpenFile(fpath)
    if err != nil {
        return nil, err
    }

    sheets := f.GetSheetMap()

    var elements = make(map[int]HazopElement, len(Hazop.Elements))
    for _, e := range Hazop.Elements {
        elements[e.Id] = e
    }

    var wb = &Workbook{
        File:       f,
        Elements:   elements,
        Worksheets: make([]*Worksheet, len(sheets)),
    }

    for i, name := range sheets {
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
                Index:  i,
                Name:   name,
                NCols:  len(cols),
                NRows:  len(rows),
                NCells: len(cols) * len(rows),
                Report: &Report{},
            }

            if err := wb.searchHazopElements(ws); err != nil {
                log.Println(err)
                return
            }

            if err := ws.checkHeaderAlignment(); err != nil {
                log.Println(err)
                return
            }

            if ws.Valid {
                if err := wb.readHazopGraph(ws); err != nil {
                    log.Println(err)
                    return
                }
            }

            wb.Worksheets[i-1] = ws
        }(i, name)
    }

    if err := wb.File.Close(); err != nil {
        return nil, err
    }

    return wb, nil
}

func (wb *Workbook) searchHazopElements(ws *Worksheet) error {
    ws.Header = make(map[int]string)
    for _, el := range wb.Elements {
        coords, err := wb.File.SearchSheet(ws.Name, el.Regex, true)
        if err != nil {
            return err
        }

        if len(coords) != 1 {
            ws.Report.NewError(fmt.Sprintf("%s `%d:%s` %v",
                ErrHeaderNotFound,
                el.Id,
                el.Name,
                coords,
            ))
            continue
        }

        ws.Header[el.Id] = coords[0]
        ws.Report.NewInfo(fmt.Sprintf("%s `%d:%s` %v",
            InfoHeaderFound,
            el.Id,
            el.Name,
            coords,
        ))
    }

    return nil
}

func (ws *Worksheet) checkHeaderAlignment() error {
    l := len(ws.Header)
    if l < 2 {
        ws.Valid = false
        ws.Report.NewError(fmt.Sprintf("%s `%d`", ErrHeaderLength, l))
        return nil
    }

    var headerX = make(map[int]int, l)
    var headerY = make(map[int]int, l)
    for k, v := range ws.Header {
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
            ws.Valid = false
            ws.Report.NewError(fmt.Sprintf("%s %v",
                ErrHeaderNotAligned,
                ws.Header,
            ))
            return nil
        }
    }

    ws.Valid = true
    ws.GraphSizeX = ws.NRows - headerY[k0]
    ws.GraphSizeY = l
    ws.HeaderX = headerX
    ws.HeaderY = headerY
    ws.Report.NewInfo(fmt.Sprintf("%s %v", InfoHeaderAligned, ws.Header))
    return nil
}

func (wb *Workbook) readHazopGraph(ws *Worksheet) error {
    ws.Graph = make([]map[string]interface{}, ws.GraphSizeX)
    for i := 0; i < ws.GraphSizeX; i++ {
        ws.Graph[i] = make(map[string]interface{}, ws.GraphSizeY)
    }

    for k := range ws.Header {
        checker, err := newChecker(wb.Elements[k].Type)
        if err != nil {
            return err
        }

        for i := 0; i < ws.GraphSizeX; i++ {
            cname, err := excelize.CoordinatesToCellName(
                ws.HeaderX[k],
                ws.HeaderY[k]+1+i,
            )
            if err != nil {
                return err
            }

            val, err := wb.File.GetCellValue(ws.Name, cname)
            if err != nil {
                return err
            }

            v, err := checker.checkValueType(val)
            if err != nil {
                ws.Report.NewError(fmt.Sprintf("%v `%v`", err, cname))
                continue
            }

            min, max := wb.Elements[k].Min, wb.Elements[k].Max
            if err := checker.checkValueLength(v, min, max); err != nil {
                ws.Report.NewError(fmt.Sprintf("%v `%v`", err, cname))
                continue
            }

            ws.Report.NewInfo(fmt.Sprintf("%s: `%s`", InfoValueIsValid, cname))

            ws.NValidCells += 1
            ws.Graph[i][wb.Elements[k].Name] = v
        }

        ws.NValidCells += 1
    }

    p := float64(ws.NValidCells) / float64(ws.NCells)
    ws.PValidCells = math.Round(p*10000) / 100

    return nil
}

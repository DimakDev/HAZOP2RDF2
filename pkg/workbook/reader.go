package workbook

import (
    "fmt"
    "log"

    "github.com/xuri/excelize/v2"
)

func (wb *Workbook) readWorkbook() error {
    for i, sname := range wb.File.GetSheetMap() {
        wb.initWorksheet(i, sname)
        if err := wb.readHazopHeader(i, sname); err != nil {
            return err
        }
        if err := wb.verifyHazopHeader(i, sname); err != nil {
            return err
        }
        if wb.HazopValidity[i].NodeMetadata {
            if err := wb.readNodeMetadata(i, sname); err != nil {
                return err
            }
        }
        if wb.HazopValidity[i].NodeHazop {
            if err := wb.readNodeHazop(i, sname); err != nil {
                return err
            }
        }
    }
    return nil
}

func (wb *Workbook) initWorksheet(i int, sname string) {
    wb.SheetMap[i] = sname
    wb.HazopData[i] = &HazopData{}
    wb.HazopHeader[i] = &HazopHeader{
        NodeMetadata: make(map[int]string),
        NodeHazop:    make(map[int]string)}
    wb.HazopValidity[i] = &HazopValidity{}
    wb.HazopDataReport[i] = &Report{}
    wb.HazopHeaderReport[i] = &Report{}
}

func (wb *Workbook) readHazopHeader(i int, sname string) error {
    for j, el := range settings.Hazop.Element {
        coord, err := wb.File.SearchSheet(sname, el.Regex, true)
        if err != nil {
            return fmt.Errorf("Error scanning elements: %v", err)
        }
        switch len(coord) {
        case 0:
            msg := fmt.Sprintf("Element not found: `%s`", el.Name)
            wb.HazopHeaderReport[i].newWarning(msg)
        case 1:
            wb.HazopHeader[i].newHeader(j, el.DataType, coord[0])
            msg := fmt.Sprintf("Element found: `%s`", el.Name)
            wb.HazopHeaderReport[i].newInfo(msg)
        default:
            msg := fmt.Sprintf("Element multiple coordinates: `%s` %v",
                el.Name, coord)
            wb.HazopHeaderReport[i].newWarning(msg)
        }
    }
    return nil
}

type argsEnvelope struct {
    cellType   int
    minLen     int
    maxLen     int
    header     string
    start      int
    fix        int
    end        int
    length     int
    sheetIndex int
    sheetName  string
}

func (wb *Workbook) readNodeMetadata(i int, sname string) error {
    cols, err := wb.File.GetCols(sname)
    if err != nil {
        return fmt.Errorf("Error reading columns: %v", err)
    }

    for k, v := range wb.HazopHeader[i].NodeMetadata {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return fmt.Errorf("Error parsing coordinate name: %v", err)
        }

        args := &argsEnvelope{
            cellType:   settings.Hazop.Element[k].CellType,
            minLen:     settings.Hazop.Element[k].MinLen,
            maxLen:     settings.Hazop.Element[k].MaxLen,
            header:     settings.Hazop.Element[k].Name,
            start:      x,
            fix:        y,
            end:        len(cols),
            length:     len(cols) - x,
            sheetIndex: i,
            sheetName:  sname,
        }

        row, err := wb.newRow(args)
        if err != nil {
            return err
        }

        wb.HazopData[i].NodeMetadata = append(wb.HazopData[i].NodeMetadata, row)
    }
    log.Println(wb.HazopData[i].NodeMetadata)
    log.Println(wb.HazopDataReport[i].Errors)
    log.Println(wb.HazopHeader[i].NodeMetadata)
    log.Println(wb.HazopHeaderReport[i].Errors)
    return nil
}

func (wb *Workbook) readNodeHazop(i int, sname string) error {
    rows, err := wb.File.GetRows(sname)
    if err != nil {
        return fmt.Errorf("Error reading rows: %v", err)
    }

    for k, v := range wb.HazopHeader[i].NodeHazop {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return fmt.Errorf("Error parsing coordinate name: %v", err)
        }

        args := &argsEnvelope{
            cellType:   settings.Hazop.Element[k].CellType,
            minLen:     settings.Hazop.Element[k].MinLen,
            maxLen:     settings.Hazop.Element[k].MaxLen,
            header:     settings.Hazop.Element[k].Name,
            start:      y,
            fix:        x,
            end:        len(rows),
            length:     len(rows) - y,
            sheetIndex: i,
            sheetName:  sname,
        }

        col, err := wb.newCol(args)
        if err != nil {
            return err
        }

        wb.HazopData[i].NodeHazop = append(wb.HazopData[i].NodeHazop, col)
    }

    log.Println(wb.HazopData[i].NodeHazop)
    log.Println(wb.HazopDataReport[i].Errors)
    log.Println(wb.HazopHeader[i].NodeHazop)
    log.Println(wb.HazopHeaderReport[i].Errors)
    return nil
}

func (wb *Workbook) newRow(args *argsEnvelope) ([]interface{}, error) {
    row := make([]interface{}, args.length)
    row[0] = args.header
    for i := 1; i < args.length; i++ {
        cname, err := excelize.CoordinatesToCellName(i+args.start, args.fix)
        if err != nil {
            return nil, fmt.Errorf("Error parsing coordniates: %v", err)
        }

        cell, cerr, serr := wb.readVerifyCell(cname, args)
        if serr != nil {
            return nil, serr
        } else if cerr != nil {
            wb.HazopDataReport[args.sheetIndex].newError(cerr)
            continue
        } else {
            row[i] = cell
        }

        row[i] = cell
    }

    return row, nil
}

func (wb *Workbook) newCol(args *argsEnvelope) ([]interface{}, error) {
    col := make([]interface{}, args.length)
    col[0] = args.header
    for i := 1; i < args.length; i++ {
        cname, err := excelize.CoordinatesToCellName(args.fix, i+args.start)
        if err != nil {
            return nil, fmt.Errorf("Error parsing coordniates: %v", err)
        }

        cell, cerr, serr := wb.readVerifyCell(cname, args)
        if serr != nil {
            return nil, serr
        } else if cerr != nil {
            wb.HazopDataReport[args.sheetIndex].newError(cerr)
            continue
        } else {
            col[i] = cell
        }
    }

    return col, nil
}

func (wb *Workbook) readVerifyCell(cname string, args *argsEnvelope) (interface{}, error, error) {
    val, err := wb.File.GetCellValue(args.sheetName, cname)
    if err != nil {
        return nil, nil, fmt.Errorf("Error reading cell value: %v", err)
    }

    verifier, err := newVerifier(args.cellType)
    if err != nil {
        return nil, nil, err
    }

    cell, err := verifier.parse(val)
    if err != nil {
        return nil, fmt.Errorf("Value `%s`: %v", cname, err), nil
    }

    if err := verifier.check(cell, args.minLen, args.maxLen); err != nil {
        return nil, fmt.Errorf("Value `%s`: %v", cname, err), nil
    }

    return cell, nil, nil
}

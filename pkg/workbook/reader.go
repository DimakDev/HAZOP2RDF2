package workbook

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

func (wb *Workbook) readHazopHeader(sname string, elements map[int]Element, n *NodeData) error {
    for k, e := range elements {
        coord, err := wb.File.SearchSheet(sname, e.Regex, true)
        if err != nil {
            return fmt.Errorf("Error scanning elements: %v", err)
        }

        switch len(coord) {
        case 0:
            n.HeaderReport.newWarning(fmt.Sprintf("Element not found: `%s`", e.Name))
        case 1:
            n.Header[k] = coord[0]
            n.HeaderReport.newInfo(fmt.Sprintf("Element found: `%s`", e.Name))
        default:
            n.HeaderReport.newWarning(fmt.Sprintf("Element multiple coordinates: `%s` %v", e.Name, coord))
        }
    }

    return nil
}

type readXYCnames func(int, int, int) ([]string, error)
type readXYCoord func(string) (int, error)

type reader struct {
    runner readXYCoord
    fixer  readXYCoord
    cnames readXYCnames
}

func (wb *Workbook) readHazopData(sname string, total int, r *reader, n *NodeData) error {
    for k, v := range n.Header {
        e := getElement(k)

        runner, err := r.runner(v)
        if err != nil {
            return err
        }

        fixer, err := r.fixer(v)
        if err != nil {
            return err
        }

        cnames, err := r.cnames(runner, fixer, total-runner)
        if err != nil {
            return err
        }

        v, err := newVerifier(e.CellType)
        if err != nil {
            return err
        }

        vec := make([]interface{}, len(cnames))
        vec[0] = e.Name

        for i := 1; i < len(cnames); i++ {
            cell, err := wb.File.GetCellValue(sname, cnames[i])
            if err != nil {
                return fmt.Errorf("Error reading cell value: %v", err)
            }

            c, err := v.parse(cell)
            if err != nil {
                n.DataReport.newError(fmt.Errorf("Value `%s` %v", cnames[i], err))
                continue
            }

            if err := v.check(c, e.MinLen, e.MaxLen); err != nil {
                n.DataReport.newError(fmt.Errorf("Value `%s` %v", cnames[i], err))
                continue
            }

            n.DataReport.newInfo(fmt.Sprintf("Value `%s` parsed and verified", cnames[i]))

            vec[i] = c
        }

        n.Data[k] = vec
    }

    return nil
}

func (wb *Workbook) getNCols(sname string) (int, error) {
    cols, err := wb.File.GetCols(sname)
    if err != nil {
        return 0, fmt.Errorf("Error reading columns: %v", err)
    }

    return len(cols), nil
}

func (wb *Workbook) getNRows(sname string) (int, error) {
    rows, err := wb.File.GetRows(sname)
    if err != nil {
        return 0, fmt.Errorf("Error reading columns: %v", err)
    }

    return len(rows), nil
}

func readXCnames(x, y, length int) ([]string, error) {
    cnames := make([]string, length)
    for i := 0; i < length; i++ {
        cname, err := excelize.CoordinatesToCellName(x+i, y)
        if err != nil {
            return nil, fmt.Errorf("Error parsing coordniates: %v", err)
        }
        cnames[i] = cname
    }
    return cnames, nil
}

func readYCnames(y, x, length int) ([]string, error) {
    cnames := make([]string, length)
    for i := 0; i < length; i++ {
        cname, err := excelize.CoordinatesToCellName(x, y+i)
        if err != nil {
            return nil, fmt.Errorf("Error parsing coordniates: %v", err)
        }
        cnames[i] = cname
    }
    return cnames, nil
}

func readXCoord(coord string) (int, error) {
    x, _, err := excelize.CellNameToCoordinates(coord)
    if err != nil {
        return 0, fmt.Errorf("Error parsing coordinate name: %v", err)
    }
    return x, nil
}

func readYCoord(coord string) (int, error) {
    _, y, err := excelize.CellNameToCoordinates(coord)
    if err != nil {
        return 0, fmt.Errorf("Error parsing coordinate name: %v", err)
    }
    return y, nil
}

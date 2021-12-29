package workbook

import (
    "fmt"
    "strconv"

    "github.com/xuri/excelize/v2"
)

func (wb *Workbook) verifyWorkbook() error {
    for _, ws := range wb.Worksheets {
        if err := ws.categorizeHazopHeader(); err != nil {
            return err
        }
        if err := ws.verifyNodeMetadataHeader(); err != nil {
            return err
        }
        if err := ws.verifyNodeHazopHeader(); err != nil {
            return err
        }
        if ws.HazopValidity.NodeMetadata {
            if err := ws.verifyNodeMetadata(); err != nil {
                return err
            }
        }
        if ws.HazopValidity.NodeHazop {
            if err := ws.verifyNodeHazop(); err != nil {
                return err
            }
        }
    }
    return nil
}

func (ws *Worksheet) categorizeHazopHeader() error {
    for i, v := range ws.HazopHeader.Raw {
        switch dtype := settings.Hazop.Element[i].DataType; dtype {
        case settings.Hazop.DataType.NodeMetadata:
            ws.HazopHeader.NodeMetadata[i] = v
        case settings.Hazop.DataType.NodeHazop:
            ws.HazopHeader.NodeHazop[i] = v
        default:
            return fmt.Errorf("Error unknown data type: %d", dtype)
        }
    }

    return nil
}

func (ws *Worksheet) verifyNodeMetadataHeader() error {
    h, err := newHeader(ws.HazopHeader.NodeMetadata)
    if err != nil {
        return nil
    }

    if !checkEqualVector(h.xs) {
        err := fmt.Errorf("Node Metadata alignment: %v", h.vs)
        ws.HazopValidity.NodeMetadata = false
        ws.HazopHeader.newError(err)
    } else {
        ws.HazopValidity.NodeMetadata = true
    }

    return nil
}

func (ws *Worksheet) verifyNodeHazopHeader() error {
    h, err := newHeader(ws.HazopHeader.NodeHazop)
    if err != nil {
        return nil
    }

    if !checkEqualVector(h.ys) {
        err := fmt.Errorf("Node Hazop alignment: %v", h.vs)
        ws.HazopValidity.NodeHazop = false
        ws.HazopHeader.newError(err)
    } else {
        ws.HazopValidity.NodeHazop = true
    }

    return nil
}

type header struct {
    ks []int
    vs []string
    xs []int
    ys []int
}

func newHeader(s []string) (*header, error) {
    var h header
    for i, v := range s {
        if len(v) != 0 {
            x, y, err := excelize.CellNameToCoordinates(v)
            if err != nil {
                return nil, fmt.Errorf("Error parsing coordinate name: %v", err)
            }
            h.ks = append(h.ks, i)
            h.vs = append(h.vs, v)
            h.xs = append(h.xs, x)
            h.ys = append(h.ys, y)
        }
    }

    return &h, nil
}

func (ws *Worksheet) verifyNodeMetadata() error {
    h, err := newHeader(ws.HazopHeader.NodeMetadata)
    if err != nil {
        return nil
    }
    ws.HazopData.NodeMetadata = make([][]interface{}, len(h.ks))
    for i := 0; i < len(h.ks); i++ {
        x, y, err := excelize.CellNameToCoordinates(h.vs[i])
        if err != nil {
            return fmt.Errorf("Error parsing coordinate name: %v", err)
        }
        vec := make([]interface{}, ws.NumOfCols-x+1)
        vec[0] = settings.Hazop.Element[h.ks[i]].Name
        for j := 0; j < ws.NumOfCols-x; j++ {
            val, cerr, serr := verifyCell(h.ks[i], ws.HazopData.Raw[x+j][y-1])
            if serr != nil {
                return serr
            } else if cerr != nil {
                cname, err := excelize.CoordinatesToCellName(x+j+1, y)
                if err != nil {
                    return fmt.Errorf("Error parsing coordniates: %v", err)
                }
                ws.HazopData.newError(fmt.Errorf("Value %s: %v", cname, cerr))
            } else {
                vec[j+1] = val
            }
        }
        ws.HazopData.NodeMetadata[i] = vec
    }
    return nil
}

func (ws *Worksheet) verifyNodeHazop() error {
    h, err := newHeader(ws.HazopHeader.NodeHazop)
    if err != nil {
        return nil
    }
    ws.HazopData.NodeHazop = make([][]interface{}, len(h.ks))
    for i := 0; i < len(h.ks); i++ {
        x, y, err := excelize.CellNameToCoordinates(h.vs[i])
        if err != nil {
            return fmt.Errorf("Error parsing coordinate name: %v", err)
        }
        vec := make([]interface{}, ws.NumOfRows-y+1)
        vec[0] = settings.Hazop.Element[h.ks[i]].Name
        for j := 0; j < ws.NumOfRows-y; j++ {
            val, cerr, serr := verifyCell(h.ks[i], ws.HazopData.Raw[x-1][y+j])
            if serr != nil {
                return serr
            } else if cerr != nil {
                cname, err := excelize.CoordinatesToCellName(x, y+j+1)
                if err != nil {
                    return fmt.Errorf("Error parsing coordniates: %v", err)
                }
                ws.HazopData.newError(fmt.Errorf("Value %s: %v", cname, cerr))
            } else {
                vec[j+1] = val
            }
        }
        ws.HazopData.NodeHazop[i] = vec
    }
    return nil
}

type checkCell func(interface{}, int, int) error
type parseCell func(string) (interface{}, error)

type cellVerifier struct {
    parse  parseCell
    check  checkCell
    value  string
    minLen int
    maxLen int
}

func newCellVerifier(idx int, val string) (*cellVerifier, error) {
    switch ctype := settings.Hazop.Element[idx].CellType; ctype {
    case settings.Hazop.CellType.String:
        return &cellVerifier{
            parse:  parseStr,
            check:  checkStrLen,
            value:  val,
            minLen: settings.Hazop.Element[idx].MinLen,
            maxLen: settings.Hazop.Element[idx].MaxLen,
        }, nil
    case settings.Hazop.CellType.Integer:
        return &cellVerifier{
            parse:  parseInt,
            check:  checkIntRange,
            value:  val,
            minLen: settings.Hazop.Element[idx].MinLen,
            maxLen: settings.Hazop.Element[idx].MaxLen,
        }, nil
    case settings.Hazop.CellType.Float:
        return &cellVerifier{
            parse:  parseFloat,
            check:  checkFloatRange,
            value:  val,
            minLen: settings.Hazop.Element[idx].MinLen,
            maxLen: settings.Hazop.Element[idx].MaxLen,
        }, nil
    default:
        return nil, fmt.Errorf("Unknown cell type: %d", ctype)
    }
}

func verifyCell(k int, val string) (interface{}, error, error) {
    ver, err := newCellVerifier(k, val)
    if err != nil {
        return nil, nil, err
    }
    cell, err := ver.parse(ver.value)
    if err != nil {
        return nil, err, nil
    }
    if err := ver.check(cell, ver.minLen, ver.maxLen); err != nil {
        return nil, err, nil
    }
    return cell, nil, nil
}

func checkEqualVector(a []int) bool {
    for i := 1; i < len(a); i++ {
        if a[0] != a[i] {
            return false
        }
    }
    return true
}

func parseStr(val string) (interface{}, error) {
    return val, nil
}

func parseInt(val string) (interface{}, error) {
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

func checkStrLen(val interface{}, min, max int) error {
    if len(val.(string)) < min || len(val.(string)) > max {
        return fmt.Errorf("out of range %d-%d", min, max)
    } else {
        return nil
    }
}

func checkIntRange(val interface{}, min, max int) error {
    if val.(int) < min || val.(int) > max {
        return fmt.Errorf("out of range %d-%d", min, max)
    } else {
        return nil
    }
}

func checkFloatRange(val interface{}, min, max int) error {
    if val.(float32) < float32(min) || val.(float32) > float32(max) {
        return fmt.Errorf("out of range %d-%d", min, max)
    } else {
        return nil
    }
}

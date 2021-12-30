package workbook

import (
    "fmt"
    "strconv"

    "github.com/xuri/excelize/v2"
)

func (wb *Workbook) verifyHazopHeader(i int, _ string) error {
    nm, err := newHeader(wb.HazopHeader[i].NodeMetadata)
    if err != nil {
        return nil
    }

    if !checkEqualVector(nm.xs) {
        err := fmt.Errorf("Node Metadata alignment: %v", nm.vs)
        wb.HazopValidity[i].NodeMetadata = false
        wb.HazopHeaderReport[i].newError(err)
    } else {
        wb.HazopValidity[i].NodeMetadata = true
    }

    nh, err := newHeader(wb.HazopHeader[i].NodeHazop)
    if err != nil {
        return nil
    }

    if !checkEqualVector(nh.ys) {
        err := fmt.Errorf("Node Hazop alignment: %v", nh.vs)
        wb.HazopValidity[i].NodeHazop = false
        wb.HazopHeaderReport[i].newError(err)
    } else {
        wb.HazopValidity[i].NodeHazop = true
    }

    return nil
}

func checkEqualVector(a []int) bool {
    for i := 1; i < len(a); i++ {
        if a[0] != a[i] {
            return false
        }
    }
    return true
}

type header struct {
    ks []int
    vs []string
    xs []int
    ys []int
}

func newHeader(m map[int]string) (*header, error) {
    var h header
    for k, v := range m {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return nil, fmt.Errorf("Error parsing coordinate name: %v", err)
        }
        h.ks = append(h.ks, k)
        h.vs = append(h.vs, v)
        h.xs = append(h.xs, x)
        h.ys = append(h.ys, y)
    }

    return &h, nil
}

type checker func(interface{}, int, int) error
type parser func(string) (interface{}, error)

type verifier struct {
    parse parser
    check checker
}

func newVerifier(ctype int) (*verifier, error) {
    switch ctype {
    case settings.Hazop.CellType.String:
        return &verifier{
            parse: parseStr,
            check: checkStrLen,
        }, nil
    case settings.Hazop.CellType.Integer:
        return &verifier{
            parse: parseInt,
            check: checkIntRange,
        }, nil
    case settings.Hazop.CellType.Float:
        return &verifier{
            parse: parseFloat,
            check: checkFloatRange,
        }, nil
    default:
        return nil, fmt.Errorf("Unknown cell type: %d", ctype)
    }
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

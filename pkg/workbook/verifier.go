package workbook

import (
    "errors"
    "fmt"
    "strconv"

    "github.com/xuri/excelize/v2"
)

func verifyHeaderAlignment(header []int, headerNames []string, n *NodeData) {
    if len(header) == 0 {
        n.HeaderAligned = false
        n.HeaderReport.newError(errors.New("No header found"))
        return
    }

    if !checkEqualVector(header) {
        n.HeaderAligned = false
        n.HeaderReport.newError(fmt.Errorf("Header not aligned: %v", headerNames))
    } else {
        n.HeaderAligned = true
        n.HeaderReport.newInfo(fmt.Sprintf("Header aligned: %v", headerNames))
    }
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
    keys   []int
    coords []string
    coordX []int
    coordY []int
}

func splitHeader(coord map[int]string) (*header, error) {
    var h header
    for k, v := range coord {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return nil, fmt.Errorf("Error parsing coordinate name: %v", err)
        }
        h.keys = append(h.keys, k)
        h.coords = append(h.coords, v)
        h.coordX = append(h.coordX, x)
        h.coordY = append(h.coordY, y)
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

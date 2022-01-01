package workbook

import (
    "errors"
    "fmt"
    "strconv"

    "github.com/xuri/excelize/v2"
)

var (
    ErrNoHeaderFound    = errors.New("No valid header found")
    ErrNotEnoughHeader  = errors.New("Not enough header to align")
    ErrHeaderNotAligned = errors.New("Header not aligned")
    ErrParsingCellNames = errors.New("Error parsing cell names")
    ErrUnknownCellType  = errors.New("Uknown cell type")
    ErrParsingInteger   = errors.New("Failed parsing integer")
    ErrParsingFloat     = errors.New("Failed parsing float")
    ErrValueOutOfRange  = errors.New("Value out of range")
    HeaderAligned       = "Header aligned"
)

func verifyHeaderAlignment(coord []int, cnames []string, n *NodeData) {
    if len(coord) == 0 {
        n.HeaderAligned = false
        n.HeaderLogger.newError(ErrNoHeaderFound.Error())
        return
    }

    if len(coord) == 1 {
        n.HeaderAligned = false
        msg := fmt.Sprintf("%v: %v", ErrNotEnoughHeader, cnames)
        n.HeaderLogger.newError(msg)
        return
    }

    if !checkEqualVector(coord) {
        n.HeaderAligned = false
        msg := fmt.Sprintf("%v: %v", ErrHeaderNotAligned, cnames)
        n.HeaderLogger.newError(msg)
    } else {
        n.HeaderAligned = true
        msg := fmt.Sprintf("%s: %v", HeaderAligned, cnames)
        n.HeaderLogger.newInfo(msg)
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

func splitHeader(cnames map[int]string) (*header, error) {
    var h header
    for k, v := range cnames {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrParsingCellNames, err)
        }
        h.keys = append(h.keys, k)
        h.coords = append(h.coords, v)
        h.coordX = append(h.coordX, x)
        h.coordY = append(h.coordY, y)
    }

    return &h, nil
}

type parser func(string) (interface{}, error)
type checker func(interface{}, int, int) error

type verifier struct {
    parse parser
    check checker
}

func newVerifier(ctype int) (*verifier, error) {
    switch ctype {
    case Hazop.CellType.String:
        return &verifier{
            parse: parseStr,
            check: checkStrLen,
        }, nil
    case Hazop.CellType.Integer:
        return &verifier{
            parse: parseInt,
            check: checkIntRange,
        }, nil
    case Hazop.CellType.Float:
        return &verifier{
            parse: parseFloat,
            check: checkFloatRange,
        }, nil
    default:
        return nil, fmt.Errorf("%v: %d", ErrUnknownCellType, ctype)
    }
}

func parseStr(val string) (interface{}, error) {
    return val, nil
}

func parseInt(val string) (interface{}, error) {
    if v, err := strconv.Atoi(val); err != nil {
        return nil, ErrParsingInteger
    } else {
        return v, nil
    }
}

func parseFloat(val string) (interface{}, error) {
    if v, err := strconv.ParseFloat(val, 32); err == nil {
        return nil, ErrParsingFloat
    } else {
        return v, nil
    }
}

func checkStrLen(val interface{}, min, max int) error {
    if len(val.(string)) < min || len(val.(string)) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func checkIntRange(val interface{}, min, max int) error {
    if val.(int) < min || val.(int) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func checkFloatRange(val interface{}, min, max int) error {
    if val.(float32) < float32(min) || val.(float32) > float32(max) {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

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
        n.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrNotEnoughHeader,
                cnames,
            ),
        )
        return
    }

    if !checkHeaderAlignment(coord) {
        n.HeaderAligned = false
        n.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrHeaderNotAligned,
                cnames,
            ),
        )
        return
    }

    n.HeaderAligned = true
    n.HeaderLogger.newInfo(fmt.Sprintf("%s: %v", HeaderAligned, cnames))
}

func checkHeaderAlignment(index []int) bool {
    for i := 1; i < len(index); i++ {
        if index[0] != index[i] {
            return false
        }
    }
    return true
}

type coordinates struct {
    keys   []int
    cnames []string
    coordX []int
    coordY []int
}

func cellNamesToCoordinates(cnames map[int]string) (*coordinates, error) {
    var coords coordinates
    for k, v := range cnames {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrParsingCellNames, err)
        }
        coords.keys = append(coords.keys, k)
        coords.cnames = append(coords.cnames, v)
        coords.coordX = append(coords.coordX, x)
        coords.coordY = append(coords.coordY, y)
    }

    return &coords, nil
}

type cellVerifier interface {
    checkCellType(string) (interface{}, error)
    checkCellLength(interface{}, int, int) error
}

type verifyString struct{}
type verifyFloat struct{}
type verifyInteger struct{}

func (v verifyString) checkCellType(val string) (interface{}, error) {
    return val, nil
}

func (v verifyString) checkCellLength(val interface{}, min, max int) error {
    if len(val.(string)) < min || len(val.(string)) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifyInteger) checkCellType(val string) (interface{}, error) {
    if v, err := strconv.Atoi(val); err != nil {
        return nil, ErrParsingInteger
    } else {
        return v, nil
    }
}

func (v verifyInteger) checkCellLength(val interface{}, min, max int) error {
    if val.(int) < min || val.(int) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifyFloat) checkCellType(val string) (interface{}, error) {
    if v, err := strconv.ParseFloat(val, 32); err == nil {
        return nil, ErrParsingFloat
    } else {
        return v, nil
    }
}

func (v verifyFloat) checkCellLength(val interface{}, min, max int) error {
    if val.(float32) < float32(min) || val.(float32) > float32(max) {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

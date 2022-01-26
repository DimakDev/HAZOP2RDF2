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

func verifyHeaderAlignment(coords []int, cnames []string, node *NodeData) {
    if len(coords) == 0 {
        node.HeaderAligned = false
        node.HeaderLogger.newError(ErrNoHeaderFound.Error())
        return
    }

    if len(coords) == 1 {
        node.HeaderAligned = false
        node.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrNotEnoughHeader,
                cnames,
            ),
        )
        return
    }

    if !checkHeaderAlignment(coords) {
        node.HeaderAligned = false
        node.HeaderLogger.newError(
            fmt.Sprintf("%v: %v",
                ErrHeaderNotAligned,
                cnames,
            ),
        )
        return
    }

    node.HeaderAligned = true
    node.HeaderLogger.newInfo(fmt.Sprintf("%s: %v", HeaderAligned, cnames))
}

func checkHeaderAlignment(coords []int) bool {
    for i := 1; i < len(coords); i++ {
        if coords[0] != coords[i] {
            return false
        }
    }
    return true
}

type coordinates struct {
    elkeys  []int
    cnames  []string
    coordsX []int
    coordsY []int
}

func cellNamesToCoordinates(cnames map[int]string) (*coordinates, error) {
    var coords coordinates
    for k, v := range cnames {
        x, y, err := excelize.CellNameToCoordinates(v)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrParsingCellNames, err)
        }
        coords.elkeys = append(coords.elkeys, k)
        coords.cnames = append(coords.cnames, v)
        coords.coordsX = append(coords.coordsX, x)
        coords.coordsY = append(coords.coordsY, y)
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

func (v verifyString) checkCellType(cell string) (interface{}, error) {
    return cell, nil
}

func (v verifyString) checkCellLength(cell interface{}, min, max int) error {
    if len(cell.(string)) < min || len(cell.(string)) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifyInteger) checkCellType(cell string) (interface{}, error) {
    if v, err := strconv.Atoi(cell); err != nil {
        return nil, ErrParsingInteger
    } else {
        return v, nil
    }
}

func (v verifyInteger) checkCellLength(cell interface{}, min, max int) error {
    if cell.(int) < min || cell.(int) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifyFloat) checkCellType(cell string) (interface{}, error) {
    if v, err := strconv.ParseFloat(cell, 32); err == nil {
        return nil, ErrParsingFloat
    } else {
        return v, nil
    }
}

func (v verifyFloat) checkCellLength(cell interface{}, min, max int) error {
    if cell.(float32) < float32(min) || cell.(float32) > float32(max) {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

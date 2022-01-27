package workbook

import (
    "errors"
    "fmt"
    "strconv"
)

var (
    ErrParsingInteger  = errors.New("Error parsing integer")
    ErrParsingFloat    = errors.New("Error parsing float")
    ErrValueOutOfRange = errors.New("Error value out of range")
)

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

func (v verifyInteger) checkCellType(cell string) (interface{}, error) {
    if v, err := strconv.Atoi(cell); err != nil {
        return nil, ErrParsingInteger
    } else {
        return v, nil
    }
}

func (v verifyFloat) checkCellType(cell string) (interface{}, error) {
    if v, err := strconv.ParseFloat(cell, 32); err == nil {
        return nil, ErrParsingFloat
    } else {
        return v, nil
    }
}

func (v verifyString) checkCellLength(cell interface{}, min, max int) error {
    if len(cell.(string)) < min || len(cell.(string)) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifyInteger) checkCellLength(cell interface{}, min, max int) error {
    if cell.(int) < min || cell.(int) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifyFloat) checkCellLength(cell interface{}, min, max int) error {
    if cell.(float32) < float32(min) || cell.(float32) > float32(max) {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

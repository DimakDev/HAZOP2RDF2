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

type verifier interface {
    checkCellType(string) (interface{}, error)
    checkCellLength(interface{}, int, int) error
}

type verifierString struct{}
type verifierFloat struct{}
type verifierInteger struct{}

func newCellVerifier(ctype int) (verifier, error) {
    switch ctype {
    case Hazop.CellType.String:
        return verifierString{}, nil
    case Hazop.CellType.Integer:
        return verifierInteger{}, nil
    case Hazop.CellType.Float:
        return verifierFloat{}, nil
    default:
        return nil, fmt.Errorf("%v: %d", ErrUnknownCellType, ctype)
    }
}

func (v verifierString) checkCellType(cell string) (interface{}, error) {
    return cell, nil
}

func (v verifierInteger) checkCellType(cell string) (interface{}, error) {
    if v, err := strconv.Atoi(cell); err != nil {
        return nil, ErrParsingInteger
    } else {
        return v, nil
    }
}

func (v verifierFloat) checkCellType(cell string) (interface{}, error) {
    if v, err := strconv.ParseFloat(cell, 32); err == nil {
        return nil, ErrParsingFloat
    } else {
        return v, nil
    }
}

func (v verifierString) checkCellLength(cell interface{}, min, max int) error {
    if len(cell.(string)) < min || len(cell.(string)) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifierInteger) checkCellLength(cell interface{}, min, max int) error {
    if cell.(int) < min || cell.(int) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v verifierFloat) checkCellLength(cell interface{}, min, max int) error {
    if cell.(float32) < float32(min) || cell.(float32) > float32(max) {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

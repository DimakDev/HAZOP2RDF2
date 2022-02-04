package workbook

import (
    "fmt"
    "strconv"
)

const (
    STRING int = iota
    INTEGER
    FLOAT
)

var (
    ErrParsingInteger  = "Error parsing integer"
    ErrParsingFloat    = "Error parsing float"
    ErrValueOutOfRange = "Error value out of range"
    ErrUnknownDatatype = "Error unknown data type"
)

type checker interface {
    checkValueType(interface{}) (interface{}, error)
    checkValueLength(interface{}, int, int) error
}

type checkString struct{}
type checkFloat struct{}
type checkInteger struct{}

func newChecker(cell int) (checker, error) {
    switch cell {
    case STRING:
        return checkString{}, nil
    case INTEGER:
        return checkInteger{}, nil
    case FLOAT:
        return checkFloat{}, nil
    default:
        return nil, fmt.Errorf("%s: %d", ErrUnknownDatatype, cell)
    }
}

func (v checkString) checkValueType(value interface{}) (interface{}, error) {
    return value.(string), nil
}

func (c checkInteger) checkValueType(value interface{}) (interface{}, error) {
    if v, err := strconv.Atoi(value.(string)); err != nil {
        return nil, fmt.Errorf(ErrParsingInteger)
    } else {
        return v, nil
    }
}

func (c checkFloat) checkValueType(value interface{}) (interface{}, error) {
    if v, err := strconv.ParseFloat(value.(string), 32); err != nil {
        return nil, fmt.Errorf(ErrParsingFloat)
    } else {
        return v, nil
    }
}

func (c checkString) checkValueLength(value interface{}, min, max int) error {
    if len(value.(string)) < min || len(value.(string)) > max {
        return fmt.Errorf("%s %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (c checkInteger) checkValueLength(value interface{}, min, max int) error {
    if value.(int) < min || value.(int) > max {
        return fmt.Errorf("%s %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (c checkFloat) checkValueLength(value interface{}, min, max int) error {
    if value.(float32) < float32(min) || value.(float32) > float32(max) {
        return fmt.Errorf("%s %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

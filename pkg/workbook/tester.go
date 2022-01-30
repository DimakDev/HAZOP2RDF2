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
    ErrUnknownDatatype = errors.New("Error unknown data type")
)

type tester interface {
    testValueType(interface{}) (interface{}, error)
    testValueLength(interface{}, int, int) error
}

type testString struct{}
type testFloat struct{}
type testInteger struct{}

func newTester(dtype datatype) (tester, error) {
    switch dtype {
    case STRING:
        return testString{}, nil
    case INTEGER:
        return testInteger{}, nil
    case FLOAT:
        return testFloat{}, nil
    default:
        return nil, fmt.Errorf("%v: %d", ErrUnknownDatatype, dtype)
    }
}

func (v testString) testValueType(value interface{}) (interface{}, error) {
    return value.(string), nil
}

func (v testInteger) testValueType(value interface{}) (interface{}, error) {
    if v, err := strconv.Atoi(value.(string)); err != nil {
        return nil, fmt.Errorf("%v", ErrParsingInteger)
    } else {
        return v, nil
    }
}

func (v testFloat) testValueType(value interface{}) (interface{}, error) {
    if v, err := strconv.ParseFloat(value.(string), 32); err != nil {
        return nil, fmt.Errorf("%v", ErrParsingFloat)
    } else {
        return v, nil
    }
}

func (v testString) testValueLength(value interface{}, min, max int) error {
    if len(value.(string)) < min || len(value.(string)) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v testInteger) testValueLength(value interface{}, min, max int) error {
    if value.(int) < min || value.(int) > max {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (v testFloat) testValueLength(value interface{}, min, max int) error {
    if value.(float32) < float32(min) || value.(float32) > float32(max) {
        return fmt.Errorf("%v %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

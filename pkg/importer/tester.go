package importer

import (
    "fmt"
    "strconv"
)

var (
    ErrParsingInteger  = "Error parsing integer"
    ErrParsingFloat    = "Error parsing float"
    ErrValueOutOfRange = "Error value out of range"
    ErrUnknownDatatype = "Error unknown data type"
)

type tester interface {
    testCellType(interface{}) (interface{}, error)
    testCellLength(interface{}, int, int) error
}

type testString struct{}
type testFloat struct{}
type testInteger struct{}

// 0 - STRING, 1 - INTEGER, 2 - FLOAT
func newTester(dataType int) (tester, error) {
    switch dataType {
    case 0:
        return testString{}, nil
    case 1:
        return testInteger{}, nil
    case 2:
        return testFloat{}, nil
    default:
        return nil, fmt.Errorf("%s: %d", ErrUnknownDatatype, dataType)
    }
}

func (v testString) testCellType(value interface{}) (interface{}, error) {
    return value.(string), nil
}

func (c testInteger) testCellType(value interface{}) (interface{}, error) {
    if v, err := strconv.Atoi(value.(string)); err != nil {
        return nil, fmt.Errorf(ErrParsingInteger)
    } else {
        return v, nil
    }
}

func (c testFloat) testCellType(value interface{}) (interface{}, error) {
    if v, err := strconv.ParseFloat(value.(string), 32); err != nil {
        return nil, fmt.Errorf(ErrParsingFloat)
    } else {
        return v, nil
    }
}

func (c testString) testCellLength(value interface{}, min, max int) error {
    if len(value.(string)) < min || len(value.(string)) > max {
        return fmt.Errorf("%s %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (c testInteger) testCellLength(value interface{}, min, max int) error {
    if value.(int) < min || value.(int) > max {
        return fmt.Errorf("%s %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

func (c testFloat) testCellLength(value interface{}, min, max int) error {
    if value.(float32) < float32(min) || value.(float32) > float32(max) {
        return fmt.Errorf("%s %d-%d", ErrValueOutOfRange, min, max)
    } else {
        return nil
    }
}

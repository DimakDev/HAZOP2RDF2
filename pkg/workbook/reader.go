package workbook

import (
    "errors"
    "fmt"

    "github.com/xuri/excelize/v2"
)

var (
    ErrParsingCoordinate     = errors.New("Error parsing coordniate")
    ErrParsingCoordinateName = errors.New("Error parsing coordinate name")
    ErrParsingCellNames      = errors.New("Error parsing cell names")
)

type readXYCellNames func(int, int, int) ([]string, error)
type readXYCoordinate func(string) (int, error)

type reader struct {
    varDimension readXYCoordinate
    fixDimension readXYCoordinate
    cellNames    readXYCellNames
}

func (r *reader) readCellNames(start string, end int) ([]string, error) {
    d1, err := r.varDimension(start)
    if err != nil {
        return nil, err
    }

    d2, err := r.fixDimension(start)
    if err != nil {
        return nil, err
    }

    cnames, err := r.cellNames(d1, d2, end-d1)
    if err != nil {
        return nil, err
    }

    return cnames, nil
}

func readXCellNames(x, y, size int) ([]string, error) {
    cnames := make([]string, size)
    for i := 0; i < size; i++ {
        cname, err := excelize.CoordinatesToCellName(x+i, y)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrParsingCoordinate, err)
        }
        cnames[i] = cname
    }

    return cnames, nil
}

func readYCellNames(y, x, size int) ([]string, error) {
    cnames := make([]string, size)
    for i := 0; i < size; i++ {
        cname, err := excelize.CoordinatesToCellName(x, y+i)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrParsingCoordinate, err)
        }
        cnames[i] = cname
    }

    return cnames, nil
}

func readXCoordinate(cname string) (int, error) {
    x, _, err := excelize.CellNameToCoordinates(cname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrParsingCoordinateName, err)
    }

    return x, nil
}

func readYCoordinate(cname string) (int, error) {
    _, y, err := excelize.CellNameToCoordinates(cname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrParsingCoordinateName, err)
    }

    return y, nil
}

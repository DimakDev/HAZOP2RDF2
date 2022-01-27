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

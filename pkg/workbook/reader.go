package workbook

import (
    "errors"
    "fmt"

    "github.com/xuri/excelize/v2"
)

var (
    ErrScanningHeader         = errors.New("Error scanning header")
    ErrReadingCellValue       = errors.New("Error reading cell value")
    ErrReadingColumns         = errors.New("Error reading columns")
    ErrReadingRows            = errors.New("Error reading rows")
    ErrParsingCoordinates     = errors.New("Error parsing coordniates")
    ErrParsingCoordinateName  = errors.New("Error parsing coordinate name")
    HeaderNotFound            = "Header not found"
    HeaderFound               = "Header found"
    HeaderMultipleCoordinates = "Header multiple coordinates found"
    ValueParsedVerified       = "Value parsed/verified"
)

func (wb *Workbook) readHazopElements(
    sname string,
    elements map[int]Element,
    node *NodeData,
) error {
    for i, element := range elements {
        coords, err := wb.File.SearchSheet(sname, element.Regex, true)
        if err != nil {
            return fmt.Errorf("%v: %v", ErrScanningHeader, err)
        }

        switch len(coords) {
        case 0:
            node.HeaderLogger.newWarning(
                fmt.Sprintf("%s: `%s`",
                    HeaderNotFound,
                    element.Name,
                ),
            )
        case 1:
            node.Header[i], node.Element[i] = coords[0], element
            node.HeaderLogger.newInfo(
                fmt.Sprintf("%s: `%s` `%s`",
                    HeaderFound,
                    element.Name,
                    coords[0],
                ),
            )
        default:
            node.HeaderLogger.newWarning(
                fmt.Sprintf("%v: `%s` %v",
                    HeaderMultipleCoordinates,
                    element.Name,
                    coords,
                ),
            )
        }
    }

    return nil
}

func (wb *Workbook) getNCols(sname string) (int, error) {
    cols, err := wb.File.GetCols(sname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrReadingColumns, err)
    }

    return len(cols), nil
}

func (wb *Workbook) getNRows(sname string) (int, error) {
    rows, err := wb.File.GetRows(sname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrReadingRows, err)
    }

    return len(rows), nil
}

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
            return nil, fmt.Errorf("%v: %v", ErrParsingCoordinates, err)
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
            return nil, fmt.Errorf("%v: %v", ErrParsingCoordinates, err)
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

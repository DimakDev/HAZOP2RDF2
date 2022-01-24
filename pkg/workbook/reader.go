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
    n *NodeData,
) error {
    for k, e := range elements {
        coord, err := wb.File.SearchSheet(sname, e.Regex, true)
        if err != nil {
            return fmt.Errorf("%v: %v", ErrScanningHeader, err)
        }

        switch len(coord) {
        case 0:
            n.HeaderLogger.newWarning(
                fmt.Sprintf("%s: `%s`",
                    HeaderNotFound,
                    e.Name,
                ),
            )
        case 1:
            n.Header[k], n.Element[k] = coord[0], e
            n.HeaderLogger.newInfo(
                fmt.Sprintf("%s: `%s` `%s`",
                    HeaderFound,
                    e.Name,
                    coord[0],
                ),
            )
        default:
            n.HeaderLogger.newWarning(
                fmt.Sprintf("%v: `%s` %v",
                    HeaderMultipleCoordinates,
                    e.Name,
                    coord,
                ),
            )
        }
    }

    return nil
}

type readXYCellNames func(int, int, int) ([]string, error)
type readXYCoord func(string) (int, error)

type reader struct {
    varDimension readXYCoord
    fixDimension readXYCoord
    cellNames    readXYCellNames
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

func readXCellNames(x, y, length int) ([]string, error) {
    cnames := make([]string, length)
    for i := 0; i < length; i++ {
        cname, err := excelize.CoordinatesToCellName(x+i, y)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrParsingCoordinates, err)
        }
        cnames[i] = cname
    }

    return cnames, nil
}

func readYCellNames(y, x, length int) ([]string, error) {
    cnames := make([]string, length)
    for i := 0; i < length; i++ {
        cname, err := excelize.CoordinatesToCellName(x, y+i)
        if err != nil {
            return nil, fmt.Errorf("%v: %v", ErrParsingCoordinates, err)
        }
        cnames[i] = cname
    }

    return cnames, nil
}

func readXCoord(cname string) (int, error) {
    x, _, err := excelize.CellNameToCoordinates(cname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrParsingCoordinateName, err)
    }

    return x, nil
}

func readYCoord(cname string) (int, error) {
    _, y, err := excelize.CellNameToCoordinates(cname)
    if err != nil {
        return 0, fmt.Errorf("%v: %v", ErrParsingCoordinateName, err)
    }

    return y, nil
}

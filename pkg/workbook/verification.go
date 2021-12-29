package workbook

import (
    "errors"
    "fmt"
    "strconv"

    "github.com/xuri/excelize/v2"
)

func (wb *Workbook) verifyWorkbook() error {
    for _, ws := range wb.Worksheets {
        if err := ws.categorizeHazopHeader(); err != nil {
            return err
        }
        if err := ws.verifyHazopHeaderNodeMetadata(); err != nil {
            return err
        }
        if err := ws.verifyHazopHeaderNodeHazop(); err != nil {
            return err
        }
        if ws.HazopValidity.NodeMetadata {
            if err := ws.verifyHazopDataNodeMetadata(); err != nil {
                return err
            }
        }
        // if ws.HazopValidity.NodeHazop {
        //     if err := ws.verifyHazopDataNodeHazop(); err != nil {
        //         return err
        //     }
        // }
    }
    return nil
}

func (ws *Worksheet) categorizeHazopHeader() error {
    for j, hhr := range ws.HazopHeader.Raw {
        switch dtype := settings.Hazop.Element[j].DataType; dtype {
        case settings.Hazop.DataType.NodeMetadata:
            ws.HazopHeader.NodeMetadata[j] = hhr
        case settings.Hazop.DataType.NodeHazop:
            ws.HazopHeader.NodeHazop[j] = hhr
        default:
            return fmt.Errorf("Error unknown data type: %d", dtype)
        }
    }
    return nil
}

func (ws *Worksheet) verifyHazopHeaderNodeMetadata() error {
    coord := getMapValues(ws.HazopHeader.NodeMetadata)
    vX, _, err := cellsNameToCoordinates(coord)
    if err != nil {
        return err
    }

    if !checkEqualVector(vX) {
        err := fmt.Errorf("Node Metadata header alignment: `%s`", coord)
        ws.HazopValidity.NodeMetadata = false
        ws.HazopHeader.newError(err)
    } else {
        ws.HazopValidity.NodeMetadata = true
    }

    return nil
}

func (ws *Worksheet) verifyHazopHeaderNodeHazop() error {
    coord := getMapValues(ws.HazopHeader.NodeHazop)
    _, vY, err := cellsNameToCoordinates(coord)
    if err != nil {
        return err
    }

    if !checkEqualVector(vY) {
        err := fmt.Errorf("Node Hazop header alignment: `%s`", coord)
        ws.HazopValidity.NodeHazop = false
        ws.HazopHeader.newError(err)
    } else {
        ws.HazopValidity.NodeHazop = true
    }

    return nil
}

func (ws *Worksheet) verifyHazopDataNodeMetadata() error {
    elements := getMapKeys(ws.HazopHeader.NodeMetadata)
    coord := getMapValues(ws.HazopHeader.NodeMetadata)

    ws.HazopData.NodeMetadata = make([][]interface{}, len(elements))
    for i := 0; i < len(elements); i++ {
        x, y, err := cellNameToCoordinates(coord[i])
        if err != nil {
            return err
        }
        ws.HazopData.NodeMetadata[i] = make([]interface{}, ws.NumOfCols-x)
        for j := x; j < ws.NumOfCols; j++ {
            if err := ws.verifyCell(elements[i], j, y); err != nil {
                return err
            }
        }
    }
    return nil
}

func (ws *Worksheet) verifyCell(idx, j, y int) error {
    verifier, err := newVerifier(idx, ws.HazopData.Raw[j][y-1])
    if err != nil {
        return err
    }
    val, err := verifier.run()
    if err != nil {
        err := fmt.Errorf("Value error %d %d: %v", j, y, err)
        ws.HazopData.newError(err)
    } else {
        ws.HazopData.NodeMetadata[j-1][y-1] = val
    }
    return nil
}

type checker func(interface{}, int, int) error
type parser func(string) (interface{}, error)

type verifier struct {
    parse parser
    check checker
    val   string
    min   int
    max   int
}

func newVerifier(idx int, val string) (*verifier, error) {
    switch ctype := settings.Hazop.Element[idx].CellType; ctype {
    case settings.Hazop.CellType.String:
        return &verifier{
            parse: parseStr,
            check: checkStrLen,
            val:   val,
            min:   settings.Hazop.Element[idx].MinLen,
            max:   settings.Hazop.Element[idx].MaxLen,
        }, nil
    case settings.Hazop.CellType.Integer:
        return &verifier{
            parse: parseInt,
            check: checkIntRange,
            val:   val,
            min:   settings.Hazop.Element[idx].MinLen,
            max:   settings.Hazop.Element[idx].MaxLen,
        }, nil
    case settings.Hazop.CellType.Float:
        return &verifier{
            parse: parseFloat,
            check: checkFloatRange,
            val:   val,
            min:   settings.Hazop.Element[idx].MinLen,
            max:   settings.Hazop.Element[idx].MaxLen,
        }, nil
    default:
        return nil, fmt.Errorf("Unknown cell type: %d", ctype)
    }
}

func (v *verifier) run() (interface{}, error) {
    newVal, err := v.parse(v.val)
    if err != nil {
        return nil, err
    }
    if err := v.check(newVal, v.min, v.max); err != nil {
        return nil, err
    }
    return newVal, nil
}

func getMapValues(a map[int]string) []string {
    values := []string{}
    for _, v := range a {
        values = append(values, v)
    }
    return values
}

func getMapKeys(a map[int]string) []int {
    keys := []int{}
    for i := range a {
        keys = append(keys, i)
    }
    return keys
}

func cellsNameToCoordinates(a []string) ([]int, []int, error) {
    vX, vY := make([]int, len(a)), make([]int, len(a))
    for i := 0; i < len(a); i++ {
        x, y, err := cellNameToCoordinates(a[i])
        if err != nil {
            return nil, nil, err
        }
        vX[i], vY[i] = x, y
    }
    return vX, vY, nil
}

func cellNameToCoordinates(a string) (int, int, error) {
    x, y, err := excelize.CellNameToCoordinates(a)
    if err != nil {
        return 0, 0, fmt.Errorf("Error parsing coordinate name: %v", err)
    }
    return x, y, nil
}

func checkEqualVector(a []int) bool {
    for i := 1; i < len(a); i++ {
        if a[0] != a[i] {
            return false
        }
    }
    return true
}

func parseStr(val string) (interface{}, error) {
    return val, nil
}

func parseInt(val string) (interface{}, error) {
    if v, err := strconv.Atoi(val); err != nil {
        return nil, err
    } else {
        return v, nil
    }
}

func parseFloat(val string) (interface{}, error) {
    if v, err := strconv.ParseFloat(val, 32); err == nil {
        return nil, err
    } else {
        return v, nil
    }
}

func checkStrLen(val interface{}, min, max int) error {
    if len(val.(string)) <= min || len(val.(string)) >= max {
        return errors.New("Out of range")
    } else {
        return nil
    }
}

func checkIntRange(val interface{}, min, max int) error {
    if val.(int) <= min || val.(int) >= max {
        return errors.New("Out of range")
    } else {
        return nil
    }
}

func checkFloatRange(val interface{}, min, max int) error {
    if val.(float32) <= float32(min) || val.(float32) >= float32(max) {
        return errors.New("Out of range")
    } else {
        return nil
    }
}

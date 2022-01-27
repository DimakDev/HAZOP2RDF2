package workbook

import (
    "errors"
    "fmt"
)

var ErrNoHazopElementsFound = errors.New("Error no Hazop elements found")

type DataType struct {
    Metadata int `mapstructure:"metadata"`
    Analysis int `mapstructure:"analysis"`
}

type CellType struct {
    String  int `mapstructure:"string"`
    Integer int `mapstructure:"integer"`
    Float   int `mapstructure:"float"`
}

type Element struct {
    DataType int    `mapstructure:"data_type"`
    Name     string `mapstructure:"name"`
    Regex    string `mapstructure:"regex"`
    CellType int    `mapstructure:"cell_type"`
    MinLen   int    `mapstructure:"min_len"`
    MaxLen   int    `mapstructure:"max_len"`
}

type HazopSettings struct {
    DataType DataType  `mapstructure:"data_type"`
    CellType CellType  `mapstructure:"cell_type"`
    Element  []Element `mapstructure:"element"`
}

var Hazop HazopSettings

func (h *HazopSettings) Elements(dtype int) ([]Element, error) {
    elements := []Element{}
    for _, e := range Hazop.Element {
        if e.DataType != dtype {
            continue
        }
        elements = append(elements, e)
    }

    if len(elements) == 0 {
        return nil, fmt.Errorf("%v %d", ErrNoHazopElementsFound, dtype)
    }

    return elements, nil
}

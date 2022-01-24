package workbook

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

func (h *HazopSettings) groupHazopElements(dtype int) map[int]Element {
    elements := map[int]Element{}
    for i, e := range Hazop.Element {
        if e.DataType != dtype {
            continue
        }
        elements[i] = e
    }
    return elements
}

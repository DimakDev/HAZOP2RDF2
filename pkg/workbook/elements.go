package workbook

type datatype int

const (
    STRING datatype = iota
    INTEGER
    FLOAT
)

type HazopElement struct {
    Id    int      `mapstructure:"id"`
    Name  string   `mapstructure:"name"`
    Regex string   `mapstructure:"regex"`
    Type  datatype `mapstructure:"data_type"`
    Min   int      `mapstructure:"min_len"`
    Max   int      `mapstructure:"max_len"`
}

type HazopElements struct {
    Elements []HazopElement `mapstructure:"elements"`
}

var Hazop HazopElements

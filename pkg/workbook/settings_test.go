package workbook

import (
    "os"
    "testing"

    "github.com/spf13/viper"
    "github.com/stretchr/testify/assert"
)

func TestElements(t *testing.T) {
    assert := assert.New(t)

    var err error

    err = os.Chdir("../..")
    assert.Empty(err)

    viper.SetConfigName("settings")
    viper.SetConfigType("toml")
    viper.AddConfigPath(".")

    err = viper.ReadInConfig()
    assert.Empty(err)

    err = viper.UnmarshalKey("hazop", nil)
    assert.Error(err)

    err = viper.UnmarshalKey("hazop", &Hazop)
    assert.Empty(err)

    var elements []Element

    elements, err = Hazop.Elements(Hazop.DataType.Metadata)
    assert.Empty(err)
    assert.Len(elements, 2)

    elements, err = Hazop.Elements(Hazop.DataType.Analysis)
    assert.Empty(err)
    assert.Len(elements, 13)

    elements, err = Hazop.Elements(3)
    assert.Error(err)
    assert.Empty(elements)
}

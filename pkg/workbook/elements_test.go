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

    viper.SetConfigName("config")
    viper.SetConfigType("toml")
    viper.AddConfigPath(".")

    err = viper.ReadInConfig()
    assert.Empty(err)

    err = viper.UnmarshalKey("hazop", nil)
    assert.Error(err)

    err = viper.UnmarshalKey("hazop", &Elements)
    assert.Empty(err)
}

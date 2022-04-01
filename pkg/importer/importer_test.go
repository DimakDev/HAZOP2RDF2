package importer

import (
    "io/ioutil"
    "log"
    "os"
    "path/filepath"
    "strings"
    "testing"

    "github.com/spf13/viper"
    "github.com/stretchr/testify/assert"
)

func init() {
    os.Chdir("../..")

    viper.SetConfigName("manifest")
    viper.SetConfigType("toml")
    viper.AddConfigPath(".")

    if err := viper.ReadInConfig(); err != nil {
        log.Fatal(err)
    }

    if err := viper.UnmarshalKey("hazop", &Hazop); err != nil {
        log.Fatal(err)
    }
}

func BenchmarkReadVerifyWorkbook(b *testing.B) {
    hazopFiles, err := ioutil.ReadDir("hazop")
    if err != nil {
        log.Fatal(err)
    }

    for i := 0; i < b.N; i++ {
        for _, f := range hazopFiles {
            if strings.HasSuffix(f.Name(), ".xlsx") {
                fpath := filepath.Join("hazop", f.Name())
                _, err := ImportWorkbook(fpath)
                if err != nil {
                    log.Fatal(err)
                }
            }
        }
    }
}

func TestReadVerifyWorkbook(t *testing.T) {
    assert := assert.New(t)

    wb, err := ImportWorkbook("")
    assert.Error(err)
    assert.Empty(wb)

    hazopFiles, err := ioutil.ReadDir("hazop")
    assert.Empty(err)
    assert.NotEmpty(hazopFiles)

    for _, f := range hazopFiles {
        if strings.HasSuffix(f.Name(), ".xlsx") {
            fpath := filepath.Join("hazop", f.Name())
            wb, err := ImportWorkbook(fpath)
            assert.Empty(err)
            assert.NotEmpty(wb)
        }
    }
}

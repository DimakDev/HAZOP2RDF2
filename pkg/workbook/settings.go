package workbook

import (
    "log"

    "github.com/spf13/viper"
)

type Settings struct {
    Package struct {
        Name        string `mapstructure:"name"`
        Description string `mapstructure:"description"`
        Help        string `mapstructure:"help"`
        Version     string `mapstructure:"version"`
        Author      string `mapstructure:"author"`
    } `mapstructure:"package"`
    Common struct {
        DataDir   string `mapstructure:"data_dir"`
        DataExt   string `mapstructure:"data_ext"`
        ReportDir string `mapstructure:"report_dir"`
    } `mapstructure:"common"`
    Hazop struct {
        DataType struct {
            NodeMetadata int `mapstructure:"node_metadata"`
            NodeHazop    int `mapstructure:"node_hazop"`
        } `mapstructure:"data_type"`
        CellType struct {
            String  int `mapstructure:"string"`
            Integer int `mapstructure:"integer"`
            Float   int `mapstructure:"float"`
        } `mapstructure:"cell_type"`
        Element []struct {
            DataType int    `mapstructure:"data_type"`
            Name     string `mapstructure:"name"`
            Regex    string `mapstructure:"regex"`
            CellType int    `mapstructure:"cell_type"`
            MinLen   int    `mapstructure:"min_len"`
            MaxLen   int    `mapstructure:"max_len"`
        } `mapstructure:"element"`
    } `mapstructure:"hazop"`
}

var settings Settings

func init() {
    viper.SetConfigName("cfg")
    viper.SetConfigType("toml")
    viper.AddConfigPath(".")

    if err := viper.ReadInConfig(); err != nil {
        log.Fatal("Error reading `.toml` settings: ", err)
    }

    if err := viper.Unmarshal(&settings); err != nil {
        log.Fatal("Error parsing `.toml` settings: ", err)
    }
}

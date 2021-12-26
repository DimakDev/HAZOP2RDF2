package workbook

import (
    "log"

    "github.com/spf13/viper"
)

type Settings struct {
    Package struct {
        Name        string `toml:"name"`
        Description string `toml:"description"`
        Help        string `toml:"help"`
        Version     string `toml:"version"`
        Author      string `toml:"author"`
    } `toml:"package"`
    Common struct {
        DataDir string `toml:"data_dir"`
        DataExt string `toml:"data_ext"`
        TextDir string `toml:"text_dir"`
    } `toml:"common"`
    Hazop struct {
        Pattern  string `toml:"pattern"`
        Elements []struct {
            Name  string `toml:"name"`
            Regex string `toml:"regex"`
            Type  string `toml:"type"`
            Min   int    `toml:"min"`
            Max   int    `toml:"max"`
        } `toml:"elements"`
    } `toml:"hazop"`
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

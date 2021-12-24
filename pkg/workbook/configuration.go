package workbook

import (
	"log"

	"github.com/spf13/viper"
)

type Configuration struct {
	Package struct {
		Name        string `mapstructure"name"`
		Description string `mapstructure"description"`
		Help        string `mapstructure"help"`
		Version     string `mapstructure"version"`
		Author      string `mapstructure"author"`
	} `mapstructure"package"`
	Common struct {
		DataDir string `mapstructure"data_dir"`
		DataExt string `mapstructure"data_ext"`
		TextDir string `mapstructure"text_dir"`
	} `mapstructure"common"`
	Hazop struct {
		Worksheet struct {
			Metadata struct {
				Regex string `mapstructure"regex"`
				Name  string `mapstructure"name"`
			} `mapstructure"metadata"`
			Analysis struct {
				Regex string `mapstructure"regex"`
				Name  string `mapstructure"name"`
			} `mapstructure"analysis"`
		} `mapstructure"worksheet"`
		Metadata struct {
			Element struct {
				Label struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"label"`
				Description struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"description"`
			} `mapstructure"element"`
		} `mapstructure"metadata"`
		Analysis struct {
			Element struct {
				Reference struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"reference"`
				GuideWord struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"guide_word"`
				Parameter struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"parameter"`
				Deviation struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"deviation"`
				Cause struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"cause"`
				Consequense struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"consequense"`
				Safeguard struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"safeguard"`
				ActionRef struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"action_ref"`
				Action struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"action"`
				ActionOn struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"action_on"`
				Severity struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"severity"`
				Likehood struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"likehood"`
				RiskPriority struct {
					Regex  string `mapstructure"regex"`
					Name   string `mapstructure"name"`
					Type   string `mapstructure"type"`
					MinLen int    `mapstructure"min_len"`
					MaxLen int    `mapstructure"max_len"`
					Range  int    `mapstructure"range"`
				} `mapstructure"risk_priority"`
			} `mapstructure"element"`
		} `mapstructure"analysis"`
	} `mapstructure"hazop"`
}

var Config Configuration

func init() {
	viper.SetConfigName("cfg")
	viper.SetConfigType("toml")
	viper.AddConfigPath(".")

	if err := viper.ReadInConfig(); err != nil {
		log.Fatal("Error reading `.toml` config: ", err)
	}

	err := viper.Unmarshal(&Config)
	if err != nil {
		log.Fatal("Error parsing `.toml` config: ", err)
	}
}

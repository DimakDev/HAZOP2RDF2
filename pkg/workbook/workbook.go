package workbook

import (
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Workbook struct {
	File         *excelize.File
	Name         string
	Path         string
	Worksheets   map[int]*Worksheet
	Verification map[int][]Verification
}

type Verification struct {
	Action    string
	IsValid   bool
	Message   string
	Worksheet string
}

type Worksheet struct {
	NodeName   string
	IsMetadata bool
	IsAnalysis bool
	Rows       [][]string
}

func New(datapath string) *Workbook {
	_, fname := filepath.Split(datapath)
	return &Workbook{
		Name:         strings.TrimSuffix(fname, filepath.Ext(fname)),
		Path:         datapath,
		Worksheets:   make(map[int]*Worksheet),
		Verification: make(map[int][]Verification),
	}
}

package workbook

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func (wb *Workbook) ReadWorkbook() error {
	if f, err := excelize.OpenFile(wb.Path); err == nil {
		wb.File = f
		return nil
	} else {
		return fmt.Errorf("Error opening `%s`: %v", wb.Path, err)
	}
}

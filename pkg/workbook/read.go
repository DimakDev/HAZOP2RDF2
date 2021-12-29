package workbook

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func (wb *Workbook) readWorkbook() error {
	if err := wb.getNumOfCols(); err != nil {
		return err
	}
	if err := wb.getNumOfRows(); err != nil {
		return err
	}
	if err := wb.readHazopData(); err != nil {
		return err
	}
	if err := wb.readHazopHeader(); err != nil {
		return err
	}
	return nil
}

func (wb *Workbook) getNumOfCols() error {
	for _, ws := range wb.Worksheets {
		cols, err := wb.File.GetCols(ws.SheetName)
		if err != nil {
			return fmt.Errorf("Error reading columns: %v", err)
		}
		ws.NumOfCols = len(cols)
	}
	return nil
}

func (wb *Workbook) getNumOfRows() error {
	for _, ws := range wb.Worksheets {
		rows, err := wb.File.GetRows(ws.SheetName)
		if err != nil {
			return fmt.Errorf("Error reading rows: %v", err)
		}
		ws.NumOfRows = len(rows)
	}
	return nil
}

func (wb *Workbook) readHazopData() error {
	for _, ws := range wb.Worksheets {
		ws.HazopData.Raw = make([][]string, ws.NumOfCols)
		for i := 0; i < ws.NumOfCols; i++ {
			ws.HazopData.Raw[i] = make([]string, ws.NumOfRows)
			for j := 0; j < ws.NumOfRows; j++ {
				cname, err := excelize.CoordinatesToCellName(i+1, j+1)
				if err != nil {
					return fmt.Errorf("Error converting cell name: %v", err)
				}
				cell, err := wb.File.GetCellValue(ws.SheetName, cname)
				if err != nil {
					return fmt.Errorf("Error reading cell value: %v", err)
				}
				ws.HazopData.Raw[i][j] = cell
			}
		}
	}
	return nil
}

func (wb *Workbook) readHazopHeader() error {
	for _, ws := range wb.Worksheets {
		for j, el := range settings.Hazop.Element {
			hhc, err := wb.File.SearchSheet(ws.SheetName, el.Regex, true)
			if err != nil {
				return fmt.Errorf("Error scanning elements: %v", err)
			}
			switch len(hhc) {
			case 0:
				msg := fmt.Sprintf("Element not found: `%s`", el.Name)
				ws.HazopHeader.newWarning(msg)
			case 1:
				ws.HazopHeader.Raw[j] = hhc[0]
				msg := fmt.Sprintf("Element found: `%s`", el.Name)
				ws.HazopHeader.newInfo(msg)
			default:
				msg := fmt.Sprintf("Element multiple coordinates: `%s` %v",
					el.Name, hhc)
				ws.HazopHeader.newWarning(msg)
			}
		}
	}
	return nil
}

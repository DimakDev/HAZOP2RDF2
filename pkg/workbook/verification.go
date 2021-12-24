package workbook

import (
    "fmt"
    "regexp"
    "strings"
)

func (wb *Workbook) appendVerification(i int, action, message, worksheet string, isValid bool) {
    wb.Verification[i] = append(wb.Verification[i], Verification{Action: action, IsValid: isValid, Message: message, Worksheet: worksheet})
}

func (wb *Workbook) appendWorksheet(i int, nodeName string, isMetadata, isAnalysis bool) {
    wb.Worksheets[i] = &Worksheet{NodeName: nodeName, IsMetadata: isMetadata, IsAnalysis: isAnalysis}
}

func (wb *Workbook) VerifyWorksheets() error {
    for i, sname := range wb.File.GetSheetMap() {
        // Name convention: << Node Name >> - << Worksheet Type >>
        regex := strings.Split(regexp.MustCompile(Config.Hazop.Worksheet.Metadata.Regex).FindString(sname), "-")

        if len(regex[0]) == 0 {
            wb.appendVerification(i, "Worksheet Verification", fmt.Sprintf("Worksheet `%s` does not contain a valid Node name", sname), sname, false)
            continue
        }

        if strings.ToLower(regex[1]) == strings.ToLower(Config.Hazop.Worksheet.Metadata.Name) {
            wb.appendWorksheet(i, regex[0], true, false)
            wb.appendVerification(i, "Worksheet Verification", fmt.Sprintf("Worksheet `%s` contain a valid name", sname), sname, true)
            continue
        }

        if strings.ToLower(regex[1]) == strings.ToLower(Config.Hazop.Worksheet.Analysis.Name) {
            wb.appendWorksheet(i, regex[0], false, true)
            wb.appendVerification(i, "Worksheet Verification", fmt.Sprintf("Worksheet `%s` contain a valid name", sname), sname, true)
            continue
        }

        wb.appendVerification(i, "Worksheet Verification", fmt.Sprintf("Worksheet `%s` does not contain a valid type", sname), sname, false)
    }

    return nil
}

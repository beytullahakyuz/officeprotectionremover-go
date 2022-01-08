package tools

import (
	"errors"

	"github.com/beevik/etree"
)

func ExecExcelFile(tp string, xmlFilePath string) (bool, error) {
	xmlFile := etree.NewDocument()
	if err := xmlFile.ReadFromFile(xmlFilePath); err != nil {
		return false, err
	}
	if tp == "workbook" {
		for _, lElement := range xmlFile.ChildElements() {
			if lElement.Tag == "workbook" {
				rootElement := xmlFile.SelectElement("workbook")
				for _, cElement := range rootElement.ChildElements() {
					if cElement.Tag == "workbookProtection" {
						workbookProcElement := rootElement.SelectElement("workbookProtection")
						_ = rootElement.RemoveChild(workbookProcElement)
					}
					if cElement.Tag == "fileSharing" {
						workbookProcElement := rootElement.SelectElement("fileSharing")
						_ = rootElement.RemoveChild(workbookProcElement)
					}
				}
			}
		}
		err := xmlFile.WriteToFile(xmlFilePath)
		if err != nil {
			return false, err
		}
	} else if tp == "worksheet" {
		for _, lElement := range xmlFile.ChildElements() {
			if lElement.Tag == "worksheet" {
				rootElement := xmlFile.SelectElement("worksheet")
				for _, cElement := range rootElement.ChildElements() {
					if cElement.Tag == "sheetProtection" {
						workbookProcElement := rootElement.SelectElement("sheetProtection")
						_ = rootElement.RemoveChild(workbookProcElement)
					}
				}
			}
		}
		err := xmlFile.WriteToFile(xmlFilePath)
		if err != nil {
			return false, err
		}
	} else {
		return false, errors.New("invalid parameter!")
	}
	return true, nil
}

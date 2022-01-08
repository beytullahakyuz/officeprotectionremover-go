package tools

import (
	"github.com/beevik/etree"
)

func ExecWordFile(xmlFilePath string) (bool, error) {
	xmlFile := etree.NewDocument()
	if err := xmlFile.ReadFromFile(xmlFilePath); err != nil {
		return false, err
	}
	for _, lElement := range xmlFile.ChildElements() {
		if lElement.Tag == "settings" {
			rootElement := xmlFile.SelectElement("settings")
			for _, cElement := range rootElement.ChildElements() {
				if cElement.Tag == "documentProtection" {
					workbookProcElement := rootElement.SelectElement("documentProtection")
					_ = rootElement.RemoveChild(workbookProcElement)
				}
				if cElement.Tag == "writeProtection" {
					workbookProcElement := rootElement.SelectElement("writeProtection")
					_ = rootElement.RemoveChild(workbookProcElement)
				}
			}
		}
	}
	err := xmlFile.WriteToFile(xmlFilePath)
	if err != nil {
		return false, err
	}
	return true, nil
}

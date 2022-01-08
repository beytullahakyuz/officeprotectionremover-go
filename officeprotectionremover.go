/*

	author: beytullahakyuz
	last update: 08.01.2022

*/

package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	mtool "./tools"
)

type eOfficeTool struct {
	FilePath       string
	FolderPath     string
	FileName       string
	FileNameWExt   string
	FileExtension  string
	TempDir        string
	FileList       []string
	WorkbookDir    string
	WorkbookFile   string
	WorksheetDir   string
	WorksheetFiles []string
}

type wOfficeTool struct {
	FilePath      string
	FolderPath    string
	FileName      string
	FileNameWExt  string
	FileExtension string
	TempDir       string
	FileList      []string
	SettingsDir   string
	SettingsFile  string
}

var EOfficeFile eOfficeTool
var WOfficeFile wOfficeTool

func main() {

	if len(os.Args) < 2 {
		log.Fatal("Invalid parameters!\r\nThis program is support 1 or more than arguments.\r\n- Microsoft Excel or Word file path.\r\n\r\n- Supported file extensions: .xlsx, .docx")
	} else {
		fmt.Println("------------\t\t Results \t\t------------")
		for findex, fitem := range os.Args {
			if findex == 0 {
				continue
			}
			folderpath, filename := filepath.Split(strings.ToLower(fitem))
			fileextension := path.Ext(filename)

			if (fileextension != ".docx") && (fileextension != ".xlsx") {
				fmt.Println(filename + "\t : Failed! - Invalid file type! Please enter the correct parameters with correct file type.")
			} else {
				if fileextension == ".docx" {
					WOfficeFile.FilePath = fitem
					WOfficeFile.FolderPath = folderpath
					WOfficeFile.FileName = filename
					WOfficeFile.FileExtension = fileextension
					WOfficeFile.TempDir = os.TempDir() + "\\temp_" + strconv.Itoa(randFileName())
					WOfficeFile.FileNameWExt = strings.TrimSuffix(WOfficeFile.FileName, WOfficeFile.FileExtension)
					if state, err := mtool.UnZipFile(WOfficeFile.FilePath, WOfficeFile.TempDir); !state {
						fmt.Println(filename + "\t : Failed! - Error: " + err.Error())
					} else {
						allFiles := getFiles(WOfficeFile.TempDir)
						WOfficeFile.FileList = append(WOfficeFile.FileList, allFiles...)
						WOfficeFile.SettingsDir = WOfficeFile.TempDir + "\\word\\"
						WOfficeFile.SettingsFile = WOfficeFile.SettingsDir + "settings.xml"
						if WOfficeFile.SettingsFile != "" {
							state, err := mtool.ExecWordFile(WOfficeFile.SettingsFile)
							if !state {
								fmt.Println(filename + "\t : Failed! - Error: " + err.Error())
							}
						}
						state, err := mtool.MakeZipFile(WOfficeFile.FileList, WOfficeFile.TempDir, WOfficeFile.FolderPath+WOfficeFile.FileNameWExt+"_protectionremoved"+WOfficeFile.FileExtension)
						if !state {
							fmt.Println(filename + "\t : Failed! - Error: " + err.Error())
						}
						defer removeTempFiles()
						fmt.Println(filename + "\t : Successfully completed.")
					}
				} else {
					EOfficeFile.FilePath = fitem
					EOfficeFile.FolderPath = folderpath
					EOfficeFile.FileName = filename
					EOfficeFile.FileExtension = fileextension
					EOfficeFile.TempDir = os.TempDir() + "\\temp_" + strconv.Itoa(randFileName())
					EOfficeFile.FileNameWExt = strings.TrimSuffix(EOfficeFile.FileName, EOfficeFile.FileExtension)
					if state, err := mtool.UnZipFile(EOfficeFile.FilePath, EOfficeFile.TempDir); !state {
						fmt.Println(filename + "\t : Failed! - Error: " + err.Error())
					} else {

						allFiles := getFiles(EOfficeFile.TempDir)
						EOfficeFile.FileList = append(EOfficeFile.FileList, allFiles...)
						EOfficeFile.WorkbookDir = EOfficeFile.TempDir + "\\xl\\"
						EOfficeFile.WorkbookFile = EOfficeFile.WorkbookDir + "workbook.xml"
						EOfficeFile.WorksheetDir = EOfficeFile.TempDir + "\\xl\\worksheets\\"
						worksheetfiles, _ := ioutil.ReadDir(EOfficeFile.WorksheetDir)
						for _, item := range worksheetfiles {
							if !item.IsDir() {
								EOfficeFile.WorksheetFiles = append(EOfficeFile.WorksheetFiles, EOfficeFile.WorksheetDir+item.Name())
							}
						}
						if EOfficeFile.WorkbookFile != "" {
							state, err := mtool.ExecExcelFile("workbook", EOfficeFile.WorkbookFile)
							if !state {
								fmt.Println(filename + "\t : Failed! - Error: " + err.Error())
							}
						}
						for _, wsFile := range EOfficeFile.WorksheetFiles {
							state, err := mtool.ExecExcelFile("worksheet", wsFile)
							if !state {
								fmt.Println(filename + "\t : Failed! - Error: " + err.Error())
							}
						}
						state, err := mtool.MakeZipFile(EOfficeFile.FileList, EOfficeFile.TempDir, EOfficeFile.FolderPath+EOfficeFile.FileNameWExt+"_protectionremoved"+EOfficeFile.FileExtension)
						if !state {
							fmt.Println(filename + "\t : Failed! - Error: " + err.Error())
						}
						defer removeTempFiles()
						fmt.Println(filename + "\t : Successfully completed.")
					}
				}
			}

		}
		fmt.Println("------------\t Developed by beytullahakyuz \t------------")
	}
}

func getFiles(dir string) []string {
	var fileList []string
	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() {
			fileList = append(fileList, path)
		}
		return nil
	})
	if err != nil {
		log.Println(err)
	}
	return fileList
}

func randFileName() int {
	t := time.Now()
	return t.Nanosecond()
}

func removeTempFiles() {
	_ = os.RemoveAll(EOfficeFile.TempDir)
}

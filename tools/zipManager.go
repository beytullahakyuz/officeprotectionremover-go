package tools

import (
	"archive/zip"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
)

func MakeZipFile(files []string, tempFolderPath string, outputfile string) (bool, error) {
	archive, _ := os.Create(outputfile)
	defer archive.Close()

	zipWriter := zip.NewWriter(archive)

	for _, pFile := range files {
		fTemp, err := os.Open(pFile)
		if err != nil {
			return false, err
		}
		defer fTemp.Close()

		//fmt.Println(strings.Replace(pFile, tempFolderPath+"\\", "", -1))
		wTemp, _ := zipWriter.Create(strings.Replace(pFile, tempFolderPath+"\\", "", -1))

		if _, err := io.Copy(wTemp, fTemp); err != nil {
			return false, err
		}

	}
	zipWriter.Close()
	return true, nil
}

func UnZipFile(inputFile, outputDir string) (bool, error) {
	dst := outputDir
	archive, err := zip.OpenReader(inputFile)
	if err != nil {
		panic(err)
	}
	defer archive.Close()

	for _, f := range archive.File {
		filePath := filepath.Join(dst, f.Name)

		if !strings.HasPrefix(filePath, filepath.Clean(dst)+string(os.PathSeparator)) {
			return false, errors.New("invalid file path!")
		}
		if f.FileInfo().IsDir() {
			os.MkdirAll(filePath, os.ModePerm)
			continue
		}

		if err := os.MkdirAll(filepath.Dir(filePath), os.ModePerm); err != nil {
			panic(err)
		}

		dstFile, err := os.OpenFile(filePath, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
		if err != nil {
			panic(err)
		}

		fileInArchive, err := f.Open()
		if err != nil {
			panic(err)
		}

		if _, err := io.Copy(dstFile, fileInArchive); err != nil {
			panic(err)
		}

		dstFile.Close()
		fileInArchive.Close()
	}
	return true, nil
}

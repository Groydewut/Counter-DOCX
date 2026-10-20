package main

import (
	"archive/zip"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"io/fs"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"
)

const (
	charsPerLine = 50
	linesPerPage = 35
)

var (
	ErrDocumentXmlNotFound = errors.New("компонент word/document.xml не найден внутри архива")
	ErrFileNotExist        = errors.New("указаный файл не существует, проверьте имя")
	ErrNotARegularFile     = errors.New("указанный путь не является обычным файлом")
	ErrIsDir               = errors.New("указанный путь ведет к директории (папке)")
	ErrInvalidExt          = errors.New("файл имеет неверное расширение")
	ErrCorruptedDocx       = errors.New("файл поврежден или не является валидным документом Word OpenXML")
)

func CheckFile(path string, ext string) (bool, error) {
	info, err := os.Stat(path)
	if err == nil {
	} else if errors.Is(err, fs.ErrNotExist) {
		return false, ErrFileNotExist
	} else {
		return false, fmt.Errorf("Произошла не известная ошибка: %w\n", err)
	}

	if info.IsDir() {
		return false, ErrIsDir
	}

	if !info.Mode().IsRegular() {
		return false, ErrNotARegularFile
	}

	switch ext {
case ".txt":
		file, err := os.Open(path)
		if err != nil {
			return false, err
		}
		defer file.Close()
		//Читаем первые 512 байт (этого достаточно для определения MIME-типа)
		buffer := make([]byte, 512)
		n, err := file.Read(buffer)
		if err != nil && err != io.EOF {
			return false, ErrInvalidExt
		}
		contentType := http.DetectContentType(buffer[:n])
		if strings.HasPrefix(contentType, "text/") {
			return true, nil
		}
	case ".docx":
		zipReader, err := zip.OpenReader(path)
		if err != nil {
			return false, ErrCorruptedDocx
		}
		defer zipReader.Close()

		for _, file := range zipReader.File {
			if file.Name == "word/document.xml" {
				return true, nil
			}
		}
		return false, ErrCorruptedDocx
	}
	return false, ErrInvalidExt
}

func ReadDocxText(path string) (string, error) {

	zipOpen, err := zip.OpenReader(path)
	if err != nil {
		return "", ErrCorruptedDocx
	}
	defer zipOpen.Close()

	var docFile *zip.File
	for _, file := range zipOpen.File {
		if file.Name == "word/document.xml" {
			docFile = file
			break
		}
	}

	if docFile == nil {
		return "", ErrDocumentXmlNotFound
	}

	rc, err := docFile.Open()
	if err != nil {
		return "", err
	}
	defer rc.Close()

	return extractTextFromXML(rc)
}

func extractTextFromXML(r io.Reader) (string, error) {
	var builder strings.Builder
	decoder := xml.NewDecoder(r)

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", err
		}
		switch se := token.(type) {
		case xml.StartElement:
			if se.Name.Local == "t" {
				var txt string
				if err := decoder.DecodeElement(&txt, &se); err != nil {
					return "", err
				}
				builder.WriteString(txt)
			}
		case xml.EndElement:
			if se.Name.Local == "p" {
				builder.WriteString("\n")
			}
		}
	}
	return strings.TrimSpace(builder.String()), nil
}

func EstimatePages(text string) int {
	if len(text) == 0 {
		return 0
	}

	paragraphs := strings.Split(text, "\n")
	totalLines := 0

	for _, p := range paragraphs {
		charCount := utf8.RuneCountInString(p)

		pLines := 1
		if charCount > 0 {
			pLines = (charCount + charsPerLine - 1) / charsPerLine
		}

		totalLines += pLines
	}
	totalPages := (totalLines + linesPerPage - 1) / linesPerPage

	return totalPages
}

func CountPages(path string) (int, error) {

	var text string
	ext := strings.ToLower(filepath.Ext(path))
	if ext == ".docx" {
		status, err := CheckFile(path, ext)
		if err != nil {
			return 0, err
		}

		if status {
			text, err = ReadDocxText(path)
			if err != nil {
				return 0, err
			}
		} else {
			return 0, ErrFileNotExist
		}
	} else {
		return 0, ErrInvalidExt
	}
	return EstimatePages(text), nil
}

func main() {
	
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/" {
			http.NotFound(w, r)
			return
		}

		http.ServeFile(w, r, "index.html")
	})


	http.HandleFunc("/upload", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "Разрешён только POST", http.StatusMethodNotAllowed)
			return
		}

		r.Body = http.MaxBytesReader(w, r.Body, 10<<20)
		
		err := r.ParseMultipartForm(10 << 20)
		if err != nil {
			var maxBytesErr *http.MaxBytesError

			if errors.As(err, &maxBytesErr) {
				http.Error(
					w,
					"Файл слишком большой. Максимальный размер запроса — 10 МиБ",
					http.StatusRequestEntityTooLarge,
				)
				return
			}

			http.Error(w, "Не удалось обработать форму", http.StatusBadRequest)
			return
		}

		file, header, err := r.FormFile("document")
		if err != nil {
			http.Error(w, "Не удалось получить файл", http.StatusBadRequest)
			return
		}
		defer file.Close()

		tempFile, err := os.CreateTemp("", "document-*.docx")
		if err != nil {
			http.Error(w, "Не удалось создать временный файл", http.StatusInternalServerError)
			return
		}
		defer os.Remove(tempFile.Name())
		defer tempFile.Close()

		_, err = io.Copy(tempFile, file)
		if err != nil {
			http.Error(w, "Не удалось сохранить документ", http.StatusInternalServerError)
			return
		}

		valid, err := CheckFile(tempFile.Name(), ".docx")
		if err != nil || !valid {
			http.Error(w, "Файл не является корректным DOCX-документом", http.StatusBadRequest)
			return
		}

		pages, err := CountPages(tempFile.Name())
		if err != nil {
			http.Error(w, err.Error(), http.StatusBadRequest)
			return
		}

		
		
		w.Header().Set("Content-Type", "application/json; charset=utf-8")

		response := struct {
			Filename string `json:"filename"`
			Pages    int    `json:"pages"`
		}{
			Filename: filepath.Base(header.Filename),
			Pages:    pages,
		}

		err = json.NewEncoder(w).Encode(response)
		if err != nil {
			http.Error(w, "Не удалось сформировать ответ", http.StatusInternalServerError)
			return
		}
	})

	port := os.Getenv("PORT")

		if port == "" {
			port = "8080"
		}

	fmt.Printf("Сервер запущен на http://localhost:%s",port)

	

	log.Fatal(http.ListenAndServe(":"+port, nil))
}

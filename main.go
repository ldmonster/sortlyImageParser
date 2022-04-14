package main

import (
	"crypto/md5"
	"encoding/hex"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"strconv"
	"strings"
	"time"

	excelize "github.com/xuri/excelize/v2"
	cli "gopkg.in/urfave/cli.v1"
)

const ()

// Port of this service
var wPort string = "8080"
var rootFolder string = ""
var rootLinks string = ""

func uploadFile(w http.ResponseWriter, r *http.Request) {
	fmt.Println("File Upload Endpoint Hit")

	// Parse our multipart form, 10 << 20 specifies a maximum
	// upload of 10 MB files.
	r.ParseMultipartForm(10 << 20)
	// FormFile returns the first file for the given key `myFile`
	// it also returns the FileHeader so we can get the Filename,
	// the Header and the size of the file
	file, handler, err := r.FormFile("myFile")
	if err != nil {
		fmt.Println("Error Retrieving the File")
		fmt.Printf("Error: %s", err)
		return
	}
	defer file.Close()
	fmt.Printf("Uploaded File: %+v\n", handler.Filename)
	fmt.Printf("File Size: %+v\n", handler.Size)
	fmt.Printf("MIME Header: %+v\n", handler.Header)

	// read all of the contents of our uploaded file into a
	// byte array
	fileBytes, err := io.ReadAll(file)
	if err != nil {
		fmt.Printf("Error: %s", err)
	}
	// write this byte array to our temporary file

	if _, err := os.Stat(rootFolder + "temp-files"); err != nil {
		if err := os.MkdirAll("temp-files", 0777); err != nil {
			fmt.Printf("Error: %s", err)
		} else {
			fmt.Println("Directory created /temp-files")
		}
	}

	if err := os.WriteFile("temp-files/"+handler.Filename, fileBytes, 0777); err != nil {
		fmt.Printf("Error: %s", err)
	}

	// return that we have successfully uploaded our file!
	fmt.Fprintf(w, "Successfully Uploaded File\nWait for a pictures loading\n")
	go readExcel("temp-files/" + handler.Filename)
}

func readExcel(filename string) {
	fileNameExcel := ""
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	newExcel := f

	row := 2
	for {
		nameCell, _ := excelize.CoordinatesToCellName(1, row)
		entryName, err := f.GetCellValue("Sheet1", nameCell)
		if err != nil {
			log.Fatal(err)
		}
		if entryName == "" {
			break
		}
		nameCell, _ = excelize.CoordinatesToCellName(2, row)
		entryType, _ := f.GetCellValue("Sheet1", nameCell)
		folder := getExcelFolder(f, row)
		if entryType == "Folder" {
			if folder == "/" {
				folder = ""
			}
			if _, err := os.Stat(rootFolder + folder + entryName); err != nil {
				if err := os.MkdirAll(rootFolder+folder+entryName, 0777); err != nil {
					fmt.Printf("Error: %s", err)
				} else {
					fmt.Println("Directory created " + rootFolder + folder + entryName)
				}
			}
		}
		for column := 10; column < 13; column++ {
			nameCell, _ := excelize.CoordinatesToCellName(column, row)
			photoURL, _ := f.GetCellValue("Sheet1", nameCell)
			if photoURL != "" {
				if folder == "" {
					folder = entryName + "/"
					fileNameExcel = entryName
				}
				data := md5.Sum([]byte(entryName + time.Now().String()))
				hash := hex.EncodeToString(data[:])
				if _, err := os.Stat(rootFolder + folder + entryName + " photo(" + strconv.Itoa(column-9) + ")" + hash + ".jpg"); err != nil {
					saveFileFromURL(photoURL, entryName+" photo("+strconv.Itoa(column-9)+")"+hash+".jpg", rootFolder+folder)
				}
				newExcel.SetCellValue("Sheet1", nameCell, rootLinks+folder+entryName+" photo("+strconv.Itoa(column-9)+")"+hash+".jpg")
			}
		}
		row++
	}

	if _, err := os.Stat(rootFolder + "excel"); err != nil {
		if err := os.MkdirAll(rootFolder+"excel", 0777); err != nil {
			fmt.Printf("Error: %s", err)
		} else {
			fmt.Println("Directory created " + rootFolder + "excel")
		}
	}
	if err := newExcel.SaveAs(rootFolder + "excel/" + fileNameExcel + ".xlsx"); err != nil {
		log.Fatal(err)
	} else {
		fmt.Println("Excel file formed " + rootFolder + "excel/" + fileNameExcel + ".xlsx")
	}
}

func getExcelFolder(f *excelize.File, row int) string {
	folder := ""
	for i := 5; i < 10; i++ {
		nameCell, _ := excelize.CoordinatesToCellName(i, row)
		folderPart, _ := f.GetCellValue("Sheet1", nameCell)
		folder += folderPart + "/"
		folder = strings.Replace(folder, "//", "/", -1)
	}
	return folder
}

func saveFileFromURL(url string, filename string, dir string) {
	resp, err := http.Get(url)
	if err != nil {
		fmt.Printf("Error: %s", err)
		return
	}
	defer resp.Body.Close()
	if body, err := io.ReadAll(resp.Body); err != nil {
		fmt.Printf("Error: %s", err)
	} else {
		if err := os.WriteFile(dir+"/"+filename, body, 0777); err != nil {
			fmt.Printf("Error: %s", err)
		} else {
			fmt.Println("OK File " + rootFolder + dir + filename)
		}
	}
}

type Page struct {
	Title string
	Body  []byte
}

func loadPage(title string) (*Page, error) {
	filename := title + ".html"
	body, err := os.ReadFile(filename)
	if err != nil {
		return nil, err
	}
	return &Page{Title: "Upload your file here", Body: body}, nil
}

func viewHandler(w http.ResponseWriter, r *http.Request) {
	p, _ := loadPage("uploadFile")
	fmt.Fprintf(w, "<h1>%s</h1><div>%s</div>", p.Title, p.Body)
}

func setupRoutes() {
	http.HandleFunc("/upload", uploadFile)
	http.HandleFunc("/", viewHandler)
	log.Fatal(http.ListenAndServe(":"+wPort, nil))
}

func main() {
	app := cli.NewApp()
	app.Name = "sortly_excel_parser"
	app.Version = "1.0.0"
	app.Usage = "Парсит картинки из эксель файлов формата сортли и создает новый эксель файл со ссылками на картинки."
	app.Flags = []cli.Flag{
		cli.StringFlag{
			Name:  "port,p",
			Value: wPort,
			Usage: "Порт веб сервиса, на котором хостится веб страница загрузки",
		},
		cli.StringFlag{
			Name:  "dir,d",
			Value: rootFolder,
			Usage: "Корневая директория сохранения файлов (Вводить желательно в ковычках)",
		},
		cli.StringFlag{
			Name:  "links,l",
			Value: rootLinks,
			Usage: "Начальная директория для формирования ссылок",
		},
		cli.StringFlag{
			Name:  "file,f",
			Value: "",
			Usage: "Директория к файлу, чтобы обработать без хоста",
		},
	}
	app.Action = func(c *cli.Context) {
		wPort = c.GlobalString("port")
		rootFolder = c.GlobalString("dir")
		if rootFolder != "" {
			rootFolder = strings.Replace(rootFolder, "\\", "/", -1)
			if rootFolder[len(rootFolder)-1] != '/' {
				rootFolder += "/"
			}
		}
		rootLinks = c.GlobalString("links")
		if rootLinks != "" && rootLinks[len(rootLinks)-1] != '/' {
			rootLinks += "/"
		}
		filePath := c.GlobalString("file")
		if filePath != "" {
			filePath = strings.Replace(filePath, "\\", "/", -1)
		}
		fmt.Println("Welcome to The Gelikon Opera!")
		fmt.Println("Excel Sortly parser is ready to work\n")
		fmt.Println("Link prefix is " + rootLinks)
		fmt.Println("Rootfolder is " + rootFolder)
		if filePath != "" {
			fmt.Println("File mod is ON!")
			readExcel(filePath)
			fmt.Println("Parsing done!\n")
			return
		}
		fmt.Println("Server started on port " + wPort + "\n")
		fmt.Println("To start working with parser - just open in your browser (your IP):" + wPort)
		fmt.Println("For example 127.0.0.1:" + wPort)
		setupRoutes()
	}
	app.Run(os.Args)
}

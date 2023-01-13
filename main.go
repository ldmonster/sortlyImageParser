package main

import (
	"crypto/md5"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	_ "image/jpeg"
	"io"
	"log"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
	"sync/atomic"
	"time"

	"github.com/schollz/progressbar/v3"
	excelize "github.com/xuri/excelize/v2"
	"go.uber.org/zap"
	cli "gopkg.in/urfave/cli.v1"
)

const (
	DEFAULT_SERVICE_PORT = "8080"

	htmlPage = `<!DOCTYPE html>
	<html lang="en">
	  <head>
		<meta charset="UTF-8" />
		<meta name="viewport" content="width=device-width, initial-scale=1.0" />
		<meta http-eqgo builuiv="X-UA-Compatible" content="ie=edge" />
		<title>Document</title>
	  </head>
	  <body>

	  </body>
	</html>`

	configFileName = "config.cfg"
)

func NewLogger() (*zap.Logger, error) {
	cfg := zap.NewProductionConfig()
	cfg.OutputPaths = []string{
		"./logs.log",
		"stderr",
	}

	return cfg.Build()
}

type SortlyParserConfig struct {
	Port       string `json:"port"`
	RootFolder string `json:"root_folder"`
	RootLinks  string `json:"root_links"`

	FilePath string `json:"file_path"`
	Logger   *zap.Logger

	SaveConfig bool
}

func NewSortlyParserConfig(logger *zap.Logger) *SortlyParserConfig {
	return &SortlyParserConfig{
		Logger: logger,
	}
}

func (spc *SortlyParserConfig) ParseInput(c *cli.Context) error {
	var errR error

	portRe := regexp.MustCompile(`(?m)^[1-9][0-9]{1,4}$`)
	newPort := portRe.FindString(c.GlobalString("port"))
	if newPort == "" {
		errR = errors.New("bad port input")
	}

	spc.Port = newPort

	rootFolder := c.GlobalString("dir")
	if rootFolder != "" {
		rootFolder = strings.Replace(rootFolder, "\\", "/", -1)
		if rootFolder[len(rootFolder)-1] != '/' {
			rootFolder += "/"
		}

		if _, err := os.Stat(rootFolder); errors.Is(err, os.ErrNotExist) {
			errR = errors.New("root folder does not exist")
		}

		spc.RootFolder = rootFolder
	}

	rootLinks := c.GlobalString("links")
	if rootLinks != "" && rootLinks[len(rootLinks)-1] != '/' {
		rootLinks += "/"

		spc.RootLinks = rootLinks
	}

	filePath := c.GlobalString("file")
	if filePath != "" {
		filePath = strings.Replace(filePath, "\\", "/", -1)

		if _, err := os.Stat(filePath); errors.Is(err, os.ErrNotExist) {
			errR = errors.New("file does not exist")
		}

		spc.FilePath = filePath
	}

	saveConfig := c.GlobalString("cfg")
	if saveConfig == "1" {
		spc.SaveConfig = true
	}

	return errR
}

func (spc *SortlyParserConfig) ReadConfig() error {
	_, err := os.Stat(configFileName)
	if errors.Is(err, os.ErrNotExist) {
		return os.ErrNotExist
	}

	if !errors.Is(err, os.ErrNotExist) && err != nil {
		return fmt.Errorf("can not check file, reason: %v", err)
	}

	f, err := os.Open(configFileName)
	if err != nil {
		return fmt.Errorf("can not open file, reason: %v", err)
	}

	newSpc := &SortlyParserConfig{}

	err = json.NewDecoder(f).Decode(newSpc)
	if err != nil {
		return fmt.Errorf("can not decoding, reason: %v", err)
	}

	if newSpc.Port == "" {
		return errors.New("config port is empty, bad config")
	}

	if newSpc.RootFolder == "" {
		return errors.New("root folder is empty, bad config")
	}

	if newSpc.RootLinks == "" {
		return errors.New("root links is empty, bad config")
	}

	spc.Port = newSpc.Port
	spc.RootFolder = newSpc.RootFolder
	spc.RootLinks = newSpc.RootLinks

	return nil
}

func (spc *SortlyParserConfig) WriteConfig() {
	f, err := os.Create("config.cfg")
	if err != nil {
		spc.Logger.Error(
			fmt.Sprintf("can not open file to write config, reason: %v", err),
		)
	}

	configStr := fmt.Sprintf("{\n\t%s%s%s\n\t%s%s%s\n\t%s%s%s\n}",
		`"port":"`, spc.Port, `",`,
		`"root_folder":"`, spc.RootFolder, `",`,
		`"root_links":"`, spc.RootLinks, `"`)

	_, err = f.WriteString(configStr)
	if err != nil {
		spc.Logger.Error(
			fmt.Sprintf("can not write config to file, reason: %v", err),
		)
	}
}

type SortlyParser struct {
	Cfg *SortlyParserConfig
}

type Parser struct {
	Cfg *SortlyParserConfig

	excelFileOrig *excelize.File
	excelFileNew  *excelize.File

	currRow int

	ImageCount int

	downloadedItems int32

	pbar *progressbar.ProgressBar

	FolderList []*Folder
}

type Folder struct {
	Name string
	Item *Item
	Path string

	FolderList []*Folder
	ItemList   []*Item
}

type Item struct {
	Name string
	Url  []string

	Row int
}

func NewSortlyParser(spc *SortlyParserConfig) *SortlyParser {
	return &SortlyParser{
		Cfg: spc,
	}
}

func (sp *SortlyParser) CreateParser() *Parser {
	return &Parser{
		Cfg: sp.Cfg,
	}
}

func (sp *Parser) ReadExcel() {
	f, err := excelize.OpenFile(sp.Cfg.FilePath)
	if err != nil {
		sp.Cfg.Logger.Error("can not open file")
	}

	sp.excelFileOrig = f
	sp.excelFileNew = f

	sp.currRow = 2
	for {
		nameCell, _ := excelize.CoordinatesToCellName(1, sp.currRow)
		entryName, err := f.GetCellValue("Sheet1", nameCell)
		if err != nil {
			sp.Cfg.Logger.Error("cant read Entry Name cell")
		}

		if entryName == "" {
			break
		}

		nameCell, _ = excelize.CoordinatesToCellName(2, sp.currRow)
		entryType, _ := f.GetCellValue("Sheet1", nameCell)

		_ = entryType

		sp.AddItem(entryType, entryName)

		sp.currRow++
	}

	fmt.Println()
	sp.Cfg.Logger.Info("Excel parse done")
	fmt.Println()

	sp.ParseAllItems(sp.FolderList)

	fmt.Println()
	sp.Cfg.Logger.Info("Images parse done")
	fmt.Println()

	sp.SaveExcelFile()

	fmt.Println()
	sp.Cfg.Logger.Info("Sortly parsing work is done!")
	fmt.Println()
}

func (sp *Parser) AddItem(entryType, entryName string) {
	var (
		rootFolder *Folder
		path       = ""
	)

	for i := 5; i < 10; i++ {
		nameCell, _ := excelize.CoordinatesToCellName(i, sp.currRow)
		folderName, _ := sp.excelFileOrig.GetCellValue("Sheet1", nameCell)

		if folderName == "" {
			if i == 5 && entryType == "Folder" {
				sp.FolderList = append(sp.FolderList, &Folder{
					Name: entryName,
					Item: &Item{
						Name: entryName,
						Url:  sp.GetUrls(),
						Row:  sp.currRow,
					},
					Path: fmt.Sprintf("%s/", entryName),
				})

				break
			}

			break
		}

		for _, f := range sp.FolderList {
			rootFolder = GetFolderByName(f, folderName)
			path += fmt.Sprintf("%s/", rootFolder.Name)
		}

	}

	if rootFolder != nil {
		path += fmt.Sprintf("%s/", entryName)

		switch entryType {
		case "Folder":
			{
				rootFolder.FolderList = append(rootFolder.FolderList, &Folder{
					Name: entryName,
					Item: &Item{
						Name: entryName,
						Url:  sp.GetUrls(),
						Row:  sp.currRow,
					},
					Path: path,
				})
			}
		case "Item":
			{
				rootFolder.ItemList = append(rootFolder.ItemList, &Item{
					Name: entryName,
					Url:  sp.GetUrls(),
					Row:  sp.currRow,
				})
			}
		}

	}
}

func (sp *Parser) GetUrls() []string {
	urls := make([]string, 0, 3)

	for column := 10; column < 13; column++ {
		nameCell, _ := excelize.CoordinatesToCellName(column, sp.currRow)
		photoURL, _ := sp.excelFileOrig.GetCellValue("Sheet1", nameCell)
		if photoURL == "" {
			break
		}
		urls = append(urls, photoURL)
	}

	sp.ImageCount += len(urls)

	return urls
}

func GetFolderByName(rootFolder *Folder, name string) *Folder {
	if rootFolder.Name == name {
		return rootFolder
	}

	for _, f := range rootFolder.FolderList {
		folderR := GetFolderByName(f, name)
		if folderR != nil {
			return folderR
		}
	}

	return nil
}

func (sp *Parser) ParseAllItems(folderList []*Folder) {
	if folderList == nil {
		return
	}

	for _, f := range folderList {
		sp.Save(f.Item, f)

		for _, i := range f.ItemList {
			sp.Save(i, f)
		}

		sp.ParseAllItems(f.FolderList)

	}
}

func (sp *Parser) Save(img *Item, folder *Folder) {
	data := md5.Sum([]byte(img.Name + time.Now().String()))
	hash := hex.EncodeToString(data[:2])

	pictureFolder := fmt.Sprintf("%s%s", sp.Cfg.RootFolder, folder.Path)

	for i := 1; i <= len(img.Url); i++ {
		pictureFilename := fmt.Sprintf("%s photo(%d)%s.jpg", img.Name, i, hash)

		_, err := os.Stat(pictureFolder)
		if errors.Is(err, os.ErrNotExist) {
			mrdirErr := os.MkdirAll(pictureFolder, 0777)
			if mrdirErr != nil {
				sp.Cfg.Logger.Error(
					fmt.Sprintf("cannot create directory for folder, reason:%s", err),
				)

				continue
			}

			sp.Cfg.Logger.Info("Directory created: %s\n" + pictureFolder)
			fmt.Println()
		}

		_, err = os.Stat(pictureFolder + pictureFilename)
		if errors.Is(err, os.ErrNotExist) {
			err := sp.SaveFileFromURL(img.Url[i-1], pictureFilename, pictureFolder)
			if err != nil {
				sp.Cfg.Logger.Error(
					fmt.Sprintf("saving picture from url, reason: %v", err),
				)
			}
		}

		if !errors.Is(err, os.ErrNotExist) && err != nil {
			sp.Cfg.Logger.Error(
				fmt.Sprintf(`picture "%s" local check before saving, reason: %v`, img.Url[i], err),
			)

			break
		}

		nameCell, _ := excelize.CoordinatesToCellName(9+i, img.Row)
		sp.excelFileNew.SetCellValue("Sheet1", nameCell, sp.Cfg.RootLinks+folder.Path+img.Name+" photo("+strconv.Itoa(i)+")"+hash+".jpg")
	}
}

func (sp *Parser) SaveFileFromURL(url string, filename string, dir string) error {
	resp, err := http.Get(url)
	if err != nil {
		return err
	}

	bar := progressbar.NewOptions64(resp.ContentLength,
		progressbar.OptionSetWriter(os.Stderr),
		progressbar.OptionEnableColorCodes(true),
		progressbar.OptionShowBytes(true),
		progressbar.OptionSetWidth(10),
		progressbar.OptionThrottle(65*time.Millisecond),
		progressbar.OptionShowCount(),
		progressbar.OptionOnCompletion(func() {
			fmt.Fprint(os.Stderr, "\n")
		}),
		progressbar.OptionSpinnerType(14),
		progressbar.OptionFullWidth(),
		progressbar.OptionSetRenderBlankState(true),
		progressbar.OptionSetTheme(progressbar.Theme{
			Saucer:        "[green]=[reset]",
			SaucerHead:    "[green]>[reset]",
			SaucerPadding: " ",
			BarStart:      "[",
			BarEnd:        "]",
		}),
	)

	barDescription := fmt.Sprintf(`[cyan][%d/%d][reset] Downloading "%s"...`, sp.downloadedItems+1, sp.ImageCount, filename)
	bar.Describe(barDescription)

	file := fmt.Sprintf("%s/%s", dir, filename)
	f, _ := os.OpenFile(file, os.O_CREATE|os.O_WRONLY, 0777)

	io.Copy(io.MultiWriter(f, bar), resp.Body)
	if err != nil {
		return err
	}

	atomic.AddInt32(&sp.downloadedItems, 1)

	return nil
}

func (sp *Parser) SaveExcelFile() {
	excelFolder := fmt.Sprintf("%sexcel", sp.Cfg.RootFolder)
	_, err := os.Stat(excelFolder)
	if errors.Is(err, os.ErrNotExist) {
		mrdirErr := os.MkdirAll(excelFolder, 0777)
		if mrdirErr != nil {
			sp.Cfg.Logger.Error(
				fmt.Sprintf("Error: cannot create directory for excel folder, reason:%s", err),
			)

			return
		}
	}

	if err != nil && !errors.Is(err, os.ErrNotExist) {
		sp.Cfg.Logger.Error(
			fmt.Sprintf("excel folder local check, reason %v", err),
		)

		return
	}

	sp.Cfg.Logger.Info("Directory created " + sp.Cfg.RootFolder + "excel")
	fmt.Println()

	err = sp.excelFileNew.SaveAs(sp.Cfg.RootFolder + "excel/" + sp.FolderList[0].Name + ".xlsx")
	if err != nil {
		sp.Cfg.Logger.Error(
			fmt.Sprintf("saving resulting excel, reason %v", err),
		)

		return
	}

	sp.Cfg.Logger.Info("Excel file formed " + sp.Cfg.RootFolder + "excel/" + sp.FolderList[0].Name + ".xlsx")
}

func (sp *SortlyParser) UploadFile(w http.ResponseWriter, r *http.Request) {
	sp.Cfg.Logger.Info("File Upload Endpoint Hit")

	// Parse our multipart form, 10 << 20 specifies a maximum
	// upload of 10 MB files.
	r.ParseMultipartForm(10 << 20)
	// FormFile returns the first file for the given key `myFile`
	// it also returns the FileHeader so we can get the Filename,
	// the Header and the size of the file
	file, handler, err := r.FormFile("myFile")
	if err != nil {
		sp.Cfg.Logger.Error(
			fmt.Sprintf("retrieving the file, reason %v", err),
		)

		return
	}
	defer file.Close()

	sp.Cfg.Logger.Info(
		fmt.Sprintf("Uploaded File: %+v\n", handler.Filename),
	)
	sp.Cfg.Logger.Info(
		fmt.Sprintf("File Size: %+v\n", handler.Size),
	)
	sp.Cfg.Logger.Info(
		fmt.Sprintf("MIME Header: %+v\n", handler.Header),
	)

	fileBytes, err := io.ReadAll(file)
	if err != nil {
		sp.Cfg.Logger.Error(
			fmt.Sprintf("read uploading file bytes: %s", err),
		)

		return
	}

	_, err = os.Stat(sp.Cfg.RootFolder + "temp-files")
	if err != nil {
		err := os.MkdirAll("temp-files", 0777)
		if err != nil {
			sp.Cfg.Logger.Error(
				fmt.Sprintf("checking folder for uploading files: %s", err),
			)

			return
		}

		sp.Cfg.Logger.Info("Directory created /temp-files")
	}

	err = os.WriteFile("temp-files/"+handler.Filename, fileBytes, 0777)
	if err != nil {
		sp.Cfg.Logger.Error(
			fmt.Sprintf("writing uploading file: %s", err),
		)
	}

	sp.Cfg.Logger.Info("Successfully upload file")
	fmt.Fprintf(w, "Successfully Uploaded File\nWait for a pictures loading\n")

	sp.Cfg.FilePath = fmt.Sprintf("temp-files/%s", handler.Filename)

	parser := sp.CreateParser()
	go parser.ReadExcel()
}

func (sp *SortlyParser) ViewHandler(w http.ResponseWriter, r *http.Request) {
	title := "Upload your file here"
	body := []byte(
		fmt.Sprintf(`<form
		enctype="multipart/form-data"
		action="http://localhost:%s/upload"
		method="post"
	  >
		<input type="file" name="myFile" />
		<input type="submit" value="upload" />
	  </form>`,
			sp.Cfg.Port),
	)

	fmt.Fprintf(w, "<h1>%s</h1><div>%s</div>", title, body)
}

func (sp *SortlyParser) ServerRun() {
	http.HandleFunc("/upload", sp.UploadFile)
	http.HandleFunc("/", sp.ViewHandler)

	log.Fatal(http.ListenAndServe(
		fmt.Sprintf(":%s", sp.Cfg.Port),
		nil),
	)
}

func (sp *SortlyParser) GetUserInput() error {
	var (
		rootFolder string = ""
		rootLinks  string = ""

		errorWraper = func(anyErr error) func() error {
			err := anyErr

			return func() error {
				return err
			}
		}
		errFunc func() error
	)

	app := cli.NewApp()
	app.Name = "sortly_excel_parser"
	app.Version = "1.1.0"
	app.Usage = "Парсит картинки из эксель файлов формата сортли и создает новый эксель файл со ссылками на картинки."
	app.Flags = []cli.Flag{
		cli.StringFlag{
			Name:  "port,p",
			Value: DEFAULT_SERVICE_PORT,
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
		cli.StringFlag{
			Name:  "cfg",
			Value: "",
			Usage: "Сохранить конфигурацию в файл для работы без параметров командной строки. 1 чтобы сохранить",
		},
	}

	app.Action = func(c *cli.Context) {
		err := sp.Cfg.ParseInput(c)

		errFunc = errorWraper(err)
	}

	app.Run(os.Args)

	fmt.Printf("CONFIG %+v\n", sp.Cfg)
	fmt.Printf("SP %+v\n", sp)

	return errFunc()
}

func main() {
	var (
		sp = &SortlyParser{}
	)

	logger, err := NewLogger()
	if err != nil {
		panic("cannot init logger")
	}

	spc := NewSortlyParserConfig(logger)
	sp = NewSortlyParser(spc)

	logger.Info("Sortly parser is started!")

	err = sp.Cfg.ReadConfig()
	if err != nil && !errors.Is(err, os.ErrNotExist) {
		logger.Error(
			fmt.Sprintf("parse config: %v", err),
		)
	}

	errUserInput := sp.GetUserInput()
	fmt.Printf("USER IUNPUT: %v", errUserInput)

	if errUserInput != nil && err != nil {
		logger.Error(
			fmt.Sprintf("reading user input: %v", errUserInput),
		)
		logger.Error(
			fmt.Sprintf("reading config file: %v", err),
		)
		logger.Fatal("Parser config is invalid. Shutdown...")
	}

	fmt.Println()
	fmt.Println(`/////////////////////////////////////////`)
	fmt.Println(`///   Welcome to The Gelikon Opera!   ///`)
	fmt.Println(`/////////////////////////////////////////`)
	fmt.Println()
	fmt.Println(`Excel Sortly parser is ready to work!`)
	fmt.Println()

	sp.Cfg.Logger.Info(
		fmt.Sprintf(`Link prefix is "%s"`, sp.Cfg.RootLinks),
	)
	sp.Cfg.Logger.Info(
		fmt.Sprintf(`Rootfolder is "%s"`, sp.Cfg.RootFolder),
	)

	if sp.Cfg.SaveConfig {
		sp.Cfg.WriteConfig()
	}

	if sp.Cfg.FilePath != "" {
		sp.Cfg.Logger.Info("File mod is ON!")
		sp.Cfg.Logger.Info(
			fmt.Sprintf(`File path is "%s"`, sp.Cfg.FilePath),
		)

		sp.CreateParser().ReadExcel()

		return
	}

	fmt.Println("\nServer started on port " + sp.Cfg.Port + "\n")
	fmt.Println("To start working with parser - just open in your browser (your IP):" + sp.Cfg.Port)
	fmt.Println("For example 127.0.0.1:" + sp.Cfg.Port)

	sp.ServerRun()
}

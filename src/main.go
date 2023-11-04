package main

import (
	"bufio"
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"reflect"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

type Config struct {
	MaxNumReaders int
	SheetName     string
	OutputName    string
	FolderPath    string
	Delimiter     rune
	Recursive     bool
}

func main() {
	// Flags and setup
	config := parseFlags()
	printFlags(config)

	rowsch := make(chan []string, 1024)

	readwg := new(sync.WaitGroup)
	writewg := new(sync.WaitGroup)

	routineLimiter := make(chan int, config.MaxNumReaders)

	// Launch filewriter routine
	writewg.Add(1)
	go fileWriter(rowsch, config, writewg)

	// Launch fileReader routines
	iterateFolder(config, readwg, routineLimiter, rowsch)

	readwg.Wait()
	close(rowsch)
	writewg.Wait()
}

// parseFlags parses command-line flags and returns a Config with user-specified options.
// It handles options for concurrent readers, sheet name, output file, folder path, CSV delimiter, and recursion.
// Invalid flags or missing folder path terminate the program.
func parseFlags() Config {
	var config Config
	var delimiterString string
	flag.IntVar(&config.MaxNumReaders, "g", 8, "Limits the amount of files being read at once")
	flag.StringVar(&config.SheetName, "sn", "", "Name of the specific sheets to parse. (Optional)")
	flag.StringVar(&config.OutputName, "o", "Output", "Name of output file.")
	flag.StringVar(&config.FolderPath, "p", "", "Path to directory containing excel files to parse.")
	flag.StringVar(&delimiterString, "d", ";", "Csv delimiter for the output file. Maximum 1 letter.")
	flag.BoolVar(&config.Recursive, "r", false, "Recursively goes through each subdirectory and iterates through their excel documents.")

	flag.Parse()

	if len(delimiterString) > 1 || len(delimiterString) == 0 {
		log.Fatal("The csv delimiter provided through [-d] can only be 1 character long. Input provided: [" + delimiterString + "]")
	}

	for _, v := range delimiterString {
		config.Delimiter = v
	}
	if config.FolderPath == "" {
		config.FolderPath = promptuserforpath("Enter path to directory: ")
	}
	if config.FolderPath == "" {
		log.Fatal("A folder path to a directory is mandatory.")
	}

	return config
}

// printFlags prints the values of the provided Config struct's fields with their corresponding field names.
func printFlags(config Config) {
	v := reflect.ValueOf(config)
	fmt.Println("----Variables----")
	for i := 0; i < v.NumField(); i++ {
		// Print field name and value
		fmt.Println(v.Type().Field(i).Name, "\t\t", v.Field(i).Interface())
	}
	fmt.Printf("-----------------\n\n")
}

// iterateFolder walks the folder provided in [folderPath] and launches a filereader routine for each excel file it finds.
// If the recursive flag is provided [-r] it also calls itself upon all subdirectories of [folderPath].
func iterateFolder(c Config, readwg *sync.WaitGroup, routineLimiter chan int, rowsch chan<- []string) {
	filepath.WalkDir(c.FolderPath, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			fmt.Println(err)
			return err
		}

		// Recursive call if user want to read files in subdirectories
		if d.IsDir() && c.Recursive && c.FolderPath != path {
			newConfig := c
			newConfig.FolderPath = path
			iterateFolder(newConfig, readwg, routineLimiter, rowsch)
			return nil
		}

		// Make sure we're in the folderpath provided by the function caller.
		if filepath.Dir(path) != c.FolderPath {
			return nil
		}

		// Check if file is .xlsx or .xlsm
		if filepath.Ext(path) != ".xlsx" && filepath.Ext(path) != ".xlsm" {
			fmt.Println("Skipping file:", "["+d.Name()+"]", " extension:", filepath.Ext(path))
			return nil
		}

		readwg.Add(1)
		routineLimiter <- 1
		go fileReader(path, c.SheetName, rowsch, readwg, routineLimiter)
		return nil
	})
}

// fileReader opens and parses the rows of an Excel document sheet into the channel [ch].
// If a specific sheet name is not provided, it processes the first sheet of each file.
// The amount of concurrent fileReaders is limited by the channel [limiter].
func fileReader(filename string, sheetName string, ch chan<- []string, wg *sync.WaitGroup, limiter chan int) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println("Unable to open file: ", err)
	}
	defer f.Close()

	sheets := f.GetSheetMap()

	// Check if sheetName is populated.
	if sheetName == "" {
		sheetName = sheets[1]
	} else if !checkIfSheetExists(sheetName, sheets) {
		wg.Done()
		fmt.Printf("Unable to find sheet: [%v]\tin file: [%v]\t\tSkipping file..\n", sheetName, filepath.Base(filename))
		return
	}

	rows, err := f.Rows(sheetName)
	if err != nil {
		fmt.Println("unable to get rows from sheet: ", err)
	}

	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println("unable to get row columns, ", err)
		}
		ch <- row
	}
	err = rows.Close()
	if err != nil {
		fmt.Println("unable to close file: ", err)
	}
	<-limiter
	wg.Done()
}

// fileWriter creates and writes a CSV file from the rows received on the input channel [ch].
// The output file is named based on the provided configuration and the current timestamp.
func fileWriter(ch <-chan []string, c Config, wg *sync.WaitGroup) {
	// Make sure path has a trailing slash and create file
	if c.FolderPath[len(c.FolderPath)-1:] != "/" {
		c.FolderPath = c.FolderPath + "/"
	}
	f, err := os.Create(c.FolderPath + c.OutputName + " " + time.Now().Format("2006-01-02 15_04_05") + ".csv")
	if err != nil {
		log.Fatal("Unable to create file: ", err)
	}
	defer f.Close()

	// initialize csv writer
	writer := csv.NewWriter(f)
	writer.Comma = c.Delimiter
	rowCount := 0

	// write rows to file
	for row := range ch {
		if rowCount%1000 == 0 {
			fmt.Println("Rows processed: ", rowCount)
			writer.Flush()
		}
		writer.Write(row)
		rowCount++
	}
	fmt.Println("Rows processed: ", rowCount)
	writer.Flush()
	wg.Done()
}

// checkIfSheetExists takes a sheetname and checks if that exists within a map of sheets.
func checkIfSheetExists(sheetName string, sheets map[int]string) bool {
	for _, v := range sheets {
		if v == sheetName {
			return true
		}
	}
	return false
}

// promptuserforpath is used when a user excludes the -p flag. A folder path is mandatory for the script to function.
func promptuserforpath(prompt string) string {
	var userInput string
	scanner := bufio.NewScanner(os.Stdin)
	fmt.Print(prompt)
	scanner.Scan()
	userInput = scanner.Text()
	return userInput
}

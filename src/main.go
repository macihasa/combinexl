package main

import (
	"C"
	"bufio"
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Flags and setup
	var maxNumGoroutines int
	var sheetName, outputName, folderPath, delimiterString string
	var csvDelimiter rune
	var recursive bool

	flag.IntVar(&maxNumGoroutines, "g", 8, "Amount of goroutines reading files at once")
	flag.StringVar(&sheetName, "sn", "", "Name of the specific sheets to parse. (Optional)")
	flag.StringVar(&outputName, "o", "Output", "Name of output file.")
	flag.StringVar(&folderPath, "p", "", "Path to directory containing excel files to parse.")
	flag.StringVar(&delimiterString, "d", ";", "Csv delimiter for the output file. Maximum 1 letter.")
	flag.BoolVar(&recursive, "r", false, "Recursively goes through each subdirectory and iterates through their excel documents.")

	flag.Parse()

	if len(delimiterString) > 1 || len(delimiterString) == 0 {
		fmt.Print("The csv delimiter provided through [-d] can only be 1 character long. Input provided: [" + delimiterString + "]")
		return
	}

	for _, v := range delimiterString {
		csvDelimiter = v
	}

	if folderPath == "" {
		folderPath = promptuserforpath("Enter path to directory: ")
	}

	fmt.Printf("\n----Variables----\nmaxNumGoroutines:\t%v\nsheetName:\t\t%v\noutputName:\t\t%v\nfolderPath:\t\t%v\ncsvDelimiter:\t\t%s\nrecursive:\t\t%v\n-----------------\n\n", maxNumGoroutines, sheetName, outputName, folderPath, delimiterString, recursive)

	// Start of processing
	rowsch := make(chan []string, 1024)

	readwg := new(sync.WaitGroup)
	writewg := new(sync.WaitGroup)

	routineLimiter := make(chan int, maxNumGoroutines)

	writewg.Add(1)
	go fileWriter(rowsch, folderPath, outputName, csvDelimiter, writewg)

	iterateFolder(folderPath, sheetName, readwg, rowsch, routineLimiter, recursive)

	readwg.Wait()
	close(rowsch)
	writewg.Wait()
}

func iterateFolder(folderPath, sheetName string, readwg *sync.WaitGroup, rowsch chan<- []string, routineLimiter chan int, recursive bool) {
	filepath.WalkDir(folderPath, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			fmt.Println(err)
			return err
		}

		// Recursive call if user want to read files in subdirectories
		if d.IsDir() && recursive && folderPath != path {
			iterateFolder(path, sheetName, readwg, rowsch, routineLimiter, recursive)
			return nil
		}

		// Make sure we're in the folderpath provided by the function caller.
		if filepath.Dir(path) != folderPath {
			return nil
		}

		// Check if file is .xlsx or .xlsm
		if filepath.Ext(path) != ".xlsx" && filepath.Ext(path) != ".xlsm" {
			fmt.Println("Skipping file:", "["+d.Name()+"]", " extension:", filepath.Ext(path))
			return nil
		}

		readwg.Add(1)
		routineLimiter <- 1
		go fileReader(path, sheetName, rowsch, readwg, routineLimiter)
		return nil
	})
}

// fileReader opens and parses the rows of an excel document sheet into the channel ch.
// If the user does not provide a specific sheet name through [-sn], the first sheet of each file is processed instead.
func fileReader(filename string, sheetName string, ch chan<- []string, wg *sync.WaitGroup, limiter chan int) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println("Unable to open file: ", err)
	}

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

// Filewriter creates and writes a csv file from the rows channel ch.
func fileWriter(ch <-chan []string, path string, fileName string, delimiter rune, wg *sync.WaitGroup) {
	// Make sure path has a trailing slash and create file
	if path[len(path)-1:] != "/" {
		path = path + "/"
	}
	f, err := os.Create(path + fileName + " " + time.Now().Format("2006-01-02 15_04_05") + ".csv")
	if err != nil {
		log.Fatal("Unable to create file: ", err)
	}
	defer f.Close()

	// initialize csv writer
	writer := csv.NewWriter(f)
	writer.Comma = delimiter
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

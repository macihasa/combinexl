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

// Config contains the flag properties provided by the user.
type Config struct {
	MaxNumReaders  int
	SheetName      string
	StartsWith     string
	OutputFileName string
	OutputFilePath string
	FolderPath     string
	Delimiter      rune
	Recursive      bool
	HistoricalData bool
}

func main() {
	// Flags and setup
	config := parseFlags()
	printFlags(config)

	// Rows channel is used for transmission of data between the filereaders and the filewriter
	rowsch := make(chan []string, 1024)

	// Wait groups
	readwg := new(sync.WaitGroup)
	writewg := new(sync.WaitGroup)

	// Launch filewriter routine
	writewg.Add(1)
	go fileWriter(rowsch, config, writewg)

	// Retrieve filepaths/names
	filenames := []string{}
	iterateFolder(config, &filenames)

	// Add filenames to filenames channel and close it when done
	filenamesch := make(chan string, len(filenames))
	for _, v := range filenames {
		filenamesch <- v
	}
	close(filenamesch)

	// Launch filereader routines
	if len(filenames) < config.MaxNumReaders {
		for i := 0; i < len(filenames); i++ {
			readwg.Add(1)
			go fileReader(filenamesch, config, rowsch, readwg)
		}
	} else {
		for i := 0; i < config.MaxNumReaders; i++ {
			readwg.Add(1)
			go fileReader(filenamesch, config, rowsch, readwg)
		}
	}

	readwg.Wait()
	close(rowsch)
	writewg.Wait()

	if config.HistoricalData {
		moveFilesToFolder(filenames, "Historical_Data")
	}

}

// iterateFolder walks the folder provided in [folderPath] and adds all .
// If the recursive flag is provided [-r] it also calls itself upon all subdirectories of [folderPath].
func iterateFolder(c Config, filesToRead *[]string) {
	filepath.WalkDir(c.FolderPath, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			fmt.Println(err)
			return err
		}

		// Recursive call if user want to read files in subdirectories
		if d.IsDir() && c.Recursive && c.FolderPath != path {
			newConfig := c
			newConfig.FolderPath = path
			iterateFolder(newConfig, filesToRead)
			return nil
		}

		// Make sure we're in the folderpath provided by the function caller.
		if filepath.Dir(path) != c.FolderPath {
			return nil
		}

		// Check if file is .xlsx or .xlsm
		if filepath.Ext(path) != ".xlsx" && filepath.Ext(path) != ".xlsm" {
			if d.IsDir() {
				fmt.Println("Skipping directory:", "["+d.Name()+"]")
				return nil
			}
			fmt.Println("Skipping file:", "["+d.Name()+"]", " extension:", filepath.Ext(path))
			return nil
		}

		// Check if startswith flag is populated and if the file name starts with that string
		if c.StartsWith != "" && d.Name()[:len(c.StartsWith)] != c.StartsWith {
			fmt.Println("Skipping file:", "["+d.Name()+"]", "does not start with: ["+c.StartsWith+"]")
			return nil
		}

		// Add filepath to filestoread.
		*filesToRead = append(*filesToRead, path)
		return nil
	})
}

// fileReaders opens and parses the rows of an Excel document which location is provided by the [filenames] channel.
// It sends the rows into the rows channel [ch] for the filewriter.
// If a specific sheet name is not provided, it processes the first sheet of each file.
func fileReader(filenames chan string, c Config, ch chan<- []string, wg *sync.WaitGroup) {
	for v := range filenames {
		f, err := excelize.OpenFile(v)
		if err != nil {
			fmt.Println("Unable to open file: ", err)
		}
		sheets := f.GetSheetMap()
		sheetName := c.SheetName

		// Check if sheetName is populated. If not set it to the first sheet in the file
		if sheetName == "" {
			sheetName = sheets[1]
		} else if !checkIfSheetExists(sheetName, sheets) {
			fmt.Printf("Unable to find sheet: [%v]\tin file: [%v]\t\tSkipping file..\n", sheetName, filepath.Base(v))
			continue
		}
		// Get rows iterator from sheet
		rows, err := f.Rows(sheetName)
		if err != nil {
			fmt.Println("unable to get rows from sheet: ", err)
			continue
		}

		// Iterate over rows and send them to the rows channel for the writer to write
		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				fmt.Println("unable to get row columns, ", err)
			}
			ch <- row
		}
		err = rows.Close()
		if err != nil {
			fmt.Println("unable to close rows iterator: ", err)
			continue
		}

		err = f.Close()
		if err != nil {
			fmt.Println("unable to close file", err)
			continue
		}
	}
	wg.Done()
}

// fileWriter creates and writes a CSV file from the rows received on the input channel [ch].
// The output file is named based on the provided configuration and the current timestamp.
func fileWriter(ch <-chan []string, c Config, wg *sync.WaitGroup) {
	// Make sure path has a trailing slash and create file
	outputFileName := c.OutputFileName + " " + time.Now().Format("2006-01-02 15_04_05") + ".csv"
	outputFilePath := filepath.Join(c.FolderPath, outputFileName)

	// If [op] flag is populated, replace outputfilepath.
	if c.OutputFilePath != "" {
		outputFilePath = filepath.Join(c.OutputFilePath, outputFileName)
	}

	f, err := os.Create(outputFilePath)

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
	fmt.Println("Rows processed: ", rowCount, "- Finished")
	writer.Flush()
	wg.Done()
}

// moveFilesToFolder moves files specified in [filenames] to a new folder named [foldername].
func moveFilesToFolder(filenames []string, foldername string) {
	if len(filenames) == 0 {
		fmt.Println("No files found.")
		return
	}
	err := os.Mkdir(filepath.Join(filepath.Dir(filenames[0]), foldername), 0755)
	if err != nil {
		fmt.Println("unable to create dir", foldername, err)
	}
	for _, v := range filenames {
		err = os.Rename(v, filepath.Join(filepath.Dir(v), foldername, filepath.Base(v)))
		if err != nil {
			fmt.Println("unable to move files.", err)
		}
	}
}

// parseFlags parses command-line flags and returns a Config with user-specified options.
// It handles options for concurrent readers, sheet name, output file, folder path, CSV delimiter, and recursion.
// Invalid flags or missing folder path terminate the program.
func parseFlags() Config {
	var config Config
	var delimiterString string
	flag.StringVar(&config.FolderPath, "p", "", "Path to the directory containing Excel files to parse. (Required).")
	flag.StringVar(&config.SheetName, "sn", "", "Specify the target sheet name in Excel files. (Defaults to first sheet)")
	flag.StringVar(&config.StartsWith, "sw", "", "Filters files to only include those whose names start with the specified string.")
	flag.StringVar(&config.OutputFileName, "o", "Output", "Sets the name of the output CSV file. ")
	flag.StringVar(&config.OutputFilePath, "op", "", "Sets directory location of output file")
	flag.StringVar(&delimiterString, "d", ";", "Sets the CSV delimiter for the output file. Must be a single character.")
	flag.IntVar(&config.MaxNumReaders, "g", 8, "Limits the number of concurrent file readers.")
	flag.BoolVar(&config.Recursive, "r", false, "Enables recursive processing of subdirectories and all excel files within them.")
	flag.BoolVar(&config.HistoricalData, "h", false, "Option for used files to be moved to a historical data folder after usage.")

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

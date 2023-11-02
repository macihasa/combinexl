package main

import (
	"C"
	"bufio"
	"encoding/csv"
	"log"
	"os"
	"path/filepath"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

const MAX_NUM_GOROUTINES = 8

// This script combines the first sheet of different Excel files into a consolidated Output.csv file.
func main() {
	root := promptuserforpath("Enter path to directory: ")

	rowsch := make(chan []string, 1024)

	readwg := new(sync.WaitGroup)
	writewg := new(sync.WaitGroup)

	routineLimiter := make(chan int, MAX_NUM_GOROUTINES)

	go fileWriter(rowsch, root, writewg)

	var filecount int

	filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			log.Println(err)
			return err
		}
		if d.IsDir() {
			return nil
		}
		// Check if file is .xlsx or .xlsm
		if filepath.Ext(path) == ".xlsx" || filepath.Ext(path) == ".xlsm" {

		} else {
			log.Println("Skipping file:", d.Name(), " extension:", filepath.Ext(path))
			return nil
		}

		filecount++

		if filecount%100 == 0 {
			log.Println(filecount)
		}
		readwg.Add(1)
		routineLimiter <- 1
		go fileReader(path, rowsch, readwg, routineLimiter)

		return nil
	})
	readwg.Wait()
	close(rowsch)
	writewg.Wait()
	_ = promptuserforpath("Press enter to continue")

}

func fileReader(filename string, ch chan<- []string, wg *sync.WaitGroup, limiter chan int) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Println("Unable to open file: ", err)
	}

	sheets := f.GetSheetMap()

	rows, err := f.Rows(sheets[1])
	if err != nil {
		log.Println("not able to process row: ", err)
	}

	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			log.Println("unable to get row columns, ", err)
		}
		ch <- row
	}
	err = rows.Close()
	if err != nil {
		log.Println("unable to close file: ", err)
	}
	<-limiter
	wg.Done()
}

func fileWriter(ch <-chan []string, path string, wg *sync.WaitGroup) {
	// Make sure path has a trailing slash and create file
	if path[len(path)-1:] != "/" {
		path = path + "/"
	}
	f, err := os.Create(path + "Output " + time.Now().Format("2006-01-02 15_04_05") + ".csv")
	if err != nil {
		log.Fatal("Unable to create file: ", err)
	}
	defer f.Close()

	// initialize csv writer
	writer := csv.NewWriter(f)
	writer.Comma = rune(';')
	rowCount := 1

	// write rows to file
	for row := range ch {
		if rowCount%1000 == 0 {
			log.Println("Rows processed: ", rowCount)
			writer.Flush()
		}
		writer.Write(row)
		rowCount++
	}
	log.Println("Rows processed: ", rowCount)
	writer.Flush()
	wg.Done()
}

func promptuserforpath(prompt string) string {
	var userInput string
	scanner := bufio.NewScanner(os.Stdin)
	log.Print(prompt)
	scanner.Scan()
	userInput = scanner.Text()
	return userInput
}

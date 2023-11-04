# CombineXL - Excel Data Extraction Tool

CombineXL is a command line tool for extracting data from Excel files in a specified directory and combining it into a CSV file. It allows you to customize various parameters for the extraction process.

## Usage

```
$ combinexl -p "path_to_directory" -sn "sheet_name" -o "output_file_name"
```

## Command Line Options

- `-p`: Path to the directory containing Excel files to parse.
- `-sn`: Specify the target sheet name in Excel files. (Defaults to first sheet of each file).
- `-o`: Sets the name of the output CSV file.
- `-d`: Sets the CSV delimiter for the output file. Must be a single character. (defaults to `;`).
- `-g`: Limits the number of concurrent file readers (defaults to `8`).
- `-r`: Enables recursive processing of subdirectories and all excel files within them.

## Example

```
`$ combinexl -p "C:\Users\testing\Year to date\" -sn "TestSheet" -o "OutputFileName"`
```

Resulting logs:

```

Variables
maxNumGoroutines:       8
sheetName:              TestSheet
outputName:             OutputFileName
folderPath:             C:\Users\testing\Year to date\
csvDelimiter:           ;
recursive:   		false


Rows processed:  0
Rows processed:  1000
Skipping file: [Output 2023-10-23 16_07_05.csv]  extension: .csv
Unable to find sheet: [TestSheet]        in file: [SomeFile.xlsx]            Skipping file..
Rows processed:  2000
Rows processed:  2778 - Finished

```

## Installation

To install CombineXL, you'll need to have Go (Golang) installed on your system. If you don't have it already, you can download and install it from the official [Go website](https://golang.org/dl/).

Once you have Go installed, you can follow these steps to install CombineXL:

### Option 1: Building from Source

1. Clone the CombineXL repository to your local machine:

   ```shell
   git clone https://github.com/macihasa/combinexl.git
   ```
2. Navigate to the CombineXL directory:

   ```shell
   cd combinexl
   ```
3. Build the CombineXL executable:

   ```shell
   go build -o combinexl main.go
   ```

Now, you should have the combinexl executable in the same directory.

You can optionally move the combinexl executable to a location in your system's PATH to make it globally accessible. For example:

```shell
mv combinexl /usr/local/bin/
```

This step is optional but recommended for convenience.

### Option 2: Using go install

Install CombineXL using go install:

```shell
go install github.com/macihasa/combinexl@latest
```

This will install the combinexl executable in your Go binary directory.

You can then run combinexl from anywhere in your terminal.

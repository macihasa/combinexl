# CombineXL - Excel Data Extraction Tool

CombineXL is a command line tool for extracting data from Excel files in a specified directory and combining it into a CSV file. It allows you to customize various parameters for the extraction process.

## Usage

```shell
$ combinexl -p "path_to_directory" -sn "sheet_name" -o "output_file_name"

Command Line Options
-p: Path to the directory containing Excel files to parse.
-sn: Name of the specific Excel sheet(s) to parse (optional).
-o: Name of the output CSV file.
-d: CSV delimiter for the output file (default is ;).
-g: Amount of goroutines for reading files at once (default is 8).

Example
$ combinexl -p "path/to/excel/files" -sn "Shipment" -o "Batch_Summary"

Variables
maxNumGoroutines: Number of goroutines used for reading files.
sheetName: Name of the specific Excel sheet to parse.
outputName: Name of the output CSV file.
folderPath: Path to the directory containing Excel files.
csvDelimiter: CSV delimiter used in the output file.
Output
The tool provides feedback on the number of rows processed and can skip files that don't contain the specified sheet.

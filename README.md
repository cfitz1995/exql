# exql

This is a command line project which can be used to read in Excel files and generate SQL insert statements.

## Getting Started

These instructions will get a copy of the project up and running on your local machine. Please note that this has only been tested on a Windows 10 machine.

### Installation

Install globally via NPM to access via command line.

```
npm i -g exql
```
### Usage

There is currently only one command provided by exql, run.

This will read in the Excel file at the given path, using the name of each sheet as the table to insert to, and generate indiviual SQL insert statements for each row.

Row 1 of each worksheet must contain column headers corresponding to table columns. Prepend a column header with # to instruct exql that it's a numeric field and it should not to wrap that column's value in quotes.

Run accepts the following options:

```
-i --input  [REQUIRED] - Full path to the Excel file
-o --output [OPTIONAL] - Full path of the output file
-q --quiet  [OPTIONAL] - Suppress console output
-f --force  [OPTIONAL] - Overwrite the output file if it already exists
-h --help   [OPTIONAL] - Help
```

#### Example

```
exql run -i C:\Users\cfitz1995\Desktop\myExcelFile.xlsx -o C:\Users\cfitz1995\Desktop\mySqlOutput.sql -fq
```

## Authors

* **Connor Fitzgerald** - *Initial work*
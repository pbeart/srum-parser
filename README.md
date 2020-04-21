# srum-parser

A tool to assist with parsing the SRUDB.dat ESE database

## Installation
1. Get the repository on your computer by cloning the repository or downloading as a zip file and unzipping to a sensible place
1. Install the dependencies in requirements.txt with `pip install -r requirements.txt`
1. Install [pyesedb](https://www.google.com) by cloning the repository then building the Python bindings using the instructions [here](https://github.com/libyal/libesedb/wiki/Building#using-setuppy)

## Usage

Usage: cli.py [OPTIONS]

### Options:  
  -i, --input FILE                [required]  
  -o, --output FILE               [required]  
  -p, --omit-processed  
  -P, --only-processed  
  --help                          Show this message and exit.

### --input:  
The input SRUDB.dat file to use

### --output:  
The output .xlsx file to write to

### --omit-processed:
Whether to skip the secondary processing of the tables

### --only-processed
Whether to skip outputting the 'raw' tables

## Examples
### Output both the raw and processed tables from `SRUDB.dat` to `output.xlsx`
`./cli.py --input SRUDB.dat --output output.xlsx`

### Output only the raw tables from `SRUDB.dat` to `output.xlsx`
`./cli.py --input SRUDB.dat --output output.xlsx --omit-processed`

### Output only the processed tables from `SRUDB.dat` to `output.xlsx`
`./cli.py --input SRUDB.dat --output output.xlsx --only-processed`
# README for `pqc_arks`

This is a script to prepopulate EZID ark IDs and some boilerplate metadata in a spreadsheet for PQC-compliant objects, from a source CSV or XLSX file.

## Requirements
* Ruby 2.5.1 or above
* EZID client credentials (for minting ark IDs for production)

## Setup

1. Install Ruby dependencies:

  ```bash 
  $ bundle install
  ```

2. *** ***If you are minting ark IDs for production*** ***, source environment variables for the EZID account credentials.

  ```bash
  $ export EZID_DEFAULT_SHOULDER='$SHOULDER';
  $ export EZID_USER='$USERNAME';
  $ export EZID_PASSWORD='$PASSWORD';
  ```

  Where `$SHOULDER`, `$USERNAME`, and `$PASSWORD` are the EZID account values for production.

## Usage

To generate a spreadsheet and mint one ARK identifier per row, which will be included in the spreadsheet:
```
bundle exec ruby pqc_arks.rb $NUMBER_OF_ARKS
```

Where `$NUMBER_OF_ARKS` is an integer value specifying the number of ark IDs you want to create at this time, in one sheet.

New values will be saved to a spreadsheet either called `$<NUMBER_OF_ARKS>_POPULATED.xlsx` in the same directory from which you ran the script, or at a path and filename optionally specified as a second argument.

Example:

```
bundle exec ruby pqc_arks.rb $NUMBER_OF_ARKS output_file.xlsx
```

The spreadsheet containing the ark IDs in the written in the above example would be `output_file.xlsx`, located in the same directory from which you ran the script.

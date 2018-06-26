## README for prepop_rows

This is a script to prepopulate a rows-based metadata spreadsheet for PQC from a source CSV file.

### Requirements
* Ruby 2.2.5 or above
* rubyXL gem
* SmarterCSV gem
* EZID client credentials (for minting true arks)

### Usage

To generate a spreadsheet without minting ARK identifiers:
```
ruby prepop_rows.rb source_file.csv
```

To generate a spreadsheet and mint one ARK identifier per row (this will be included in the spreadsheet):
```
ruby prepop_rows.rb source_file.csv --ark
```

or 
```
ruby prepop_rows.rb source_file.csv -a
```

For help on the command line:
```
ruby prepop_rows.rb -h
```

New values will be saved to a spreadsheet either called `default.xlsx` in the same directory from which you ran the script, or at a path and filename optionally specified as a second argument.

### Validation checks

This script performs the following validation checks as each row of data is created:

* Ark identifier and directory name matches
* Title value is present
* Filename(s) value is present

Any errors detected are displayed in the terminal at the end of the run.
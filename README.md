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

New values will be saved to a spreadsheet called `default.xlsx` in the same directory from which you ran the script.
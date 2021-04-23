# Bulk Import CSV Generator`

This is a script to pre-populate EZID ark IDs in a bulk import spreadsheet.

## Requirements
* Ruby 2.7.1 or above
* EZID client credentials (for minting ark IDs for production)

## Setup

1. Install Ruby dependencies:

  ```bash 
  $ bundle install
  ```

2. *** ***If you are minting ark IDs for production*** ***, source environment variables for the EZID account credentials. Add the variables to `config/ezid.yml`.

  ```yml
  default_shoulder: '$SHOULDER'
  user: '$USERNAME'
  password: '$PASSWORD'
  ```

  Where `$SHOULDER`, `$USERNAME`, and `$PASSWORD` are the EZID account values for production.

## Usage

```bash
Usage: bulk_import_csv.rb --lines=LINES --output=FILEPATH --dryrun
    -o, --output FILEPATH            Output filepath for csv, required
    -l, --lines LINES                Number of lines, required
    -d, --dry-run                    Creates spreadsheet with arks from EZID Test API
    -h, --help                       Prints this help
```

#### To generate a spreadsheet and mint one ARK identifier per row:
```
ruby bulk_import_csv.rb --lines 100 --output /home/important/place/output_file.csv
```

A bulk import spreadsheet with 100 minted arks will be creates and written to `/home/important/place/output_file.csv`. The arks will be minted with production credentials.

#### To test the script:
```
ruby bulk_import_csv.rb --lines 5 --output /home/important/place/output_file.csv --dry-run
```

A bulk import spreadsheet with 5 test arks will be created and written to `/home/important/place/output_file.csv`. The arks will be minted using the public test EZID credentials.

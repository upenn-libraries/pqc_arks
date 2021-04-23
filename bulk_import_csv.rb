# Script to generate a bulk import spreadsheet with a minted ark for each row.

require 'optparse'
require 'csv'
require 'yaml'
require 'ezid-client'

options = {}
OptionParser.new do |opts|
  opts.banner = "Usage: bulk_import_csv.rb --lines=LINES --output=FILEPATH --dryrun"

  opts.on("-oFILEPATH", "--output FILEPATH", "Output filepath for csv, required") do |output|
    options[:output] = output
  end

  opts.on("-lLINES", "--lines LINES", "Number of lines, required") do |lines|
    options[:lines] = lines.to_i
  end

  opts.on("-d", "--dry-run", "Creates spreadsheet with arks from EZID Test API") do |dry_run|
    options[:dry_run] = dry_run
  end

  opts.on("-h", "--help", "Prints this help") do
    puts opts
    exit
  end
end.parse!


##### Helper methods

def headers
  metadata_fields =  [
    'item_type', 'abstract', 'call_number', 'collection', 'contributor',
    'corporate_name', 'coverage', 'creator', 'date', 'description', 'format',
    'geographic_subject', 'identifier', 'includes', 'language', 'notes',
    'personal_name', 'provenance', 'publisher', 'relation', 'rights', 'source',
    'subject', 'title'
  ]

  structural_fields = ['filenames']

  list = ['unique_identifier', 'action', 'directive_name']
  list.concat metadata_fields.map { |f| "metadata.#{f}[1]" }
  list.concat structural_fields.map { |f| "structural.#{f}" }
  list
end

# Returns a minted ark.
def mint_ark
  Ezid::Identifier.mint
end

def ezid_credentials(test = false)
  if test
    { 'default_shoulder' => 'ark:/99999/fk4', 'user' => 'apitest', 'password' => 'apitest' }
  else
    config = YAML.load(File.read("config/ezid.yml"))
    raise 'Missing EZID credentials' unless config['default_shoulder'] && config['user'] && config['password']
    config
  end
end


##### Start of script.

# Check that required parameters are provided
raise '--lines must be provided' unless options[:lines]
raise '--output must be provided' unless options[:output]

output_path = File.absolute_path(options[:output])

raise "#{output_path} does not exist" unless File.exist?(File.dirname(output_path))

# Set up EZID connection
credentials = ezid_credentials(options[:dry_run])
Ezid::Client.configure do |conf|
  conf.default_shoulder = credentials['default_shoulder']
  conf.user = credentials['user']
  conf.password = credentials['password']
end

# Generate CSV
CSV.open(options[:output], 'wb') do |csv|
  csv << headers
  options[:lines].times { csv << [mint_ark,'create'] }
end
#!/usr/bin/ruby

require 'optparse'
require 'smarter_csv'
require 'rubyXL'
require 'ezid-client'

def missing_args?
  return (ARGV[0].nil?)
end

def set_up_spreadsheet(workbook, headers)
  worksheet = workbook.worksheets[0]
  set_headers(worksheet, headers)
end

def set_headers(worksheet, headers_hash)
  headers_hash.each_with_index do |(key, value), index|
    worksheet.add_cell(0,index, value)
  end
end

def mint_arkid
  identifier = Ezid::Identifier.mint
  return identifier, identifier.to_s.gsub(/[:\/]/, ':' => '+', '/' => '=')
end

abort('Specify a path to a CSV file') if missing_args?

Ezid::Client.configure do |conf|
  conf.default_shoulder = 'ark:/99999/fk4' unless ENV['EZID_DEFAULT_SHOULDER']
  conf.user = 'apitest' unless ENV['EZID_USER']
  conf.password = 'apitest' unless ENV['EZID_PASSWORD']
end

QUALIFIED_HEADERS = { :type => 'Type',
                      :unique_identifier => 'Unique Identifier',
                      :abstract => 'Abstract',
                      :call_number => 'Call Number',
                      :collection_name => 'Collection Name',
                      :contributor => 'Contributor',
                      :corporate_name => 'Corporate Name',
                      :coverage => 'Coverage',
                      :creator => 'Creator',
                      :date => 'Date',
                      :raw_description => 'Raw Description',
                      :description => 'Description',
                      :format => 'Format',
                      :geographic_subject => 'Geographic Subject',
                      :identifier => 'Identifier',
                      :includes => 'Includes',
                      :language => 'Language',
                      :notes => 'Notes',
                      :personal_name => 'Personal Name',
                      :provenance => 'Provenance',
                      :publisher => 'Publisher',
                      :relation => 'Relation',
                      :rights => 'Rights',
                      :source => 'Source',
                      :subject => 'Subject',
                      :title => 'Title',
                      :directory_name => 'Directory Name',
                      :filenames => 'Filename(s)' }.freeze

CROSSWALKING_TERMS_SINGLE = { :corporatio => :corporate_name,
                 :date => :date,
                 :descriptio => :raw_description,
                 :descript_1 => :description,
                 :genre_subl => :type,
                 :langcode => :language,
                 :location => :geographic_subject,
                 :tiff_locat => :filenames }.freeze

CROSSWALKING_TERMS_MULTIPLE = { :identifier => [ :thing_uuid,
                                                 :objects_refno,
                                                 :ref_1,
                                                 :ref_2 ],
                                :personal_name => [ :person_nam,
                                                    :person_n_1 ] }.freeze

CROSSWALKING_OPTIONS = { :delimiter => '|' }

BOILERPLATE_TERMS_VALUES = { :collection_name => 'Arnold and Deanne Kaplan collection of Early American Judaica, 1555-1977',
                             :rights => 'http://rightsstatements.org/vocab/UND/1.0/' }

workbook = RubyXL::Workbook.new

def workbook.set_up_spreadsheet
  worksheet = worksheets[0]
  worksheet.sheet_name = 'descriptive'
  set_headers(worksheet, QUALIFIED_HEADERS)
end

def workbook.add_custom_field(y, x, value)
  worksheet = worksheets[0]
  worksheet.add_cell(y, x, value)
end

def workbook.prepop(dataset, opts = {})

  multi_values = {}
  worksheet = worksheets[0]

  dataset.each_with_index do |row, y_index|

    QUALIFIED_HEADERS.each_with_index do  |(key, values), x|
      worksheet.add_cell(y_index+1, x, row[key]) unless row[key].nil?
    end

    CROSSWALKING_TERMS_MULTIPLE.each do |key, values|
      multi_values[key] = ""
      values.each do |value|
        multi_values[key].concat("#{row[value]}#{CROSSWALKING_OPTIONS[:delimiter]}") unless row[value].nil?
      end
    end

    multi_values.each do |key, value|
      worksheet.add_cell(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == key }, value) unless value.empty?
    end

    BOILERPLATE_TERMS_VALUES.each do |key, value|
      worksheet.add_cell(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == key }, value)
    end

    if opts[:ark]
      identifier, directory = mint_arkid
      add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :unique_identifier }, identifier)
      add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :directory_name }, directory)
    end

  end
end

flags = {}
OptionParser.new do |opts|
  opts.banner = "Usage: prepop_rows.rb [options] FILE"

  opts.on("-a", "--ark", "Mint arks, one per row in output spreadsheet") do |a|
    flags[:ark] = a
  end
end.parse!
spreadsheet_name = 'default.xlsx'

csv = ARGV[0]
options = { :encoding => 'ISO8859-1:utf-8', :key_mapping => CROSSWALKING_TERMS_SINGLE }
csv_parsed = SmarterCSV.process(csv, options)

puts 'Writing spreadsheet...'
workbook.set_up_spreadsheet
workbook.prepop(csv_parsed, flags)
workbook.write(spreadsheet_name)
puts "Spreadsheet written to #{spreadsheet_name}."

#!/usr/bin/ruby

require 'optparse'
require 'smarter_csv'
require 'rubyXL'
require 'ezid-client'
require 'active_support/inflector'

require 'pry'

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
  return identifier, directorify_ark(identifier)
end

def directorify_ark(ark)
  return ark.to_s.gsub(/[:\/]/, ':' => '+', '/' => '=') unless ark.nil?
end

def rollup(header, row, multi_headers=[])

  term = ''
  score = 0
  ROLLUP_TERMS[header].each do |key|
    if row[key].to_s.empty?
      score += 1
    else
      if key == :type
        value = "#{row[key]}".gsub(/\|$/,'').singularize
      elsif key == :geographic_subject
        value = row[key].split('|').max_by(&:length)
      elsif multi_headers.include? key
        value = row[key] ? row[key].to_s.split('|').join('; ') : ''
      else
        value = "#{row[key]}"
      end
      term << "#{value.strip}; "
    end
  end
  return score <= 5 ? term : ''
end

def validation_check(row)
  output = []
  output << "Ark ID and directory names do not match: #{row[:unique_identifier]} and #{row[:directory_name]}" unless directorify_ark(row[:unique_identifier]) == row[:directory_name]
  output << "Missing required value title: #{row[:unique_identifier]}" if row[:title].nil? || row[:title].empty?
  output << "Missing required value filename(s): #{row[:unique_identifier]}" if row[:'filename(s)'].nil? || row[:'filename(s)'].empty?
  return output
end

abort('Specify a path to a CSV or XLSX file or a number of blank rows with arks to mint') if missing_args?

Ezid::Client.configure do |conf|
  conf.default_shoulder = 'ark:/99999/fk4' unless ENV['EZID_DEFAULT_SHOULDER']
  conf.user = 'apitest' unless ENV['EZID_USER']
  conf.password = 'apitest' unless ENV['EZID_PASSWORD']
end

QUALIFIED_HEADERS = { :type => 'Type',
                      :unique_identifier => 'Unique Identifier',
                      :abstract => 'Abstract',
                      :call_number => 'Call Number',
                      :citation_note => 'Citation Note',
                      :collection_name => 'Collection Name',
                      :container => 'Container',
                      :contributor => 'Contributor',
                      :corporate_name => 'Corporate Name',
                      :coverage => 'Coverage',
                      :creator => 'Creator',
                      :date => 'Date',
                      :description => 'Description',
                      :format => 'Format',
                      :geographic_subject => 'Geographic Subject',
                      :identifier => 'Identifier',
                      :'collectify_identifier(s)' => 'Collectify Identifier(s)',
                      :object_refno => 'Object Refno',
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
                      :'filename(s)' => 'Filename(s)',
                      :first_filename => 'First filename',
                      :notes2 => 'Notes2',
                      :done => 'Done',
                      :status => 'Status',
                      :duplicates => 'Duplicates',
                      :master => 'Master'}.freeze

CROSSWALKING_TERMS_SINGLE = { :corporate_name => :corporate_name,
                              :corporatio => :corporate_name,
                              :formatted_date => :date,
                              :description_1 => :description,
                              :descript_1 => :description,
                              :genre_sublevel => :type,
                              :genre_subl => :type,
                              :type => :type,
                              :langcode => :language,
                              :location => :geographic_subject,
                              :geographic_name => :geographic_subject,
                              :format => :format,
                              :description => :description,
                              :container => :container,
                              :first_filename => :first_filename,
                              :'filename(s)' => :filenames}.freeze

CROSSWALKING_TERMS_MULTIPLE = { :collectify_identifiers => [ :arny_thing_uuid,
                                                             :arny_objects_refno,
                                                             :other_id ],
                                :identifier => [ :ref,
                                                 :cid,
                                                 :ref_1,
                                                 :ref_2,
                                                 :arny_thing_uuid,
                                                 :arny_objects_refno,
                                                 :obj_id,
                                                 :loan_id,
                                                 :identifier,
                                                 :'collectify_identifier(s)' ],
                                :physical_type => [ :object_type,
                                                    :material  ],
                                :personal_name => [ :personal_name,
                                                    :person_nam,
                                                    :person_n_1 ] }.freeze

# headers that may can contain more than on value
MULTIPLE_HEADERS = CROSSWALKING_TERMS_MULTIPLE.values.flatten.freeze

CROSSWALKING_OPTIONS = { :delimiter => '|' }

BOILERPLATE_TERMS_VALUES = { :collection_name => 'Arnold and Deanne Kaplan Collection of Early American Judaica (University of Pennsylvania)',
                             :call_number => 'Arc.MS.56',
                             #:type => 'Ephemera',
                             #:language => 'English',
                             :rights => 'http://rightsstatements.org/page/NoC-US/1.0/?' }

ROLLUP_TERMS = { :title => [:type, :person_nam, :person_n_1, :personal_name, :corporate_name, :geographic_subject, :date] }.freeze

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
  messages = []
  worksheet = worksheets[0]
  dataset.each_with_index do |row, y_index|
    QUALIFIED_HEADERS.each_with_index do  |(key, values), x|
      worksheet.add_cell(y_index+1, x, row[key]) unless row[key].nil?
    end

    CROSSWALKING_TERMS_MULTIPLE.each do |key, values|
      multi_values[key] = ""

      values.each do |value|
        multi_values[key].concat("#{row[value].gsub(/\|$/,'')}#{CROSSWALKING_OPTIONS[:delimiter]}") unless (row[value].nil? || row[value].empty?)
      end
    end

    multi_values.each do |key, value|
      worksheet.add_cell(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == key }, value) unless value.empty?
    end

    BOILERPLATE_TERMS_VALUES.each do |key, value|
      worksheet.add_cell(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == key }, value)
    end

    # Rollup terms

    title = rollup(:title, row, MULTIPLE_HEADERS)
    add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :title }, title)
    row[:title] = title

    if opts[:ark]
      if row[:unique_identifier].nil? || row[:unique_identifier].empty?
        identifier, directory = mint_arkid
        row[:unique_identifier] = identifier.to_s
        row[:directory_name] = directory
      else
        row[:directory_name] = directorify_ark(row[:unique_identifier])
        identifier, directory = row[:unique_identifier], row[:directory]
      end
      add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :unique_identifier }, identifier)
      add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :directory_name }, directory)
    end
    messages << validation_check(row)
  end
  messages.map{|x| puts x}
  puts "All checks made.  Any errors detected are displayed above."
end

def workbook.blank_rows(num_rows)
  worksheet = worksheets[0]
  (0..(num_rows-1)).each do |y_index|
    identifier, directory = mint_arkid
    add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :unique_identifier }, identifier)
    add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :directory_name }, directory)
    BOILERPLATE_TERMS_VALUES.each do |key, value|
      worksheet.add_cell(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == key }, value)
    end
  end
end

def extract_rows(filename)
  if File.extname(filename).downcase == '.csv'
    csv = ARGV[0]
    options = { :encoding => 'ISO8859-1:utf-8', :key_mapping => CROSSWALKING_TERMS_SINGLE }
    contents_array = SmarterCSV.process(csv, options)
  elsif File.extname(filename).downcase == '.xlsx'
    contents_array = []
    xlsx = RubyXL::Parser.parse(filename)
    worksheet = xlsx[0]
    headers = worksheet.sheet_data.rows.first.cells.map do |cell|
      cell.value.downcase.gsub(' ','_').to_sym
    end
    abort('Duplicate column names in use, aborting') if headers.length != headers.uniq.length
    worksheet.sheet_data.rows.each do |row|
      row_hash = {}
      row.cells.each_with_index do |cell, position|
        value = (cell.nil? || cell.value.nil?) ? '' : cell.value.to_s
        row_hash[headers[position]] = value
      end
      contents_array << row_hash
    end
    contents_array.shift
  else
    abort("Unsupported file extension: #{File.extname(filename).downcase}")
  end
  return contents_array
end


flags = {}
OptionParser.new do |opts|
  opts.banner = "Usage: prepop_rows.rb [options] FILE(required) OUTPUTFILE(optional)"

  opts.on("-a", "--ark", "Mint arks, one per row in output spreadsheet") do |a|
    flags[:ark] = a
  end
end.parse!

spreadsheet_name = ARGV[1].nil? ? 'default.xlsx' : "#{File.basename(ARGV[1], '.*')}.xlsx"
workbook.set_up_spreadsheet

num_rows = Integer(ARGV[0]) rescue false

if num_rows && flags[:ark]
  workbook.blank_rows(num_rows)
else
  dataset = extract_rows(ARGV[0])
  workbook.prepop(dataset, flags)
end

workbook.write(spreadsheet_name)
puts "Spreadsheet written to #{spreadsheet_name}."

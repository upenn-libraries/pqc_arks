#!/usr/bin/ruby

require 'optparse'
require 'smarter_csv'
require 'rubyXL'
require 'ezid-client'
require 'active_support/inflector'

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
  return ark.to_s.gsub(/[:\/]/, ':' => '+', '/' => '=')
end

def rollup(header, row)
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
      else
        value = "#{row[key]}"
      end
      term << "#{value}; "
    end
  end
  return score <= 4 ? term : ''
end

abort('Specify a path to a CSV file or a number of blank rows with arks to mint') if missing_args?

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
                      :collectify_identifiers => 'Collectify Identifier(s)',
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
                      :filenames => 'Filename(s)',
                      :first_filename => 'First filename',
                      :full_address => 'full address',
                      :notes2 => 'Notes2',
                      :done => 'Done',
                      :status => 'Status'}.freeze

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
                                                 :"collectify_identifier(s)" ],
                                :physical_type => [ :object_type,
                                                    :material  ],
                                :personal_name => [ :personal_name,
                                                    :person_nam,
                                                    :person_n_1 ] }.freeze


CROSSWALKING_OPTIONS = { :delimiter => '|' }

BOILERPLATE_TERMS_VALUES = { :collection_name => 'Arnold and Deanne Kaplan Collection of Early American Judaica (University of Pennsylvania)',
                             :call_number => 'Arc.MS.56',
                             #:type => 'Ephemera',
                             #:language => 'English',
                             :rights => 'http://rightsstatements.org/vocab/UND/1.0/' }

ROLLUP_TERMS = { :title => [:type, :person_nam, :person_n_1, :corporate_name, :geographic_subject, :date] }.freeze

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
        multi_values[key].concat("#{row[value].gsub(/\|$/,'')}#{CROSSWALKING_OPTIONS[:delimiter]}") unless row[value].nil?
      end
    end

    multi_values.each do |key, value|
      worksheet.add_cell(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == key }, value) unless value.empty?
    end

    BOILERPLATE_TERMS_VALUES.each do |key, value|
      worksheet.add_cell(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == key }, value)
    end

    # Rollup terms

    title = rollup(:title, row)
    add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :title }, title)

    if opts[:ark]
      if row[:unique_identifier].nil?
        identifier, directory = mint_arkid
      else
        identifier, directory = row[:unique_identifier], directorify_ark(row[:unique_identifier])
      end
      add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :unique_identifier }, identifier)
      add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :directory_name }, directory)
    end

  end
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

flags = {}
OptionParser.new do |opts|
  opts.banner = "Usage: prepop_rows.rb [options] FILE"

  opts.on("-a", "--ark", "Mint arks, one per row in output spreadsheet") do |a|
    flags[:ark] = a
  end
end.parse!

spreadsheet_name = 'default.xlsx'
workbook.set_up_spreadsheet

num_rows = Integer(ARGV[0]) rescue false

if num_rows && flags[:ark]
  workbook.blank_rows(num_rows)
else
  csv = ARGV[0]
  options = { :encoding => 'ISO8859-1:utf-8', :key_mapping => CROSSWALKING_TERMS_SINGLE }
  csv_parsed = SmarterCSV.process(csv, options)
  puts 'Writing spreadsheet...'
  workbook.prepop(csv_parsed, flags)
end

workbook.write(spreadsheet_name)
puts "Spreadsheet written to #{spreadsheet_name}."

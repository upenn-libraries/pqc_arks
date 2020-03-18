#!/usr/bin/env ruby

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
  return ark.to_s.gsub(/[:\/]/, ':' => '+', '/' => '=') unless ark.nil?
end

abort('Specify the number of arks to mint') if missing_args?

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
                      :physical_type => 'Material',
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

def workbook.blank_rows(num_rows)
  puts "!!!!!!!!! IMPORTANT !!!!!!!!!!
NO EZID CLIENT ACCOUNT CREDENTIALS SET, USING THE EZID TEST API.  THESE TEMPORARY ARKS WILL EXPIRE IN 14 DAYS.
!!!!!!!!!!!!!!!!!!!!!!!!!!!" if Ezid::Client.config.user == 'apitest'
  worksheet = worksheets[0]
  (0..(num_rows-1)).each do |y_index|
    identifier, directory = mint_arkid
    add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :unique_identifier }, identifier.to_s)
    add_custom_field(y_index+1, QUALIFIED_HEADERS.find_index { |k,_| k == :directory_name }, directory)
  end
end

spreadsheet_name = ARGV[1].nil? ? "#{File.basename(ARGV[0], '.*')}_POPULATED.xlsx" : "#{File.basename(ARGV[1], '.*')}.xlsx"
workbook.set_up_spreadsheet

num_rows = Integer(ARGV[0]) rescue false

workbook.blank_rows(num_rows)

workbook.write(spreadsheet_name)
puts "Spreadsheet written to #{spreadsheet_name}."

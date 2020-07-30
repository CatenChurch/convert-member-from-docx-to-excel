#!/usr/bin/env ruby
# frozen_string_literal: true

require 'logger'
logger = Logger.new(STDOUT)
logger.info { 'Bootstraping application...' }

require_relative '../config/boot'
glob, export_path, *_options = ARGV
export_path ||= './export.csv'
logger.error { 'Require a glob patten to locate target files' } and return if glob.nil?

logger.info { "Using which glob patten to find files: #{glob}" }
file_paths = Pathname.glob(glob)
logger.error 'Files not found' and return if file_paths.empty?

logger.info { "Number of files found: #{file_paths.size}" }
require_relative '../lib/models/member.rb'
error_members = []
members = file_paths.map do |file_path|
  logger.info { "Processing file: #{file_path}" }
  docx = Docx::Document.open file_path.to_s
  table = docx.tables.first
  next if table.nil?

  member = Member.new
  member.file_path = file_path
  begin
    member.load_data_from_docx(docx)
  rescue StandardError => e
    logger.error("Error happen when loading file : #{file_path} : #{e}")
    error_members << member
    next
  end
  member
end

require 'csv'
logger.info { "Writing data to #{export_path}" }
CSV.open(export_path, 'w') do |csv|
  csv << Member.csv_header
  members.each do |member|
    csv << member.to_csv
  end
end

logger.info { 'Mission completed!' }

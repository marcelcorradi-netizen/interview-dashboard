#!/usr/bin/env ruby
# frozen_string_literal: true

# Extracts stakeholders and insights data from Base_Stakeholders_DS_v4.xlsx into JSON files.
# Usage: ruby scripts/extract_data.rb

require 'fileutils'
require 'json'
require 'date'
require 'rexml/document'
require 'shellwords'

ROOT = File.expand_path('..', __dir__)
SOURCE_XLSX = File.join(ROOT, 'Base_Stakeholders_DS_v4.xlsx')
DATA_DIR = File.join(ROOT, 'data')

def read_xml_from_xlsx(path)
  `unzip -p #{Shellwords.escape(SOURCE_XLSX)} #{Shellwords.escape(path)}`
end

shared_strings_doc = REXML::Document.new(read_xml_from_xlsx('xl/sharedStrings.xml'))
shared_strings = []
shared_strings_doc.elements.each('sst/si') do |si|
  text = if (t = si.elements['t'])
           t.text.to_s
         else
           si.elements.collect('r') { |r| r.elements['t']&.text.to_s }.join
         end
  shared_strings << text
end

def column_index(ref)
  letters = ref[/[A-Z]+/]
  letters.chars.reduce(0) { |acc, ch| acc * 26 + (ch.ord - 'A'.ord + 1) } - 1
end

def load_sheet(xml_content, shared_strings)
  doc = REXML::Document.new(xml_content)
  rows = []
  max_cols = 0

  doc.elements.each('worksheet/sheetData/row') do |row|
    cells = []
    row.elements.each('c') do |cell|
      index = column_index(cell.attributes['r'])
      value =
        if cell.attributes['t'] == 's'
          idx = cell.elements['v']&.text
          idx ? shared_strings[idx.to_i] : nil
        else
          cell.elements['v']&.text
        end
      cells[index] = value
    end
    max_cols = [max_cols, cells.length].max
    rows << cells
  end

  rows.each { |row| row.fill(nil, row.length...max_cols) }
  rows
end

sheet1 = load_sheet(read_xml_from_xlsx('xl/worksheets/sheet1.xml'), shared_strings)
sheet2 = load_sheet(read_xml_from_xlsx('xl/worksheets/sheet2.xml'), shared_strings)

stake_headers = sheet1.shift
insight_headers = sheet2.shift

stakeholders = sheet1.reject { |row| row.compact.empty? }.map do |row|
  {
    id: row[stake_headers.index('ID_Stakeholders')]&.to_i,
    nome: row[stake_headers.index('Nome')],
    time: row[stake_headers.index('Time')],
    cargo: row[stake_headers.index('Cargo')],
    area: row[stake_headers.index('Área')],
    tempo_meses: row[stake_headers.index("Tempo de empresa \n(em meses)")]&.to_f,
  }
end

origin = Date.new(1899, 12, 30)
insights = sheet2.reject { |row| row.compact.empty? }.map do |row|
  raw_date = row[insight_headers.index('Data_da_Entrevista')]
  {
    data_entrevista: raw_date ? (origin + raw_date.to_f.round).iso8601 : nil,
    stakeholder_id: row[insight_headers.index('ID_Stakeholder')]&.to_i,
    tipo: row[insight_headers.index('Tipo')],
    descricao: row[insight_headers.index('Descrição')],
    categoria: row[insight_headers.index('Categoria')],
    tags: row[insight_headers.index('Tags')]&.split(/\s*,\s*/) || [],
  }
end

FileUtils.mkdir_p(DATA_DIR)
File.write(File.join(DATA_DIR, 'stakeholders.json'), JSON.pretty_generate(stakeholders))
File.write(File.join(DATA_DIR, 'insights.json'), JSON.pretty_generate(insights))

puts "Stakeholders exportados: #{stakeholders.size}"
puts "Insights exportados: #{insights.size}"

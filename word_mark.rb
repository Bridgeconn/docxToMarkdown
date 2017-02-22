#!/usr/bin/env ruby

require 'docx'
require 'yomu'
require 'libreconv'

file_data = []
name_file = "test"

directory_name = "output_folder"
Dir.mkdir(directory_name) unless File.exists?(directory_name)

output_name = "#{directory_name}/#{File.basename(name_file, '.*')}.tab"
output = File.open(output_name, 'w')

Dir.glob("**/*.docx") do |file_name|
  doc = Docx::Document.open(file_name)

  first_table = doc.tables[0]

  doc.tables.each do |table|
    table.rows.each do |row| # Row-based iteration
      row.cells.each_with_index do |cell, i|
        if i == 2 
          file_data << cell.text
        end
      end
    end
  end
end

hash_file = {}
flag = false

file_data.each_with_index do |l, d|
  if l.include? file_data[d]
    
    flag = true
    if (l[0].to_i != 0)
      
      hash_file[file_data[d]] = file_data[d+1]

      # md_des << file_data - file_data[d]
      # md_description << md_des - file_data[d+1]
      
    end
    if flag
      hash_file[file_data[d]] = file_data[d+2]
      flag = false
    end

  end
end
p hash_file

output.close




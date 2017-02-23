#!/usr/bin/env ruby

require 'docx'

file_data = []
name_file = "test"

directory_name = "output_folder"
Dir.mkdir(directory_name) unless File.exists?(directory_name)

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

file_data.each_with_index do |l, d|
  if l.include? file_data[d]

    if (l[0].to_i != 0)
      md_file_name = file_data[d].split(".")
      md_file_name.each do |r|
        output_name = "#{directory_name}/#{File.basename(r, '.*')}.md"
        output = File.open(output_name, 'w')
        
        md_file_heading = file_data[d+1]
        md_file_description = file_data[d+2]
        output << "#"+"#{md_file_heading}\n\n"
        output << "#{md_file_description} \n"
      end
      
    end
  end
end

output.close




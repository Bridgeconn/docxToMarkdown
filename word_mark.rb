#!/usr/bin/env ruby

require 'docx'

file_data = []
name_file = "test"

# directory_name = "output_folder"
# Dir.mkdir(directory_name) unless File.exists?(directory_name)

t = ""
array_desc = []
heading_hash = {}
temp = ""
output = ""
folder_name = ""
directory_name = ""

flag = true
count = 0
Dir.glob("**/*.docx") do |file_name|
  doc = Docx::Document.open(file_name)

  first_table = doc.tables[0]

  doc.tables.each do |table|
    table.rows.each do |row| # Row-based iteration
      row.cells.each_with_index do |cell, i|
        if i == 2 
          file_data << cell.text.gsub('=','')
        end
      end
    end
  end

  file_data.each_with_index do |l, d|
    if l.include? file_data[d]
      
      if ((l.strip)[0].to_i != 0)
        md_file_name = file_data[d].split(".")
        
        #start folder name
        if flag
          directory_name =  md_file_name[0].to_i
          flag =  false
        end
        count +=1
        
        t = file_data[d+1]

        if(array_desc.size > 0)
          heading_hash[temp] = array_desc
          array_desc = []
        end 
      else
        if(t != l)
          array_desc << l
          temp = t
        end
      end
    end
  end

  if(array_desc.size> 0)
    heading_hash[temp] = array_desc
    array_desc = []
  end

  puts count-1
  puts directory_name

  heading_hash.each_with_index do |(k, v), indx|
    if k != ""
      puts indx
      directory_name += (indx-1)
      Dir.mkdir("#{directory_name}") unless File.exists?("#{directory_name}")
      output_name = "#{directory_name}/#{File.basename(1.to_s, '.*')}.md"
      output = File.open(output_name, 'w')

      output << "#"+"#{k}\n\n"
      v.each do |des|
        output << "#{des} \n"
      end
    end    
  end

end
  





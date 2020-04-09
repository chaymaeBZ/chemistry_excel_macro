#!/usr/bin/ruby

require 'rubyXL'

path_to_excel = ARGV[0]

workbook = RubyXL::Parser.parse(path_to_excel)

worksheet = workbook[0]

# select the 2 relevant columns
@col_a = worksheet.collect {|row| row[0].value.strip.squeeze(" ") if row[0]}.drop(1).compact
@col_b = worksheet.collect {|row| row[1].value.strip.squeeze(" ") if row[1]}.drop(1).compact

# init empty result list
@uniq_values = []

# select elements from first column, not in second one
@col_a.each do |value|
  @uniq_values << value unless @col_b.include? value
end
# select elements from second column, not in first one
@col_b.each do |value|
  @uniq_values << value unless @col_a.include? value
end

worksheet.add_cell(0, 3, "hkl from super structure") 
# write result to column 3
@uniq_values.each_with_index do |val, index|
  worksheet.add_cell(index + 1, 3, val) 
end

# generate output path : base_path/projectname-treated.xls*
target_path = "#{File.dirname(path_to_excel)}/#{File.basename(path_to_excel, ".*")}-treated#{File.extname(path_to_excel)}"

# save
workbook.write(target_path)

# row 12280 colum B
# dis = "日本,朝鲜; 安徽,福建,甘肃,广东,广西,贵州,河北,河南,湖北,湖南,江苏,江西,陕西,山东,山西,四川,云南,浙江"
# dis = "河北,河南,山西,山东,陕西,甘肃,江西,湖南,湖北,四川,云南,福建,广西"
# above are test data set

# read excel file
require 'roo'
require_relative "wrapper.rb"

s = Roo::Excelx.new("/Users/Lisong/Desktop/data/distribution_string_20140913.xlsx")
s.default_sheet = s.sheets.first


first_row = s.first_row
last_row = s.last_row


count = 0
for i in first_row..last_row
  dis = s.cell(i, 'B').to_s
  if !dis.include? ';' and !dis.include? ','
    corrected_dis = dis
  end
  if dis.include? ';' and dis.include? ','
    corrected_dis = sort_province(array_province(dis))+"; "+dis_arbord(dis).to_s
  end
  if !dis.include? ';' and dis.include? ','
    corrected_dis = sort_province(array_province(dis))
  end
  if need_sort(dis)
    count = count + 1
    puts i, dis, corrected_dis
  end
end
puts "We have #{last_row} records!"
puts "#{count} records need sorted!"

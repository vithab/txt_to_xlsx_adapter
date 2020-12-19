require 'write_xlsx'
# Задача: 
  # Получить хэш для записи в xlsx с помощью гема 'write_xlsx'
  # Из текстового файла считываются строки вида:
  # "{:company_title=>"Элитные интерьеры", :branch=>nil, :site=>nil, :address=>nil, :phone=>"", :email=>"", 
  #   :inn=>nil, :description=>"", :affiliated_companies=>"", :persons=>"", :owned_buildings=>"", 
  #   :lease_transactions=>"ТЦ Торговый центр", 
  #   :sale_transactions=>"", :logo=>"", :url=>""}"

lines = File.open('./incoming_files/companies.txt', 'r') { |file| file.readlines }
lines.map! { |link| link.strip }

key_patterns = [
                "company_title=>", "branch=>", "site=>", "address=>", 
                "phone=>", "email=>", "inn=>", "description=>", 
                "affiliated_companies=>", "persons=>", "owned_buildings=>", 
                "lease_transactions=>", "sale_transactions=>", "logo=>", "url=>"
                ]
keys = []
values = []
attributes = []
time = Time.now.to_s.split(' ').first(2).join('_').gsub(':', '-')
file_name = "company_#{time}"
count = lines.size

# Очищаем от лишних символов полученные строки из тхт файла.
# Чтобы получить значения для будущего хэша, вырезаем из строки ключи по key_patterns
lines = lines.map do |line|
  line.gsub!('{', '').gsub!('}', '').gsub!(':', '').gsub!('https', 'https:')  
  key_patterns.each do |pattern|
    line.gsub! pattern, ""
  end
  
  # Делаем проверку если строка содержит 'nil', иначе не возможно вызывать метод у NilClass
  line.gsub!('nil', "\"\"") if line.include?('nil')
  line.split(", \"")
end

# Убираем лишнее, делаем ключи символами, добавляем в массив keys
key_patterns.map { |key| keys << key.gsub('=>', '').to_sym }

# Очищаем мусор в строках
# Объединяем 2 массива в хэш
lines.each_with_index do |line, index|
  line.map { |l| l.gsub!("\"", "") }
  puts "Writing string:   #{index + 1} from #{count}"
  attributes << Hash[keys.zip(line)]
end

def print_to_xlsx(file_name, attributes, workbook, worksheet)
  format_header = workbook.add_format
  format_header.set_bold
  format_header.set_bg_color('yellow')
  format_header.set_align('center')
  format_header.set_align('vcenter')
  format_url = workbook.add_format(:color => 'blue', :underline => 1)

  worksheet.write('A1', 'Компания', format_header)
  worksheet.write('B1', 'Отрасль', format_header)
  worksheet.write('C1', 'Сайт', format_header)
  worksheet.write('D1', 'Адрес', format_header)
  worksheet.write('E1', 'Телефон', format_header)
  worksheet.write('F1', 'Email', format_header)
  worksheet.write('G1', 'ИНН', format_header)
  worksheet.write('H1', 'Описание', format_header)
  worksheet.write('I1', 'Дочернии компании', format_header)
  worksheet.write('J1', 'Персоны', format_header)
  worksheet.write('K1', 'Здания в собственности', format_header)
  worksheet.write('L1', 'Сделки по аренде', format_header)
  worksheet.write('M1', 'Сделки по продаже', format_header)
  worksheet.write('N1', 'Лого', format_header)
  worksheet.write('O1', 'Ссылка на компанию', format_header)

  i = 2

  attributes.each do |r|
    worksheet.write_string("A#{i}", "#{r[:company_title]}")
    worksheet.write_string("B#{i}", "#{r[:branch]}")
    worksheet.write_string("C#{i}", "#{r[:site]}")
    worksheet.write_string("D#{i}", "#{r[:address]}")
    worksheet.write_string("E#{i}", "#{r[:phone]}")
    worksheet.write_string("F#{i}", "#{r[:email]}")
    worksheet.write_string("G#{i}", "#{r[:inn]}")
    worksheet.write_string("H#{i}", "#{r[:description]}")
    worksheet.write_string("I#{i}", "#{r[:affiliated_companies]}")
    worksheet.write_string("J#{i}", "#{r[:persons]}")
    worksheet.write_string("K#{i}", "#{r[:owned_buildings]}")
    worksheet.write_string("L#{i}", "#{r[:lease_transactions]}")
    worksheet.write_string("M#{i}", "#{r[:sale_transactions]}")
    worksheet.write_url(   "N#{i}", "#{r[:logo]}", format_url)
    worksheet.write_url(   "O#{i}", "#{r[:url]}", format_url)

    i += 1
  end

  workbook.close
end

puts "\nWait... I`m working."

workbook = WriteXLSX.new("./results/#{file_name}.xlsx")
worksheet = workbook.add_worksheet
print_to_xlsx(file_name, attributes, workbook, worksheet)

puts "\n\nDONE! Check file: /results/#{file_name}.xlsx\n\n"

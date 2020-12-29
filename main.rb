require 'write_xlsx'
require 'byebug'
require_relative 'lib/print_results'
# Задача: 
  # Получить хэш для записи в xlsx с помощью гема 'write_xlsx'
  # Из текстового файла считываются строки вида:
  # "{:company_title=>\"Прикосновение\", :branch=>nil, :site=>nil, :address=>nil, 
  #   :phone=>\"\", :email=>\"\", :inn=>nil, :description=>\"\", :affiliated_companies=>\"\", 
  #   :persons=>\"\", :owned_buildings=>\"\", 
  #   :lease_transactions=>\"МЦ Бизнес-центр улица Миклухо-Маклая, 36А\", 
  #   :sale_transactions=>\"\", :logo=>\"\", :url=>\"https://any-site.com/any-category/prikosnoveniye\"}"

lines = File.open('./incoming_files/companies.txt', 'r') { |file| file.readlines }
lines.map! { |link| link.strip }

# Важно: если есть общая составная часть в ключах(key_patterns), чтоб в очереди массива
# она была последней(идущей за наименьшим вхождением), например:
# "company_title=>" и "title=>" - должен идти после "company_title=>",
# иначе при итерации вырежется "title=>" из "company_title=>" (останется "company_") 
# и хэш неправильно соберётся
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
def clear_line(line, key_patterns)
  line.gsub!('{', '').gsub!('}', '').gsub!(':', '').gsub!('https', 'https:')  
  key_patterns.each do |pattern|
    line.gsub! pattern, ""
  end
  
  # Делаем проверку если строка содержит 'nil', иначе не возможно вызывать метод у NilClass
  line.gsub!('nil', "\"\"") if line.include?('nil')
  line.split("\", \"")
end

lines = lines.map do |line|
  clear_line(line, key_patterns)
end

# Убираем лишнее, делаем ключи символами, добавляем в массив keys
key_patterns.map { |key| keys << key.gsub('=>', '').to_sym }

# Очищаем мусор в строках
# Объединяем 2 массива в хэш
lines.each_with_index do |line, index|
  line.map { |l| l.gsub!("\"", "") }
  puts "Reading string:   #{index + 1} from #{count}"
  attributes << Hash[keys.zip(line)]
end

puts "\nWait...writing. I`m working."

# Создаём книгу, лист и передаем в метод
workbook = WriteXLSX.new("./results/#{file_name}.xlsx")
worksheet = workbook.add_worksheet

# Учесть тот факт, что gem 'write_xlsx-0.85.7' даёт на 1 лист записать не более 65530 ссылок.
# number of URLS is over Excel's limit of 65,530 URLS per worksheet. (RuntimeError)
PrintResults.print_to_xlsx(file_name, attributes, workbook, worksheet)

puts "\n\nDONE! Check file: /results/#{file_name}.xlsx\n\n"

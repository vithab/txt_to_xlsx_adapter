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

require 'write_xlsx'

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

# Очищаем от лишних символов полученные строки из тхт файла.
# Чтобы получить значения для будущего хэша, вырезаю из строки ключи по key_patterns
lines = lines.map do |line|
  line.gsub!('{', '').gsub!('}', '').gsub!(':', '').gsub!('https', 'https:')
  
  key_patterns.each do |pattern|
    line.gsub! pattern, ""
  end
  
  line.gsub!('nil', "\"\"").split(", \"")
end

p lines[0]

require 'write_xlsx'

lines = File.open('./incoming_files/companies.txt', 'r') { |file| file.readlines }
lines.map! { |link| link.strip }

p lines

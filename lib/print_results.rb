module PrintResults
  
  # Метод для сохранения в файл xlsx (аргументы: имя файла, аттрибуты объекта(поля), книга и лист Эксель)
  def self.print_to_xlsx(file_name, attributes, workbook, worksheet)
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
end

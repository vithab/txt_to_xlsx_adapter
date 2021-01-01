#### Задача: 
Получить массив хэшей из массива строк, для записи в xlsx с помощью гема 'write_xlsx'.

Из текстового файла считываются строки вида:

[
    "{:company_title=>\"Прикосновение\", :branch=>nil, :site=>nil, :address=>nil, 
    :phone=>\"\", :email=>\"\", :inn=>nil, :description=>\"\", :affiliated_companies=>\"\", 
    :persons=>\"\", :owned_buildings=>\"\", 
    :lease_transactions=>\"МЦ Бизнес-центр улица Миклухо-Маклая, 36А\", 
    :sale_transactions=>\"\", :logo=>\"\", :url=>\"https://any-site.com/any-category/prikosnoveniye\"}"
]

Получить нужно:

[
    {:company_title=>\"Прикосновение\", :branch=>nil, :site=>nil, :address=>nil, 
    :phone=>\"\", :email=>\"\", :inn=>nil, :description=>\"\", :affiliated_companies=>\"\", 
    :persons=>\"\", :owned_buildings=>\"\", 
    :lease_transactions=>\"МЦ Бизнес-центр улица Миклухо-Маклая, 36А\", 
    :sale_transactions=>\"\", :logo=>\"\", :url=>\"https://any-site.com/any-category/prikosnoveniye\"}
]
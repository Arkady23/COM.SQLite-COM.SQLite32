# COM.SQLite — sqlite.net.dll
### Оглавление
[Назначение](#Назначение)  
[Регистрация COM.SQLite в реестре Windows](#Регистрация-COMSQLite-в-реестре-Windows)  
[Создание объекта COM.SQLite](#Создание-объекта-COMSQLite)  
[Метод Open(sql)](#Метод-Opensql)  
[Метод DoCmd(sql)](#Метод-DoCmdsql)  
[Группа методов с параметрами DoCmd1(sql, v1) —<br>DoCmd9(sql, v1, v2, v3, v4, v5, v6, v7, v8, v9)](#Группа-методов-с-параметрами-DoCmd1sql-v1-DoCmd9sql-v1-v2-v3-v4-v5-v6-v7-v8-v9)  
### Назначение
Библиотека sqlite.net.dll реализует COM-сервер для VFP9 или VFPA, который в принципе может использоваться и в любых других языках, поддерживающих COM технологию обмена данными.  

Microsoft VFP имеет ограничение на размер файла БД и на количество записей в нем. Отчасти проблему размера файла dbf решает VFPA,
но проблема максимального числа записей в файле приблизительно до 1 миллиарда записей — не решима. К тому же чтобы получить на
сегодняшний день VFPA нужно оформлять платную подписку. Многие разработчики уже перешли на другие СУБД и в том числе на SQLite. Данный COM позволяет использовать SQLite на VFP9, VFPA и других языках Windows, например на JScript платформы WSH. COM.SQLite в том числе может использоваться C# без установки каких-либо дополнительных пакетов.  

При создании COM-объекта COM.SQLite, файл sqlite3.dll с API от разработчика SQLite Consortium должен находиться рядом с файлом sqlite32.net.dll или sqlite.net.dll в зависимости от разрядности sqlite3.dll и разрядности программы, которую вы разрабатываете.  

В COM.SQLite реализованы методы:
- Open(),
- DoCmd(),
- DoCmdN(),
- Next(),
- Eof() и
- Close().

Дополнительно для удобства разработчиков вместо параметрического метода DoCmdN() можно использовать методы с номерами от DoCmd1() до DoCmd9() для включения в команду SQL соответствующего числа параметров.
### Регистрация COM.SQLite в реестре Windows
#### Для VFPA и другого 64-х разрядного ПО
Чтобы объект COM.SQLite был доступен в разрабатываемых программах 64-х битных версий, его нужно зарегистрировать в ОС с помощью
утилиты regasm.exe. Например:
```
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\sqlite.net.dll /codebase
```
Предварительно поместите файл sqlite.net.dll в удобную для вас папку, например, в папку, где находятся другие библиотеки
Microsoft VFP.  

Для удаления регистрации из Windows используйте ключ /unregister. Например:
```
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\sqlite.net.dll /unregister
```
Для выполнения вышеуказанных команд требуются права администратора.
#### Для VFP9 и другого 32-х разрядного ПО
Используйте утилиту регистрации для 32-х разрядных программ, находящуюся по другому пути:
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe. Команды на регистрацию и удаление регистрации аналогичны командам
для 64-х разрядного ПО. Например, регистрация:
```
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\sqlite32.net.dll /codebase
```
и удаление регистрации:
```
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\sqlite32.net.dll /unregister
```
## Создание объекта COM.SQLite
Текст кода на VFP:
```
SQLite = CreateObject('COM.SQLite')
```
## Метод Open(sql)
Вызов метода обеспечивает открытие БД SQLite с умалчиваемыми характеристиками, т.е. с возможностью запись/чтение и создание БД, если она не существует. Строковым параметром указыватся название файла БД вместе с путем. Если путь не указан, то предполагается наличие БД в папке по умолчанию. Если БД открылась, метод возвращает 0, иначе — код ошибки.
## Метод DoCmd(sql)
Метод компилирует запрос заданный текстовой строкой sql. Если запрос верный и его можно выполнять, метод возвращает 0, иначе — код ошибки.
## Группа методов с параметрами DoCmd1(sql, v1) —<br>DoCmd9(sql, v1, v2, v3, v4, v5, v6, v7, v8, v9)
Любой из этих методов компилирует запрос заданный текстовой строкой sql с учетом значений, заданных параметрами. Количество параметров соответствует указанному в названии метода числу. Параметры могут быть произвольных типов, но предусмотренных в SQLite, т.е. blob, text, integer, double и null. Следует отметить, что несмотря на то, что текстовые значения передаются в БД как есть, SQLite предполагает, что все текстовые строки находятся в кодировке UTF-8. Если запрос верный и его можно выполнять, метод возвращает 0, иначе — код ошибки.
## Метод DoCmdN(sql, @vals)
Метод компилирует запрос заданный текстовой строкой sql с учетом значений, заданных массивом параметров. Массив передается по ссылке. Нумерация элементов массива с параметрами начинается от нуля. Смотрите ниже приведенный пример на языке Visual FoxPro. Если запрос верный и его можно выполнять, метод возвращает 0, иначе — код ошибки.
## Метод Next()
Метод обеспечивает выполнение скомпелированного запроса и возвращает резульлтат построчно. Формирование строк выполняется асинхронно. Пока вы обрабатываете результат полученой строки, в это время формируется следующая строка. Метод возвращает массив значений полей строки или null, если строк больше нет.
## Метод Eof()
Метод возвращает 0, если сформирована подготовленная SQL-запросом строка. Если строк больше нет, метод возвращает -1.
## Метод Close()
Метод возвращает 0, если БД успешно закрыта, иначе — код ошибки.
## Примеры
Пример на VFP. Создание таблицы и внесение данных:
```
? SQLite=CreateO('COM.SQLite')
? SQLite.Open('test.db')
? SQLite.DoCmd("DROP TABLE IF EXISTS people; CREATE TABLE people(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, age INTEGER)")
? SQLite.DoCmd2("INSERT INTO people (name, age) VALUES (?, ?);",strconv("Аркадий",9),64)
```

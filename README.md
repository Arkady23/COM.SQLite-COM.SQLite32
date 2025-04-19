# COM.SQLite и COM.SQLite32
### Оглавление
[Назначение](#Назначение)  
[Регистрация COM.SQLite и COM.SQLite32 в реестре Windows](#Регистрация-COMSQLite-и-COMSQLite32-в-реестре-Windows)  
[Создание объекта COM.SQLite и COM.SQLite32](#Создание-объекта-COMSQLite-и-COMSQLite32)  
[Метод Open(file, readonly)](#Метод-Openfile-readonly)  
[Метод DoCmd(sql[, v1[, v2[, v3[, v4[, v5[, v6[, v7[, v8[, v9[, v10]]]]]]]]]])](#Метод-DoCmdsql-v1-v2-v3-v4-v5-v6-v7-v8-v9-v10)  
[Метод DoCmdN(sql, @vals)](#Метод-DoCmdNsql-vals)  
[Метод Next()](#Метод-Next)  
[Метод Eof()](#Метод-Eof)  
[Метод Close()](#Метод-Close)  
[Примеры на Visual FoxPro](#Примеры-на-Visual-FoxPro)  
[Обсуждение](#Обсуждение)  

[История версий](#История-версий)  
### Назначение
Библиотека sqlite.net.dll реализует COM-сервер для VFP9 или VFPA, который в принципе может использоваться и в любых других языках, поддерживающих COM технологию обмена данными.  

Microsoft VFP имеет ограничение на размер файла БД и на количество записей в нем. Отчасти проблему размера файла dbf решает VFPA,
но проблема максимального числа записей в файле приблизительно до 1 миллиарда записей — не решима. К тому же чтобы получить на
сегодняшний день VFPA нужно оформлять платную подписку. Многие разработчики уже перешли на другие СУБД и в том числе на SQLite. Данный COM позволяет использовать SQLite на VFP9, VFPA и других языках Windows, например на JScript платформы WSH. COM.SQLite и COM.SQLite32 в том числе могут использоваться на C# без установки каких-либо дополнительных пакетов.  

При создании COM-объекта COM.SQLite или COM.SQLite32, файл sqlite3.dll с API от разработчика SQLite Consortium должен находиться рядом с файлом sqlite.net.dll или sqlite32.net.dll в зависимости от разрядности sqlite3.dll и разрядности программы, которую вы разрабатываете.  

В COM.SQLite/COM.SQLite32 реализованы методы:
- Open(),
- DoCmd(),
- DoCmdN(),
- Next(),
- Eof() и
- Close().
### Регистрация COM.SQLite и COM.SQLite32 в реестре Windows
#### Для VFPA и другого 64-х разрядного ПО
Чтобы объект COM.SQLite был доступен в разрабатываемых программах 64-х битных версий, его нужно зарегистрировать в ОС с помощью
утилиты regasm.exe с ключом /codebase. Например:
```PowerShell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\sqlite.net.dll /codebase
```
Предварительно поместите файл sqlite.net.dll в удобную для вас папку, например, в папку, где находятся другие библиотеки
Microsoft VFP.  

Для удаления регистрации из Windows используйте ключ /unregister. Например:
```PowerShell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\sqlite.net.dll /unregister
```
Для выполнения вышеуказанных команд требуются права администратора.
#### Для VFP9 и другого 32-х разрядного ПО
Используйте утилиту регистрации для 32-х разрядных программ, находящуюся по другому пути:
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe. Команды на регистрацию и удаление регистрации аналогичны командам
для 64-х разрядного ПО. Например, регистрация:
```PowerShell
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\sqlite32.net.dll /codebase
```
и удаление регистрации:
```PowerShell
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\sqlite32.net.dll /unregister
```
## Создание объекта COM.SQLite и COM.SQLite32
Текст кода на VFP:
```xBase
SQLite = CreateObject('COM.SQLite')
```
и
```xBase
SQLite = CreateObject('COM.SQLite32')
```
## Метод Open(file[, readonly])
Вызов метода обеспечивает открытие БД SQLite с возможностью запись/чтение и создание БД, если она не существует, если логический параметр readonly ложь либо он не указан. Если значение второго параметра правда, то БД открывается только для чтения. Первым строковым параметром указыватся название файла БД вместе с путем. Если путь не указан, то предполагается наличие БД в папке местонахождения библиотек sqlite.dll и sqlite.net.dll. Допускается указание имени файла в фомате URI, т.е. имя может начинаться с префикса "file://". Если БД открылась, метод возвращает 0, иначе — код ошибки.
## Метод DoCmd(sql[, v1[, v2[, v3[, v4[, v5[, v6[, v7[, v8[, v9[, v10]]]]]]]]]])
Метод выполняет запрос, заданный текстовой строкой sql и, если неоходимо — дополнительными параметрами в количестве до 10-ти значений. Если запрос выполнен, метод возвращает 0, иначе — код ошибки.
## Метод DoCmdN(sql, @vals)
Метод выполняет запрос, заданный текстовой строкой sql с учетом значений, заданных массивом параметров. Массив передается по ссылке. Нумерация элементов массива с параметрами начинается от нуля. Смотрите ниже приведенный пример на языке Visual FoxPro. Если запрос выполнен, метод возвращает 0, иначе — код ошибки.
## Метод Next()
Метод возвращает массив значений полей текущей строки результата выполнения sql-команды. Если строк больше нет, метод возвращает значение null.
## Метод Eof()
Метод возвращает 0, если текущая строка с результами запроса еще не прочитана. Если строк больше нет, метод возвращает -1.
## Метод Close()
Метод возвращает 0, если БД успешно закрыта, иначе — код ошибки.
## Примеры на Visual FoxPro
В примере приведено создание таблицы, разные варианты добавления записей, отображение записей и резервирование с уплотнением БД.
```xbase
* ТЕСТЫ РАБОТЫ С ОБЪЕКТОМ COM.SQLite
* TESTS OF WORKING WITH THE OBJECT COM.SQLite
  ? "Test 1: " + tran(Test1())
  ? "Test 2: " + tran(Test2())
  ? "Test 3: " + tran(Test3())
  ? "Test 4: " + tran(Test4())
  ? "Test 5: " + tran(Test5())

* СОЗДАНИЕ ТАБЛИЦЫ И ДОБАВЛЕНИЕ ЗАПИСЕЙ
* CREATING A TABLE AND ADDING RECORDS
FUNCTION Test1
  local ret
  SQLite = CreateO('COM.SQLite')

  ret = SQLite.Open('test.db')
     if ret<>0
        return 1
     endif
  ret = SQLite.DoCmd("DROP TABLE IF EXISTS people;"+ ;
        "CREATE TABLE people(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, age INTEGER)")
     if ret<>0
        return 2
     endif

  * В запросе требуются два параметра со значениями:
  * The request requires two parameters with values:
  ret = SQLite.DoCmd("INSERT INTO people (name, age) VALUES (?, ?)", "Bill Gates", 69)
     if ret<>0
        return 3
     endif

  * Другие записи:
  * Other records:
  ret = SQLite.DoCmd("INSERT INTO people (name, age) VALUES (?, ?)", "Richard Hipp", 64)
     if ret<>0
        return 4
     endif
  ret = SQLite.DoCmd("INSERT INTO people (name, age) VALUES (?, ?)", ;
        Strconv("Аркадий Корниенко",9), 64)
     if ret<>0
        return 5
     endif
  ret = SQLite.Close()
     if ret<>0
        return 6
     endif

RETURN 0

* ТОЖЕ САМОЕ, НО С ИСПОЛЬЗОВАНИЕМ МАССИВА ПАРАМЕТРОВ
* THE SAME, BUT USING AN ARRAY OF PARAMETERS
FUNCTION Test2
  local ret
  SQLite = CreateO('COM.SQLite')

  * Указываем тип массива параметров с нулевого элемента для COM-объекта SQLite:
  * Specify the parameter array type with a zero element for the SQLite COM object:
  ComArray(SQLite,10)

  * Формируем массив параметров размером в 6 элементов:
  * We form an array of parameters with a size of 6 elements:
  dime vals(6)
  vals(1)="Bill Gates"
  vals(2)=69
  vals(3)="Richard Hipp"
  vals(4)=64
  vals(5)=Strconv("Аркадий Корниенко",9)
  vals(6)=64

  ret = SQLite.Open('test.db')
     if ret<>0
        return 7
     endif
  ret = SQLite.DoCmdN("DROP TABLE IF EXISTS people;"+ ;
        "CREATE TABLE people(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, age INTEGER);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);", ;
        @vals)
     if ret<>0
        return 8
     endif
  ret = SQLite.Close()
     if ret<>0
        return 9
     endif

RETURN 0

* ТОЖЕ САМОЕ НО С ИСПОЛЬЗОВАНИЕМ ТРАНЗАКЦИИ
* THE SAME BUT USING A TRANSACTION
FUNCTION Test3
  local ret
  SQLite = CreateO('COM.SQLite')

  ret = SQLite.Open('test.db')
     if ret<>0
        return 10
     endif

  * Использование транзакции при выполнении нескольких SQL-комманд:
  * Using a Transaction When Executing Multiple SQL Statements:
  ret = SQLite.DoCmd("BEGIN TRANSACTION;")
     if ret<>0
        return 11
     endif

  ret = SQLite.DoCmd("DROP TABLE IF EXISTS people;"+ ;
        "CREATE TABLE people(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, age INTEGER);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);", ;
        "Bill Gates", 69, "Richard Hipp", 64, ;
        Strconv("Аркадий Корниенко",9), 64)
     if ret<>0
        ret = SQLite.DoCmd("ROLLBACK;")
        if ret<>0
           return 12
        endif
        return 13
     endif

  ret = SQLite.DoCmd("COMMIT;")
     if ret<>0
        return 14
     endif

  ret = SQLite.Close()
     if ret<>0
        return 15
     endif

RETURN 0

* ВЫВОД ЗАПИСЕЙ
* OUTPUT OF RECORDS
FUNCTION Test4
  local ret
  SQLite = CreateO('COM.SQLite')

  ret = SQLite.Open('test.db')
     if ret<>0
        return 16
     endif

  ret = SQLite.DoCmd("SELECT * FROM people")
     if ret<>0
        return 17
     endif
     do while SQLite.Eof()=0
        arec = SQLite.Next()
        ? tran(arec(1))+" | "+Strconv(arec(2),11)+" | "+tran(arec(3))
     enddo
     ?

  ret = SQLite.Close()
     if ret<>0
        return 18
     endif

RETURN 0

* СЖАТИЕ И КОПИРОВАНИЕ БД
* COMPRESSION AND COPYING OF THE DB
FUNCTION Test5
  local ret
  SQLite = CreateO('COM.SQLite')
  bak = 'sqlite.test.db.bak'
  backup = 'sqlite.test.db'

  ret = SQLite.Open('test.db')
     if ret<>0
        return 19
     endif

  if(file(m.bak))
     dele file (m.bak)
  endif
  if(file(m.backup))
     rena (m.backup) to (m.bak)
  endif

  ret = SQLite.DoCmd("VACUUM INTO ?",m.backup)
     if ret<>0
        return 20
     endif

  ret = SQLite.Close()
     if ret<>0
        return 21
     endif

RETURN 0
```
### Обсуждение
Задать вопрос или обсудить тему, касающуюся SQLite или VFP/VFPA, вы можете в разделе проекта `Issues` > `New issue`.
### История версий
0.0.0.0. 18.04.2025. Опубликована первая рабочая версия.  

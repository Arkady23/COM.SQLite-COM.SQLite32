# COM.SQLite и COM.SQLite32
### Оглавление
[Назначение](#Назначение)  
[Регистрация COM.SQLite32 и COM.SQLite в реестре Windows](#Регистрация-COMSQLite32-и-COMSQLite-в-реестре-Windows)  
[Создание объекта COM.SQLite32 и COM.SQLite](#Создание-объекта-COMSQLite32-и-COMSQLite)  
[Метод Open(file[, readonly])](#Метод-Openfile-readonly)  
[Метод DoCmd(sql[, v1[, v2[, v3[, v4[, v5[, v6[, v7[, v8[, v9[, v10]]]]]]]]]])](#Метод-DoCmdsql-v1-v2-v3-v4-v5-v6-v7-v8-v9-v10)  
[Метод DoCmdN(sql, @vals)](#Метод-DoCmdNsql-vals)  
[Метод Next()](#Метод-Next)  
[Метод Eof()](#Метод-Eof)  
[Метод Close()](#Метод-Close)  
[Примеры на Visual FoxPro](#Примеры-на-Visual-FoxPro)  
[Обсуждение](#Обсуждение)  

[История версий](#История-версий)  
### Назначение
Библиотеки sqlite32.net.dll и sqlite.net.dll реализуют COM-сервер для VFP9 или VFPA, который в принципе может использоваться и в любых других языках, поддерживающих COM технологию обмена данными.  

Microsoft VFP имеет ограничение на размер файла БД и на количество записей в нем. Отчасти проблему размера файла dbf решает VFPA, но проблема максимального числа записей в файле приблизительно до 1 миллиарда записей — не решима. К тому же чтобы получить на сегодняшний день VFPA нужно оформлять платную подписку. Многие разработчики уже перешли на другие СУБД и в том числе на SQLite. Данный COM позволяет использовать SQLite на VFP9, VFPA и других языках Windows. COM.SQLite32 и COM.SQLite в том числе могут использоваться на C# без установки каких-либо дополнительных пакетов.  

В начальной нулевой версии при создании COM-объекта COM.SQLite32 или COM.SQLite, файл sqlite3.dll с API от разработчика SQLite Consortium должен находиться рядом с файлом sqlite.net.dll или sqlite32.net.dll в зависимости от разрядности sqlite3.dll и разрядности программы, которую вы разрабатываете. Нчиная с версии библиотек 0.1.0.0 COM.SQLite32/COM.SQLite не требует наличия файла sqlite3.dll, т.к. на самом деле по умолчанию в Windows10/11 эта dll уже находится в папке C:/Windows/System32 или в папке C:/Windows/SysWOW64, причем последняя актуальная версия. Название файла — winsqlite3.dll.  

В COM.SQLite32/COM.SQLite реализованы методы:
- Open(),
- DoCmd(),
- DoCmdN(),
- Next(),
- Eof() и
- Close().
### Регистрация COM.SQLite32 и COM.SQLite в реестре Windows
#### Для VFP9 и другого 32-х разрядного ПО
Используйте утилиту регистрации regasm.exe **для 32-х разрядных программ**, находящуюся по пути C:\Windows\Microsoft.NET\Framework\v4.0.30319, с ключом /codebase. Пример регистрации:
```PowerShell
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\sqlite32.net.dll /codebase
```
Предварительно поместите файл sqlite32.net.dll в удобную для вас папку, например, в папку, где находятся другие библиотеки Microsoft VFP.  

Для удаления регистрации из Windows используйте ключ /unregister:
```PowerShell
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe D:\VFP\VFP9\sqlite32.net.dll /unregister
```
Для выполнения вышеуказанных команд требуются права администратора.
#### Для VFPA и другого 64-х разрядного ПО
Чтобы объект COM.SQLite был доступен в разрабатываемых программах 64-х битных версий, его нужно зарегистрировать в ОС с помощью
утилиты regasm.exe **для 64-х разрядных программ**, находящуюся в **другой папке** (**это важно**) C:\Windows\Microsoft.NET\Framework64\v4.0.30319, с ключом /codebase. Пример регистрации:
```PowerShell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\sqlite.net.dll /codebase
```
Предварительно поместите файл sqlite.net.dll в удобную для вас папку, например, в папку, где находятся другие библиотеки
Microsoft VFP.  

Для удаления регистрации из Windows используйте ключ /unregister:
```PowerShell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe D:\VFP\VFPA\sqlite.net.dll /unregister
```
Для выполнения вышеуказанных команд требуются права администратора.
## Создание объекта COM.SQLite32 и COM.SQLite
Текст кода на VFP:
```xBase
SQLite = CreateObject('COM.SQLite32')
```
и
```xBase
SQLite = CreateObject('COM.SQLite')
```
или универсальный вариант:
```xBase
SQLite = CreateObject('COM.SQLite'+iif(sys(17)='Pentium','32',''))
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
  ? "Test 6: " + tran(Test6())
  ? "Test 7: " + tran(Test7())

* СОЗДАНИЕ ТАБЛИЦЫ И ДОБАВЛЕНИЕ ЗАПИСЕЙ
* CREATING A TABLE AND ADDING RECORDS
FUNCTION Test1
  SQLite = CreateO('COM.SQLite'+iif(sys(17)='Pentium','32',''))

  if SQLite.Open('test.db')<>0
     return 1
  endif
  if SQLite.DoCmd("DROP TABLE IF EXISTS people;"+        ;
           "CREATE TABLE people(id INTEGER PRIMARY KEY"+ ;
           " AUTOINCREMENT, name TEXT, age INTEGER)")<>0
     return 2
  endif

  * В запросе требуются два параметра со значениями:
  * The request requires two parameters with values:
  if SQLite.DoCmd("INSERT INTO people (name, age) VALUES (?, ?)", ;
           "Bill Gates", 69)<>0
     return 3
  endif

  * Другие записи:
  * Other records:
  if SQLite.DoCmd("INSERT INTO people (name, age) VALUES (?, ?)", ;
           "Richard Hipp", 64)<>0
     return 4
  endif
  if SQLite.DoCmd("INSERT INTO people (name, age) VALUES (?, ?)", ;
          Strconv("Аркадий Корниенко",9), 64)<>0
     return 5
  endif
  if SQLite.Close()<>0
     return 6
  endif

RETURN 0

* ТОЖЕ САМОЕ, НО С ИСПОЛЬЗОВАНИЕМ МАССИВА ПАРАМЕТРОВ
* THE SAME, BUT USING AN ARRAY OF PARAMETERS
FUNCTION Test2
  SQLite = CreateO('COM.SQLite'+iif(sys(17)='Pentium','32',''))

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

  if SQLite.Open('test.db')<>0
     return 7
  endif
  if SQLite.DoCmdN("DROP TABLE IF EXISTS people;"+       ;
        "CREATE TABLE people(id INTEGER PRIMARY KEY"+    ;
        " AUTOINCREMENT, name TEXT, age INTEGER);"+      ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);", ;
        @vals)<>0
     return 8
  endif
  if SQLite.Close()<>0
     return 9
  endif

RETURN 0

* ТОЖЕ САМОЕ НО С ИСПОЛЬЗОВАНИЕМ ТРАНЗАКЦИИ
* THE SAME BUT USING A TRANSACTION
FUNCTION Test3
  SQLite = CreateO('COM.SQLite'+iif(sys(17)='Pentium','32',''))

  if SQLite.Open('test.db')<>0
     return 10
  endif

  * Использование транзакции при выполнении нескольких SQL-комманд:
  * Using a Transaction When Executing Multiple SQL Statements:
  if SQLite.DoCmd("BEGIN TRANSACTION;")<>0
     return 11
  endif

  if SQLite.DoCmd("DROP TABLE IF EXISTS people;"+        ;
        "CREATE TABLE people(id INTEGER PRIMARY KEY"+    ;
        " AUTOINCREMENT, name TEXT, age INTEGER);"+      ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);"+ ;
        "INSERT INTO people (name, age) VALUES (?, ?);", ;
        "Bill Gates", 69, "Richard Hipp", 64, ;
        Strconv("Аркадий Корниенко",9), 64)<>0
     if SQLite.DoCmd("ROLLBACK;")<>0
        return 12
     endif
     return 13
  endif

  if SQLite.DoCmd("COMMIT;")<>0
     return 14
  endif

  if SQLite.Close()<>0
     return 15
  endif

RETURN 0

* ВЫВОД ЗАПИСЕЙ
* OUTPUT OF RECORDS
FUNCTION Test4
  SQLite = CreateO('COM.SQLite'+iif(sys(17)='Pentium','32',''))

  if SQLite.Open('test.db')<>0
     return 16
  endif

  if SQLite.DoCmd("SELECT * FROM people")<>0
     return 17
  endif
  do while SQLite.Eof()=0
     arec = SQLite.Next()
     ? tran(arec(1))+" | "+Strconv(arec(2),11)+" | "+tran(arec(3))
  enddo
  ?

  if SQLite.Close()<>0
     return 18
  endif

RETURN 0

* ПРОВЕРКА РАЗНЫХ ТИПОВ ДАННЫХ (DOUBLE/REAL, BOOL, NULL, INT64, DATETIME)
* TESTING DIFFERENT DATA TYPES
FUNCTION Test5
  LOCAL SQLite, nVal, lVal, dVal, nBig
  SQLite = CreateO('COM.SQLite'+iif(sys(17)='Pentium','32',''))

  if SQLite.Open('test.db')<>0
     return 19
  endif

  * Создаём таблицу с разными типами колонок
  if SQLite.DoCmd("DROP TABLE IF EXISTS types_test;" + ;
           "CREATE TABLE types_test (" + ;
           "id INTEGER PRIMARY KEY AUTOINCREMENT, " + ;
           "f_double DOUBLE, " + ;
           "f_real REAL, " + ;
           "f_bool BOOLEAN, " + ;
           "f_big INTEGER, " + ;
           "f_date TEXT, " + ;
           "f_null TEXT)")<>0
     return 20
  endif

  * Тестовые значения из VFP
  nVal = 12345.67890123   && Дробное число высокой точности
  lVal = .T.              && Логическое значение
  nBig = 9876543210123    && Большое целое (больше 32-битного INT)
  dVal = DATETIME(2026, 11, 13, 13, 46, 00)

  * Записываем данные через параметры
  if SQLite.DoCmd("INSERT INTO types_test (f_double, f_real, f_bool, f_big, f_date, f_null) " + ;
           "VALUES (?, ?, ?, ?, ?, ?)", ;
           m.nVal, -0.000456, m.lVal, m.nBig, TTOC(m.dVal, 3), .NULL.)<>0
     return 21
  endif

  * Читаем обратно и проверяем типы и значения
  if SQLite.DoCmd("SELECT f_double, f_real, f_bool, f_big, f_date, f_null FROM types_test")<>0
     return 22
  endif

  if SQLite.Eof()=0
     arec = SQLite.Next()
     ? "--- Проверка типов ---"
     ? "DOUBLE:   ", arec(1), "| Тип VFP:", VARTYPE(arec(1)), "| Разница:", ABS(arec(1) - m.nVal)
     ? "REAL:     ", arec(2), "| Тип VFP:", VARTYPE(arec(2))
     ? "BOOL:     ", arec(3), "| Тип VFP:", VARTYPE(arec(3))
     ? "INT64:    ", arec(4), "| Тип VFP:", VARTYPE(arec(4))
     ? "DATETIME: ", arec(5), "| Тип VFP:", VARTYPE(arec(5))
     ? "NULL:     ", arec(6), "| Тип VFP:", VARTYPE(arec(6)), "| ISNULL:", ISNULL(arec(6))

     * Проверяем, что Double вернулся именно как число (N) и без потери точности
     if VARTYPE(arec(1)) <> 'N' OR ABS(arec(1) - m.nVal) > 0.0000001
        SQLite.Close()
        return 23
     endif
  else
     SQLite.Close()
     return 24
  endif
  ?

  if SQLite.Close()<>0
     return 25
  endif

RETURN 0

* ПРОВЕРКА ЗАПИСИ И ЧТЕНИЯ BLOB
* BLOB WRITE AND READ VERIFICATION
FUNCTION Test6
  LOCAL SQLite, cBinIn, cBinOut
  SQLite = CreateO('COM.SQLite'+iif(sys(17)='Pentium','32',''))

  if SQLite.Open('test.db')<>0
     return 29
  endif

  if SQLite.DoCmd("DROP TABLE IF EXISTS blob_test; CREATE TABLE blob_test (data BLOB)")<>0
     return 30
  endif

  * Формируем бинарные данные (включая нулевой байт CHR(0) и символы > 127)
  cBinIn = CREATEBINARY("PNG" + CHR(0) + CHR(1) + CHR(255) + "TestBlobData")

  if SQLite.DoCmd("INSERT INTO blob_test (data) VALUES (?)", m.cBinIn)<>0
     return 31
  endif

  if SQLite.DoCmd("SELECT data FROM blob_test")<>0
     return 32
  endif

  if SQLite.Eof()=0
     arec = SQLite.Next()
     cBinOut = arec(1)
     ? "--- Проверка BLOB ---"
     ? "Тип VFP:", VARTYPE(m.cBinOut), "| Длина:", LEN(m.cBinOut), "| Совпадает:", (m.cBinIn == m.cBinOut)
     if VARTYPE(m.cBinOut) <> 'Q' OR !(m.cBinIn == m.cBinOut)
        SQLite.Close()
        return 33
     endif
  endif
  ?

  if SQLite.Close()<>0
     return 34
  endif

RETURN 0

* СЖАТИЕ И КОПИРОВАНИЕ БД
* COMPRESSION AND COPYING OF THE DB
FUNCTION Test7
  SQLite = CreateO('COM.SQLite'+iif(sys(17)='Pentium','32',''))
  bak = 'bak.test.db.bak'
  backup = 'bak.test.db'

  if SQLite.Open('test.db')<>0
     return 26
  endif

  if file(m.bak)
     dele file (m.bak)
  endif
  if file(m.backup)
     rena (m.backup) to (m.bak)
  endif

  if SQLite.DoCmd("VACUUM INTO ?",m.backup)<>0
     return 27
  endif

  if SQLite.Close()<>0
     return 28
  endif

RETURN 0
```
### Обсуждение
Задать вопрос или обсудить тему, касающуюся SQLite или VFP/VFPA, вы можете в разделе проекта `Issues` > `New issue`.
### История версий
0.0.0.0. 18.04.2025. Опубликована первая рабочая версия.  
0.1.0.0. 24.03.2026. Теперь используется встроенная в Windows10/11 winsqlite3.dll.  
0.2.0.0. 13.11.2026. Исправлено прямое чтение чисел с плавающей точкой (Double/Real), добавлен возврат системного .NULL., устранена гонка потоков при быстрой смене запросов с параметрами, оптимизирована работа с памятью (чтение строк и прямая передача BLOB без промежуточного копирования). Спасибо ИИ.  

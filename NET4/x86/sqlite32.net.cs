//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//!!!                                                   !!!
//!!!  SQLite32.net на C#.        Автор: A.Б.Корниенко  !!!
//!!!  v0.0.0.0                             18.04.2025  !!!
//!!!                                                   !!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

using System;
using System.Text;
using System.Globalization;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace sqlite {

  [ClassInterface(ClassInterfaceType.None)]
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  [ProgId("COM.SQLite32")]  //!!
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  public class sqlite {

    Encoding vfpw = Encoding.GetEncoding(1251); // подходит для двоичных данных
    const int k_1 = -1, SQLITE_OK = 0, SQLITE_NOMEM = 7, SQLITE_ROW = 100,
      SQLITE3_INTEGER = 1, SQLITE3_FLOAT = 2, SQLITE3_TEXT = 3, SQLITE3_BLOB = 4,
      SQLITE3_NULL = 5;
    const string se = ";";

    int iCol, ise, ise1, ise2, i, i1, i2, i3, nc, nb, ret, ret1, ret2, nSql;
    IntPtr p_db, p_smt, p_smt2, p_tail;
    object[] vals, aret, aret2;
    char[] qu = {'"','\''};
    string sql1, sql2;
    bool l2, l3, l4;
    byte[] b;
    double d;
//    Array a;
    Task ts;

    // Используемые методы в sqlite3.dll
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_open_v2(string db, out IntPtr p_db, int rw, string vfs);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_close_v2(IntPtr p_db);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_prepare_v2(IntPtr p_db, string sql, int nbyte,
             out IntPtr p_smt, out IntPtr p_tail);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_step(IntPtr p_smt);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_column_count(IntPtr p_smt);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_column_type(IntPtr p_smt, int iCol);
    [DllImport("sqlite3.dll")]
    internal static extern IntPtr sqlite3_column_int(IntPtr p_smt, int iCol);
    [DllImport("sqlite3.dll")]
    internal static extern IntPtr sqlite3_column_int64(IntPtr p_smt, int iCol);
    [DllImport("sqlite3.dll")]
    internal static extern IntPtr sqlite3_column_text(IntPtr p_smt, int iCol);
    [DllImport("sqlite3.dll")]
    internal static extern IntPtr sqlite3_column_blob(IntPtr p_smt, int iCol);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_column_bytes(IntPtr p_smt, int iCol);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_finalize(IntPtr p_smt);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_bind_blob(IntPtr p_smt, int i1,
             byte[] b, int nb, int p_void);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_bind_text(IntPtr p_smt, int i1,
             string s, int nb, int p_void);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_bind_null(IntPtr p_smt, int i1);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_bind_double(IntPtr p_smt, int i1, double d);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_bind_int64(IntPtr p_smt, int i1, double i64);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_bind_int(IntPtr p_smt, int i1, int i);
    [DllImport("sqlite3.dll")]
    internal static extern int sqlite3_complete(string sql);

    // Открьть БД
    public int Open(string db, bool read_only = false) {
      if(read_only) {
        // SQLITE_OPEN_READONLY  = 0x00000001;
        // SQLITE_OPEN_URI(=64)  = 0x00000040;
        i = 65;
      } else {
        // SQLITE_OPEN_READWRITE = 0x00000002;
        // SQLITE_OPEN_CREATE    = 0x00000004;
        // SQLITE_OPEN_URI(=64)  = 0x00000040;
        i = 70;
      }
      try {
        ret = sqlite3_open_v2(db, out p_db, i, null);
      } catch(Exception) {
        ret = SQLITE_NOMEM;
      }
      return ret;
    }

    void getSql1(ref string sql) {
      ise2 = sql.IndexOf(se,ise2)+SQLITE3_INTEGER;
      if(ise2>SQLITE_OK) {
        sql1 = sql.Substring(ise, ise2-ise);
      } else {
        ise2 = sql.Length;
        sql1 = sql.Substring(ise)+";";
      }
    }

    int odd(string z) {
      return (z.Length - z.Replace("'", "").Length)%2 +
             (z.Length - z.Replace("\"", "").Length)%2;
    }

    void Exec(ref string sql) {
      ise = ise2 = SQLITE_OK;
      if(ts != null) ts.Wait();
      sqlite3_finalize(l3?p_smt2:p_smt2);
      l2 = l3 = l4 = false;
      nSql = sql.Length;
      i3 = SQLITE_OK;

      do {
        getSql1(ref sql);

        // Проверим ";" разделитель соманд или нет
        while (ise<nSql && sqlite3_complete(sql1)==SQLITE_OK) getSql1(ref sql);

        if(ts != null) ts.Wait();
        if(l4) {

          // Меняем p_smt на p_smt2 или наоборот
          l3 = !l3;
          l4 = false;
          ret1 = ret2;

        }
        if(l3) {
          ret = sqlite3_prepare_v2(p_db, sql1, k_1, out p_smt2, out p_tail);
        } else {
          ret = sqlite3_prepare_v2(p_db, sql1, k_1, out p_smt, out p_tail);
        }
        if(ret == SQLITE_OK) {
          sql2 = sql1;          // sql2 для асинхронной задачи ts
          ise = ise1 = ise2;    // ise1 для асинхронной задачи ts
          ts = Task.Run(() => {

            // Определить количество вопросов i2 в команде sql1/sql2
            i1 = i2 = SQLITE_OK;
            while (i1<sql2.Length) {
              i = sql2.IndexOf("?",i1);
              if(i<SQLITE_OK) {
                i1 = sql2.Length;
              } else {
                if(odd(sql2.Substring(i1, i-i1))==SQLITE_OK) i2++;
                i1 = i+SQLITE3_INTEGER;
              }
            }

            // Проконтролировать наличие параметров в vals
            if((i2+i3)>vals.Length) i2 = vals.Length - i3;

            if(i2>SQLITE_OK) {
              for(iCol=SQLITE3_INTEGER; iCol<=i2; iCol++) {
                switch(vals[i3].GetType().Name) {
                case "Int32":
                  sqlite3_bind_int(l3?p_smt2:p_smt, iCol, (int)vals[i3]);
                  break;
                case "Int64":
                  sqlite3_bind_int64(l3?p_smt2:p_smt, iCol, (double)vals[i3]);
                  break;
                case "Float":
                case "Double":
                  sqlite3_bind_double(l3?p_smt2:p_smt, iCol, (double)vals[i3]);
                  break;
                case "Boolean":
                  sqlite3_bind_text(l3?p_smt2:p_smt, iCol, ((bool)vals[i3])?"T":"", k_1, k_1);
                  break;
                case "DBNull":
                case "Null":
                  sqlite3_bind_null(l3?p_smt2:p_smt, iCol);
                  break;
                case "String":
                  sqlite3_bind_text(l3?p_smt2:p_smt, iCol, (string)vals[i3], k_1, k_1);
                  break;
                case "Byte[]":
                  sqlite3_bind_blob(l3?p_smt2:p_smt, iCol, (byte[])vals[i3],
                                               ((byte[])vals[i3]).Length, k_1);
                  break;
                case "Byte[*]":
                  nb = ((Array)vals[i3]).Length;
                  b = new byte[nb];
                  Array.Copy(((Array)vals[i3]), SQLITE3_INTEGER, b, SQLITE_OK, nb);
                  sqlite3_bind_blob(l3?p_smt2:p_smt, iCol, b, nb, SQLITE_OK);
                  break;
                }
                i3++;
              }
            }
            ret2 = sqlite3_step(l3?p_smt2:p_smt);
            if(ise1 == nSql) Step2(l3?p_smt2:p_smt);
          });
        } else {
          ise2 = nSql;
        }
      } while (ise2<nSql);
      if(!l4) ret2 = ret1;
    }

    // Вынесено т.к. используестя в 2-х местах. Для чтения текста и полей Double
    void SqlToText1() {
      nb = sqlite3_column_bytes(p_smt, iCol);
      b = new byte[nb];
      Marshal.Copy((IntPtr)sqlite3_column_text(p_smt, iCol), b, SQLITE_OK, nb);
    }

    // Вызов из Next
    void Step() {

      // Если в аоследней команде не было обнаружено SQLITE_ROW
      if(!l4) l3 = !l3;

      ret2 = sqlite3_step(l3?p_smt2:p_smt);
      Step2(l3?p_smt2:p_smt);
    }

    void Step2(IntPtr p_smt) {
      if(ret2 == SQLITE_ROW) {
        l4 = true;
        nc = sqlite3_column_count(p_smt);
        if(l2) {
          aret2 = new object[nc];
        } else {
          aret = new object[nc];
        }
        for(iCol = SQLITE_OK; iCol < nc; iCol++) {
          switch(sqlite3_column_type(p_smt, iCol)) {
          case 1:   // SQLITE3_INTEGER
            if(sqlite3_column_bytes(p_smt, iCol)<10) {
              if(l2) {
                aret2[iCol] = (int)sqlite3_column_int(p_smt, iCol);
              } else {
                aret[iCol] = (int)sqlite3_column_int(p_smt, iCol);
              }
            } else {
              if(l2) {
                aret2[iCol] = (double)sqlite3_column_int64(p_smt, iCol);
              } else {
                aret[iCol] = (double)sqlite3_column_int64(p_smt, iCol);
              }
            }
            break;
          case 2:   // SQLITE3_FLOAT
            // Double не извлекается из SQLite3 (баг). Обходной путь через текст.

            // Использовать: Double.TryParse().
            // Безопасная альтернатива Double.Parse(), так как при неудачном
            // преобразовании не возникает исключение. Метод возвращает логическое
            // значение, указывающее на успех преобразования.
            //              Convert.ToDouble().
            // Может обрабатывать пустые строки и возвращает 0 при неудачном
            // преобразовании.
            SqlToText1();
            string dStr = vfpw.GetString(b).Replace(',','.');
            if(!Double.TryParse(dStr, NumberStyles.Any,
                CultureInfo.InvariantCulture, out d)) {
              try{ 
                d = Convert.ToDouble(dStr);
              } catch(Exception) {
                d = 0;
              }
            }

            if(l2) {
              aret2[iCol] = d;
            } else {
              aret[iCol] = d;
            }
            break;
          case 3:   // SQLITE3_TEXT
            SqlToText1();
            if(l2) {
              aret2[iCol] = vfpw.GetString(b);
            } else {
              aret[iCol] = vfpw.GetString(b);
            }
            break;
          case 5:   // SQLITE3_NULL
            if(l2) {
              aret2[iCol] = null;
            } else {
              aret[iCol] = null;
            }
            break;
          default:  // SQLITE3_BLOB
            nb = sqlite3_column_bytes(p_smt, iCol);

//            // Танцы с бубном для формирования массива с индекса 1
//            b = new byte[nb];
//            a = Array.CreateInstance(typeof(byte), new[] { nb }, new[] { 1 });
//            Marshal.Copy(inPt, b, 0, nb);
//            Array.Copy(b, 0, a, 1, nb);
//            aret[iCol] = a;

            // Работает с нормальным массивом
            if(l2) {
              aret2[iCol] = new byte[nb];
              Marshal.Copy((IntPtr)sqlite3_column_blob(p_smt, iCol),
                          ((byte[])aret2[iCol]), SQLITE_OK, nb);
            } else {
              aret[iCol] = new byte[nb];
              Marshal.Copy((IntPtr)sqlite3_column_blob(p_smt, iCol),
                          ((byte[])aret[iCol]), SQLITE_OK, nb);
            }
            break;
          }
        }
      } else {
        sqlite3_finalize(p_smt);
      }
    }

    // Подготовить запрос имея до 10 параметров
    public int DoCmd(string sql, object v1 = null, object v2 = null,
                     object v3 = null, object v4 = null, object v5 = null,
                     object v6 = null, object v7 = null, object v8 = null,
                     object v9 = null, object v10 = null) {
      if(!(v1 != null)) {
        vals = new object[] { };
      } else if(!(v2 != null)) {
        vals = new object[] { v1 };
      } else if(!(v3 != null)) {
        vals = new object[] { v1,v2 };
      } else if(!(v4 != null)) {
        vals = new object[] { v1,v2,v3 };
      } else if(!(v5 != null)) {
        vals = new object[] { v1,v2,v3,v4 };
      } else if(!(v6 != null)) {
        vals = new object[] { v1,v2,v3,v4,v5 };
      } else if(!(v7 != null)) {
        vals = new object[] { v1,v2,v3,v4,v5,v6 };
      } else if(!(v8 != null)) {
        vals = new object[] { v1,v2,v3,v4,v5,v6,v7 };
      } else if(!(v9 != null)) {
        vals = new object[] { v1,v2,v3,v4,v5,v6,v7,v8 };
      } else if(!(v10 != null)) {
        vals = new object[] { v1,v2,v3,v4,v5,v6,v7,v8,v9 };
      } else {
        vals = new object[] { v1,v2,v3,v4,v5,v6,v7,v8,v9,v10 };
      }
      Exec(ref sql);
      return ret;
    }

    // Подготовить запрос с массивом параметров
    public int DoCmdN(string sql, ref object[] valsN) {
      vals = (object[])valsN.Clone();
      Exec(ref sql);
      return ret;
    }

    // Конец строк?
    public int Eof() {
      if(ts != null) ts.Wait();
      return (ret2==SQLITE_ROW)? SQLITE_OK : k_1;
    }

    // Подготовить запрос
    public object Next() {
      if(ts != null) ts.Wait();
      if(ret2 == SQLITE_ROW) {
        if(l2) {
          l2 = false;
          ts = Task.Run(() => { Step(); });
          return aret2;
        } else {
          l2 = true;
          ts = Task.Run(() => { Step(); });
          return aret;
        }
      }
      return null;
    }

    // Закрыть БД
    public int Close() {
      if(ts != null) ts.Wait();
      return sqlite3_close_v2(p_db);
    }
  }
}

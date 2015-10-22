using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReaderFromExcel
{
    class ExcelWorkBookToDatabaseWriter : IExcelWorkBookWriter
    {
        enum NamesOfColumn { OldCode, NewCode, Good, Mj, Quantity, OldPrice, NewPrice, Ean };

        //public string CreateTable()
        //{

        //}

        public void WriteExcelWorkBook(UploadDocument workBook, bool rewrite)
        {
            

            Column col = new Column();
            //string connectionString = "Data Source=ServerName; Initial Catalog=DatabaseName;User ID=username;Password=password";

            //string sqlCreate = "Create Table" + workBook.Name + "(";
            //foreach (var column in col)
            //{
            //    sqlCreate += column.Name;

            //    if (column.Type == "int")
            //    {
            //        sqlCreate += " int ";
            //    }

            //    else if (column.Type == "long")
            //    {
            //        sqlCreate += " bigint ";
            //    }

            //    else if (column.Type == "double")
            //    {
            //        sqlCreate += " decimal(6,2) ";
            //    }

            //    else if (column.Type == "string")
            //    {
            //        sqlCreate += " varchar(MAX) ";
            //    }

            //    sqlCreate += ",";
            //}

            //sqlCreate += ");";

            //using (SqlConnection conn = new SqlConnection(connectionString))
            //{
            //    conn.Open();
            //    SqlCommand cmd = new SqlCommand(sqlCreate, conn);
            //    cmd.ExecuteNonQuery();
            //    cmd.Dispose();
            //    cmd.Clone();
            //}
        }

  
    }
}

using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReaderFromExcel
{
    class ExcelWorkBookReader
    {
        public ExcelWorkbook ReadExcelWorkBook(string filename)
        {
            ExcelWorkbook workBook = new ExcelWorkbook();

            List<UploadDocument> list = new List<UploadDocument>();

            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();

                //Dopredu neviem nazov tabulky
                OleDbCommand command = new OleDbCommand("select * from [Nove$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                    while (dr.Read())
                    {
                        if (dr[0].ToString() != "")
                        {
                            list.Add(
                                new UploadDocument
                                {
                                    oldCode = dr[0].ToString(),
                                    newCode = Convert.ToInt32(dr[1]),
                                    good = dr[2].ToString(),
                                    mj = dr[3].ToString(),
                                    quantity = Convert.ToInt32(dr[4]),
                                    oldPrice = Convert.ToDouble(dr[5]),
                                    newPrice = Convert.ToDouble(dr[7]),
                                    EAN = Convert.ToInt64(dr[8])
                                });
                        }
                    }

                connection.Close();
            }
            return workBook;

        }
    }
}

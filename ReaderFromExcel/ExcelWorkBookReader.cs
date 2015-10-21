using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace ReaderFromExcel
{
    class ExcelWorkBookReader
    {
        public string[] NamesOfSheets(string filename)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();

                //int i = 0;
                string[] excelSheets;
                int j = 0;

                try
                {
                    DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    excelSheets = dt.Rows.Cast<DataRow>()
                        .Where(i => i["TABLE_NAME"].ToString().EndsWith("$'"))
                        .Select(i => i["TABLE_NAME"].ToString()).ToArray();

                    return excelSheets;
                    //foreach (DataRow item in dt.Rows)
                    //{
                    //    if (item["TABLE_NAME"].ToString().EndsWith("$'"))
                    //    {
                    //        i++;
                    //    }
                    //}

                    //if (i > 0)
                    //{
                    //    excelSheets = new string[i];
                    //    foreach (DataRow item in dt.Rows)
                    //    {
                    //        excelSheets[j] = item["TABLE_NAME"].ToString();
                    //        j++;
                    //    }
                    //    return excelSheets;
                    //}
                }
                catch (Exception)
                {
                    throw new Exception("Failed to get SheetName");

                }
            }

            return null;

        }

        public UploadDocument ReadExcelWorkBook(string filename, string[] sheets)
        {
            UploadDocument workBook = new UploadDocument();

            List<UploadDocument> list = new List<UploadDocument>();

            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();


                if (sheets != null)
                {
                    foreach (var item in sheets)
                    {
                        OleDbCommand command = new OleDbCommand("select * from [" + item + "]", connection);
                        using (OleDbDataReader dr = command.ExecuteReader())
                            while (dr.Read() && dr.IsDBNull(0) == false)
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
                }
            }
     
            return workBook;

        }
    }
}

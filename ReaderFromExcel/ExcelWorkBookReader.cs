using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Xml;
using System.IO;


namespace ReaderFromExcel
{
    class ExcelWorkBookReader
    {

        public List<string> NamesOfSheets(string filename)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";

            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                try
                {
                    List<string> excelSheets = new List<string>();
                    DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    excelSheets = dt.Rows.Cast<DataRow>()
                        .Where(i => i["TABLE_NAME"].ToString().EndsWith("$") || i["TABLE_NAME"].ToString().EndsWith("$'"))
                        .Select(i => i["TABLE_NAME"].ToString()).ToList();

                    return excelSheets;
                }
                catch (Exception)
                {
                    throw new Exception("Failed to get SheetName");
                }
            }
        }

        public UploadDocument ReadExcelWorkBook(string filename, List<string> sheets)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";

            UploadDocument workBook = new UploadDocument();

            List<UploadDocument> list = new List<UploadDocument>();

            using (OleDbConnection connection = new OleDbConnection(con))
            {
                ExcelWorkbook work = new ExcelWorkbook();

                connection.Open();

                if (sheets != null)
                {
                    foreach (var item in sheets)
                    {
                        //work.Sheets = ;
                        new ExcelWorkSheet
                        {
                            name = item.ToString()
                        };
                    }
                }
            }


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

        public UploadDocument2 ReadExcelWorkBook2(string filename, List<string> sheets)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";

            UploadDocument2 workBook = new UploadDocument2();

            List<UploadDocument2> list = new List<UploadDocument2>();


            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();


                if (sheets != null)
                {
                        OleDbCommand command = new OleDbCommand("select * from [" + sheets[1] + "]", connection);
                        using (OleDbDataReader dr = command.ExecuteReader())
                            while (dr.Read() && dr.IsDBNull(0) == false)
                            {
                            list.Add(
                                new UploadDocument2
                                {
                                    cisloBalicka = Convert.ToInt32(dr[0]),
                                    typ = dr[1].ToString(),
                                    upL1 = Convert.ToInt32(dr[2]),
                                    upL1String = dr[3].ToString(),
                                    upL2 = Convert.ToInt32(dr[4]),
                                    upL2String = dr[5].ToString(),
                                    minCenaUpl1 = Convert.ToInt64(dr[6]),
                                    minCenaUpl2 = Convert.ToInt64(dr[7]),
                                    minMnozstvoUpl1 = Convert.ToInt32(dr[8]),
                                    akciovaCenaUpl1 = Convert.ToInt64(dr[9]),
                                    minMnozstvoUpl2 = Convert.ToInt32(dr[10]),
                                    akciovaCenaUpl2 = Convert.ToInt64(dr[11]),
                                    zlava = Convert.ToInt64(dr[12]),
                                    cenaBalicka = Convert.ToInt64(dr[13]),
                                    platnostOd = Convert.ToDateTime(dr[14]),
                                    platnostDo = Convert.ToDateTime(dr[15]),
                                    nazovBalicka = dr[16].ToString()
                                    });
                            }   
                }
            }
            return workBook;
        }
    }
}


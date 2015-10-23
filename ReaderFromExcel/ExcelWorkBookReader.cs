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


        public EOrderRow ReadEbook(string filename, List<string> sheets)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";

            EOrderRow workBook = new EOrderRow();

            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();

                List<EOrderRow> list = new List<EOrderRow>();

                foreach (var item in sheets)
                {
                    OleDbCommand command = new OleDbCommand("select * from [" + item + "]", connection);
                    using (OleDbDataReader dr = command.ExecuteReader())
                        while (dr.Read() && dr.IsDBNull(0) == false)
                        {
                            try
                            {
                                list.Add(
                                new EOrderRow
                                {
                                    oldCode = dr["Slovnaft kód"].ToString(),
                                    newCode = Convert.ToInt32(dr["Nový SAP kód"]),
                                    good = dr["Tovar"].ToString(),
                                    mj = dr["MJ"].ToString(),
                                    quantity = Convert.ToInt32(dr["Balenie (kusov v kartóne)"]),
                                    newPrice = Convert.ToDouble(dr["SN ČS DDU 10.2.2015 EUR/KS"]),
                                    UPL = Convert.ToInt32(dr["UPL"]),
                                    EAN = Convert.ToInt64(dr["EAN"])
                                });
                            }
                            catch (IndexOutOfRangeException)
                            {
                                throw new IndexOutOfRangeException("Nazov stlpca v exceli nezodpoveda preddefinovanemu nazvu");

                            }
                        }
                }
            }

            return workBook;
        }


        //public ExcelWorkbook ReadExcelWorkBook(string filename, List<string> sheets)
        //{
        //    string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";

        //    ExcelWorkbook workBook = new ExcelWorkbook();

        //    foreach (var item in sheets)
        //    {
        //        workBook.Sheets.Add(ReadSheet(filename, item));
        //    }

        //    return workBook;
        //}


        //public ExcelWorkSheet ReadSheet(string filename, string name)
        //{
        //    string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=No;'";

        //    ExcelWorkSheet sheet = new ExcelWorkSheet();

        //    int r = 0;
        //    using (OleDbConnection connection = new OleDbConnection(con))
        //    {
        //        connection.Open();

        //        OleDbCommand command = new OleDbCommand("select * from [" + name + "]", connection);
        //        using (OleDbDataReader dr = command.ExecuteReader())
        //        {
        //            while (dr.Read() && dr.IsDBNull(0) == false)
        //            {
        //                if (r == 0)
        //                {
        //                    sheet.Columns = new Dictionary<int, Column>();

        //                    for (int i = 0; dr.IsDBNull(i) == false; i++)
        //                    {
        //                        sheet.Columns.Add(i,
        //                        new Column
        //                        {
        //                            Name = dr.GetValue(i).ToString(),
        //                            Order = i,
        //                            Type = dr.GetFieldType(i).FullName
        //                        });

        //                    }
        //                    r += 1;
        //                }
        //                else
        //                {
        //                    Row row = new Row();
        //                    row.Data = new Dictionary<int, Data>();

        //                    for (int i = 0; dr.IsDBNull(i) == false; i++)
        //                    {
        //                        row.Data.Add(i, new Data()
        //                        {
        //                            Value = dr.GetValue(i).ToString()
        //                        });
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    return sheet;
        //}

        public AbRow ReadAbRow(string filename, List<string> sheets)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";

            AbRow sheet1 = new AbRow();

            List<AbRow> list = new List<AbRow>();

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
                                new AbRow
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
            return sheet1;
        }

        public CodeRow ReadCodeRow(string filename, List<string> sheets)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=Yes;'";

            CodeRow codeRow = new CodeRow();

            List<CodeRow> list = new List<CodeRow>();

            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();

                if (sheets != null)
                {
                    OleDbCommand command = new OleDbCommand("select * from [" + sheets + "]", connection);
                    using (OleDbDataReader dr = command.ExecuteReader())
                        while (dr.Read() && dr.IsDBNull(0) == false)
                        {
                            list.Add(
                                new CodeRow
                                {


                                });
                        }
                }
            }
            return codeRow;
        }
    }
    
}



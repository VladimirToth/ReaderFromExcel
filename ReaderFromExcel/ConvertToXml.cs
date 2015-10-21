using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

//namespace ReaderFromExcel
//{
//    class ConvertToXml
//    {
//        private DataTable ReadExcelFile(string filename)
//        {
//            DataTable dt = new DataTable();

//            try
//            {
//                //Otvorenie pripojenia k excelovskemu dokumentu
//                using (SpreadsheetDocument sheetDocument = SpreadsheetDocument.Open(filename, false))
//                {
//                    WorkbookPart workBookPart = sheetDocument.WorkbookPart;

//                    // Nacitanie vsetkych listov v zosite do enumerovatelnej kolekcie
//                    IEnumerable<Sheet> sheetsCollection = sheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

//                    // Nacitanie ID konkretneho listu v excelovskom dokumente. Zatial som zadal len posledny list. Treba dorobit, aby sa nacitali vsetky listy...........
//                    string relationshipId = sheetsCollection.Last().Id.Value;
//                        //sheetsCollection.First().Id.Value;

//                    // Vytvorenie instancie konkretneho listu podla zadaneho id 
//                    WorksheetPart worksheetPart = (WorksheetPart)sheetDocument.WorkbookPart.GetPartById(relationshipId);

//                    // Nacitanie udajov z excelovskeho dokumentu
//                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

//                    // Nacitanie poctu riadkov v konkretnom liste do kolekcie
//                    IEnumerable<Row> rowcollection = sheetData.Descendants<Row>();

//                    if (rowcollection.Count() == 0)
//                    {
//                        return dt;
//                    }

//                    // Pridanie nadpisu stlpcov do DataTable s vytvorenim nadpisu s udajov uvedenych v prvom riadku listu
//                    foreach (Cell cell in rowcollection.ElementAt(0))
//                    {
//                        dt.Columns.Add(GetValueOfCell(sheetDocument, cell));
//                    }


//                    // Pridanie riadkov do DataTable 
//                    foreach (Row row in rowcollection)
//                    {
//                        DataRow temprow = dt.NewRow();
//                        int columnIndex = 0;
//                        foreach (Cell cell in row.Descendants<Cell>())
//                        {
//                            // Get Cell Column Index 
//                            int cellColumnIndex = GetColumnIndex(GetColumnName(cell.CellReference));


//                            if (columnIndex < cellColumnIndex)
//                            {
//                                do
//                                {
//                                    temprow[columnIndex] = string.Empty;
//                                    columnIndex++;
//                                }


//                                while (columnIndex < cellColumnIndex);
//                            }


//                            temprow[columnIndex] = GetValueOfCell(sheetDocument, cell);
//                            columnIndex++;
//                        }


//                        // Add the row to DataTable 
//                        // the rows include header row 
//                        dt.Rows.Add(temprow);
//                    }
//                }

//                dt.Rows.RemoveAt(0);
//                return dt;
//            }
//            catch (Exception ex)
//            {

//                throw new Exception(ex.Message);
//            }
//        }


//        //Funkcia na nacitanie hodnoty z konkretnej bunky
//        private static string GetValueOfCell(SpreadsheetDocument spreadsheetdocument, Cell cell)
//        {
//            SharedStringTablePart sharedString = spreadsheetdocument.WorkbookPart.SharedStringTablePart;
//            if (cell.CellValue == null)
//            {
//                return string.Empty;
//            }


//            string cellValue = cell.CellValue.InnerText;

//            // The condition that the Cell DataType is SharedString 
//            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
//            {
//                return sharedString.SharedStringTable.ChildElements[int.Parse(cellValue)].InnerText;
//            }
//            else
//            {
//                return cellValue;
//            }
//        }

//        //Funkcia na zistenie nazvu stlpca
//        private string GetColumnName(string cellReference)
//        {
//            Regex regex = new Regex("[A-Za-z]+");
//            Match match = regex.Match(cellReference);
//            return match.Value;
//        }

//        /// <summary> 
//        /// Get Index of Column from given column name 
//        /// </summary> 
//        /// <param name="columnName">Column Name(For Example,A or AA)</param> 
//        /// <returns>Column Index</returns> 
//        private int GetColumnIndex(string columnName)
//        {
//            int columnIndex = 0;
//            int factor = 1;

//            // From right to left 
//            for (int position = columnName.Length - 1; position >= 0; position--)
//            {
//                // For letters 
//                if (Char.IsLetter(columnName[position]))
//                {
//                    columnIndex += factor * ((columnName[position] - 'A') + 1) - 1;
//                    factor *= 26;
//                }
//            }

//            return columnIndex;
//        }

//        public string GetXML(string filename)
//        {
//            using (DataSet ds = new DataSet())
//            {
//                ds.Tables.Add(this.ReadExcelFile(filename));
//                return ds.GetXml();
//            }
//        }
//    }
//}

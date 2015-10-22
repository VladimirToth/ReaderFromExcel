using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReaderFromExcel
{
    class ExcelWorkBookToXmlWriter:IExcelWorkBookWriter
    {
        public void WriteExcelWorkBook(ExcelWorkbook workBook, bool rewrite)
        {
            XmlSerializer x = new XmlSerializer(typeof(ExcelWorkbook));
            using (FileStream file = new FileStream("cesta.xml", FileMode.OpenOrCreate))
            using (TextWriter tw = new StreamWriter(file))
            {
                x.Serialize(tw, workBook);
            }
        }
    }
}

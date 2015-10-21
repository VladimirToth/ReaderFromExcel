using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReaderFromExcel
{
    class ExcelWorkbook
    {
        public string Name { get; set; }
        public List<ExcelWorkSheet> Sheets { get; set; }

    }
}

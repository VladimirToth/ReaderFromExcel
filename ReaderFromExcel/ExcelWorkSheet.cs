using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReaderFromExcel
{
    class ExcelWorkSheet
    {
        public string name { get; set; }
        public Dictionary<int, Column> Columns { get; set; }
        public List<Row> Rows { get; set; }

        public IEnumerable<Data> GetDataInColumn(int index)
        {

            return Rows.Select(r => r.Data[index]);
        }
    }
}

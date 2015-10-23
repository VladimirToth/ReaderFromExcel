using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReaderFromExcel
{
    class EOrderRow
    {
        public string oldCode { get; set; }
        public int newCode { get; set; }
        public string good { get; set; }
        public string mj { get; set; }
        public int quantity { get; set; }
        public double newPrice { get; set; }
        public int UPL { get; set; }
        public long EAN { get; set; }
    }
}

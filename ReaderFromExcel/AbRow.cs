using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReaderFromExcel
{
    class AbRow
    {
        public int cisloBalicka { get; set; }
        public string typ { get; set; }
        public int upL1 { get; set; }
        public string upL1String { get; set; }
        public int upL2 { get; set; }
        public string upL2String { get; set; }
        public double minCenaUpl1 { get; set; }
        public double minCenaUpl2 { get; set; }
        public int minMnozstvoUpl1 { get; set; }
        public double akciovaCenaUpl1 { get; set; }
        public int minMnozstvoUpl2 { get; set; }
        public double akciovaCenaUpl2 { get; set; }
        public double zlava { get; set; }
        public double cenaBalicka { get; set; }
        public DateTime platnostOd { get; set; }
        public DateTime platnostDo { get; set; }
        public string nazovBalicka { get; set; }
        public List<CodeRow> kody { get; set; }
    }
}

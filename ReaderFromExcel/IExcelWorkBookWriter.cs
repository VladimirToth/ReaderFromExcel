﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReaderFromExcel
{
    interface IExcelWorkBookWriter
    {
        void WriteExcelWorkBook(EOrderRow workBook, bool rewrite);
    }
}

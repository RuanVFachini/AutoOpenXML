using System;
using ClosedXML.Excel;

namespace AutoOpenXml.Models
{
    internal class CellRead
    {
        public XLDataType Type { get; set; }
        public Object Value { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace AutoOpenXml.Models
{
    internal class ColumnInfo
    {
        public string Label { get; set; }
        public string Name { get; set; }
        public int Index { get; set; }
        public ColumnTypes Type { get; set; }
        public Func<Object, Object> Func { get; set; }
    }
}

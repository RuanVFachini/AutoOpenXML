using System;
using System.Drawing;

namespace AutoOpenXml.Models
{
    internal class ColumnInfo
    {
        public string Label { get; set; }
        public string Name { get; set; }
        public int Index { get; set; }
        public Type Type { get; set; }
        public Color HeaderBackgroundColor { get; set; }
    }
}

﻿using System;

namespace AutoOpenXml
{
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ExportColumnAttribute : System.Attribute
    {
        private readonly string AutoOpenXml_Header;
        private readonly int AutoOpenXml_Index;

        public ExportColumnAttribute(string header, int index = -1)
        {
            AutoOpenXml_Header = header;
            AutoOpenXml_Index = index;
        }
    }
}

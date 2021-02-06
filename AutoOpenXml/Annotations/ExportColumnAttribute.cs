namespace AutoOpenXml
{
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ExportColumnAttribute : System.Attribute
    {
        private readonly string AutoOpenXml_Header;
        private readonly int AutoOpenXml_Index;
        private readonly ColumnTypes AutoOpenXml_Type;

        public ExportColumnAttribute(string header, int index, ColumnTypes type)
        {
            AutoOpenXml_Header = header;
            AutoOpenXml_Index = index;
            AutoOpenXml_Type = type;
        }
    }
}

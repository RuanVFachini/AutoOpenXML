namespace AutoOpenXml
{
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ExportColumnAttribute : System.Attribute
    {
        private readonly string AutoOpenXml_Header;
        private readonly int AutoOpenXml_Index;
        private readonly string AutoOpenXml_Mask;

        public ExportColumnAttribute(string header, int index = -1, string mask = "")
        {
            AutoOpenXml_Header = header;
            AutoOpenXml_Index = index;
            AutoOpenXml_Mask = mask;
        }
    }
}

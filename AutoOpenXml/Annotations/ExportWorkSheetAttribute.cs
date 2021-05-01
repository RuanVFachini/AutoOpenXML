using ClosedXML.Excel;

namespace AutoOpenXml
{
    [System.AttributeUsage(System.AttributeTargets.Class)]
    public class ExportWorkSheetAttribute : System.Attribute
    {
        private readonly string AutoOpenXml_Name;
        private readonly string AutoOpenXml_Style;

        public ExportWorkSheetAttribute(string name)
        {
            AutoOpenXml_Name = name;
            AutoOpenXml_Style = null;
        }

        public ExportWorkSheetAttribute(string name, string style)
        {
            AutoOpenXml_Name = name;
            AutoOpenXml_Style = style;
        }
    }
}

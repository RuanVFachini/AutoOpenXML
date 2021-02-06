namespace AutoOpenXml
{
    [System.AttributeUsage(System.AttributeTargets.Class)]
    public class ExportWorkSheetAttribute : System.Attribute
    {
        private readonly string AutoOpenXml_Name;

        public ExportWorkSheetAttribute(string name)
        {
            AutoOpenXml_Name = name;
        }
    }
}

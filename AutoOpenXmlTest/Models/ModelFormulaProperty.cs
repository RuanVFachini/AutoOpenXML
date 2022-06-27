using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelFormulaProperty.WorksheetName)]
    public class ModelFormulaProperty
    {
        [ExportColumn(VariablesModelFormulaProperty.FieldName, 1)]
        public int Formula { get; set; }

        public ModelFormulaProperty() { }

        public ModelFormulaProperty(int formula)
        {
            Formula = formula;
        }
    }

    public static class VariablesModelFormulaProperty
    {
        public const string WorksheetName = "FormulaProperty";
        public const string FieldName = "Formula";
    }
}

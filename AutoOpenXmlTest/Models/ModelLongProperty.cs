using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelLongProperty.WorksheetName)]
    public class ModelLongProperty
    {
        public string Name { get; set; }
        
        [ExportColumn(VariablesModelDecimalProperty.FieldName, 2)]
        public long Id { get; set; }

        public ModelLongProperty() { }

        public ModelLongProperty(string name, long id)
        {
            Name = name;
            Id = id;
        }
    }

    public static class VariablesModelLongProperty
    {
        public const string WorksheetName = "LongProperty";
        public const string FieldName = "Idade da Pessoa";
        public const string FieldFormat = "";
        public static readonly ModelLongProperty[] Data = {
        new ModelLongProperty("Joao", 11),
        new ModelLongProperty("Carlos", 12),
        };
    }
}

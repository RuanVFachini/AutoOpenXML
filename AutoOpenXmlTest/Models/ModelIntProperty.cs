using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelIntProperty.WorksheetName)]
    public class ModelIntProperty
    {
        public string Name { get; set; }
        [ExportColumn(VariablesModelIntProperty.FieldName, 3)]
        public int Age { get; set; }

        public ModelIntProperty() { }

        public ModelIntProperty(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    public static class VariablesModelIntProperty
    {
        public const string WorksheetName = "IntProperty";
        public const string FieldName = "Idade da Pessoa";
        public static readonly ModelIntProperty[] Data = {
            new ModelIntProperty("Joao", 15),
            new ModelIntProperty("Carlos", 32),
        };
    }
}

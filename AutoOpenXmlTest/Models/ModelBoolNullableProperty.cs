using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelBoolNullableProperty.WorksheetName)]
    public class ModelBoolNullableProperty
    {
        public string Name { get; set; }

        [ExportColumn(VariablesModelBoolNullableProperty.FieldName, 5)]
        public bool? Active { get; set; }

        public ModelBoolNullableProperty() { }

        public ModelBoolNullableProperty(string name, bool? active)
        {
            Name = name;
            Active = active;
        }
    }

    public static class VariablesModelBoolNullableProperty
    {
        public const string WorksheetName = "BoolNullableProperty";
        public const string FieldName = "Ativo";
        public const string FieldFormat = "";
        public static readonly ModelBoolNullableProperty[] Data = {
            new ModelBoolNullableProperty("Joao", true),
            new ModelBoolNullableProperty("Carlos", false),
            new ModelBoolNullableProperty("André", null)
        };
    }
}

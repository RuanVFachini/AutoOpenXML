using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelLongNullableProperty.WorksheetName)]
    public class ModelLongNullableProperty
    {
        public string Name { get; set; }
        
        [ExportColumn(VariablesModelDecimalProperty.FieldName, 2)]
        public long? Id { get; set; }

        public ModelLongNullableProperty() { }

        public ModelLongNullableProperty(string name, long? id)
        {
            Name = name;
            Id = id;
        }
    }

    public static class VariablesModelLongNullableProperty
    {
        public const string WorksheetName = "LongNullableProperty";
        public const string FieldName = "Idade da Pessoa";
        public const string FieldFormat = "";
        public static readonly ModelLongNullableProperty[] Data = {
        new ModelLongNullableProperty("Joao", 11),
        new ModelLongNullableProperty("Carlos", 12),
        new ModelLongNullableProperty("Carlos", null),
        };
    }
}

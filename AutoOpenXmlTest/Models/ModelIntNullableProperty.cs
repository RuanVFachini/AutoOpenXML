using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelIntNullableProperty.WorksheetName)]
    public class ModelIntNullableProperty
    {
        public string Name { get; set; }
        [ExportColumn(VariablesModelIntNullableProperty.FieldName, 3)]
        public int? Age { get; set; }

        public ModelIntNullableProperty() { }

        public ModelIntNullableProperty(string name, int? age)
        {
            Name = name;
            Age = age;
        }
    }

    public static class VariablesModelIntNullableProperty
    {
        public const string WorksheetName = "Nova Aba";
        public const string FieldName = "Idade da Pessoa";
        public static readonly ModelIntNullableProperty[] Data = {
            new ModelIntNullableProperty("Joao", 15),
            new ModelIntNullableProperty("Carlos", 32),
            new ModelIntNullableProperty("André", null),
        };
    }
}

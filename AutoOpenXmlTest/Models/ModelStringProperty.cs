using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelStringProperty.WorksheetName)]
    public class ModelStringProperty
    {
        [ExportColumn(VariablesModelStringProperty.FieldName, 1)]
        public string Name { get; set; }
        public int Age { get; set; }

        public ModelStringProperty()
        { }

        public ModelStringProperty(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    public static class VariablesModelStringProperty
    {
        public const string WorksheetName = "StringProperty";
        public const string FieldName = "Nome da Pessoa";
        public static readonly ModelStringProperty[] Data = {
            new ModelStringProperty("Joao", 15),
            new ModelStringProperty("Carlos", 32),
        };
    }
}
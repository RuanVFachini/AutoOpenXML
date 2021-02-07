using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Cases.ExportIntProperty
{
    [ExportWorkSheet(Variables.WorksheetName)]
    internal class ModelIntProperty
    {
        public string Name { get; set; }
        [ExportColumn(Variables.FieldName, 1, ColumnTypes.Int)]
        public int Age { get; set; }

        public ModelIntProperty() { }

        public ModelIntProperty(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    internal static class Variables
    {
        public const string WorksheetName = "Nova Aba";
        public const string FieldName = "Idade da Pessoa";
        public static readonly ModelIntProperty[] Data = {
            new ModelIntProperty("Joao", 15),
            new ModelIntProperty("Carlos", 32),
        };
    }
}

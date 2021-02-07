using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Cases.ExportStringProperty
{
    [ExportWorkSheet(Variables.WorksheetName)]
    internal class ModelStringProperty
    {
        [ExportColumn(Variables.FieldName, 1, ColumnTypes.Text)]
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

    internal static class Variables
    {
        public const string WorksheetName = "Nova Aba";
        public const string FieldName = "Nome da Pessoa";
        public static readonly ModelStringProperty[] Data = {
            new ModelStringProperty("Joao", 15),
            new ModelStringProperty("Carlos", 32),
        };
    }
}
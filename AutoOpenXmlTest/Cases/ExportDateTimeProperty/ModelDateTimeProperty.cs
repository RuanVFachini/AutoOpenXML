using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Cases.ExportDateTimeProperty
{
    [ExportWorkSheet(Variables.WorksheetName)]
    internal class ModelDateTimeProperty
    {
        public string Name { get; set; }
        [ExportColumn(Variables.FieldName, 1, ColumnTypes.DateTime)]
        public DateTime BirthDay { get; set; }

        public ModelDateTimeProperty() { }

        public ModelDateTimeProperty(string name, DateTime birthDay)
        {
            Name = name;
            BirthDay = birthDay;
        }
    }

    internal static class Variables
    {
        public const string WorksheetName = "Nova Aba";
        public const string FieldName = "Aniversário da Pessoa";
        public const string FieldFormat = "dd/mm/yyyy hh:mm:ss";
        public static readonly ModelDateTimeProperty[] Data = {
            new ModelDateTimeProperty("Joao", new DateTime(1991, 8, 28, 10, 30, 50)),
            new ModelDateTimeProperty("Carlos", new DateTime(1992, 7, 30, 10, 30, 50)),
        };
    }
}

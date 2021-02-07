using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Cases.ExportDecilmalProperty
{
    [ExportWorkSheet(Variables.WorksheetName)]
    internal class ModelDecimalProperty
    {
        public string Name { get; set; }
        [ExportColumn(Variables.FieldName, 1, ColumnTypes.Decimal)]
        public decimal Salary { get; set; }

        public ModelDecimalProperty() { }

        public ModelDecimalProperty(string name, decimal salary)
        {
            Name = name;
            Salary = salary;
        }
    }

    internal static class Variables
    {
        public const string WorksheetName = "Nova Aba";
        public const string FieldName = "salario da Pessoa";
        public const string FieldFormat = "";
        public static readonly ModelDecimalProperty[] Data = {
        new ModelDecimalProperty("Joao", (decimal)10.25),
        new ModelDecimalProperty("Carlos", (decimal) 12.584),
    };
    }
}

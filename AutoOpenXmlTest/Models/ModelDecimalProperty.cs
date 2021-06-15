using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelDecimalProperty.WorksheetName)]
    public class ModelDecimalProperty
    {
        public string Name { get; set; }
        [ExportColumn(VariablesModelDecimalProperty.FieldName, 2)]
        public decimal Salary { get; set; }

        public ModelDecimalProperty() { }

        public ModelDecimalProperty(string name, decimal salary)
        {
            Name = name;
            Salary = salary;
        }
    }

    public static class VariablesModelDecimalProperty
    {
        public const string WorksheetName = "DecimalProperty";
        public const string FieldName = "salario da Pessoa";
        public const string FieldFormat = "";
        public static readonly ModelDecimalProperty[] Data = {
        new ModelDecimalProperty("Joao", (decimal)10.25),
        new ModelDecimalProperty("Carlos", (decimal) 12.584),
        };
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelDecimalNullableProperty.WorksheetName)]
    public class ModelDecimalNullableProperty
    {
        public string Name { get; set; }
        [ExportColumn(VariablesModelDecimalNullableProperty.FieldName, 2)]
        public decimal? Salary { get; set; }

        public ModelDecimalNullableProperty() { }

        public ModelDecimalNullableProperty(string name, decimal? salary)
        {
            Name = name;
            Salary = salary;
        }
    }

    public static class VariablesModelDecimalNullableProperty
    {
        public const string WorksheetName = "DecimalNullableProperty";
        public const string FieldName = "salario da Pessoa";
        public const string FieldFormat = "";
        public static readonly ModelDecimalNullableProperty[] Data = {
        new ModelDecimalNullableProperty("Joao", (decimal)10.25),
        new ModelDecimalNullableProperty("Carlos", (decimal) 12.584),
        new ModelDecimalNullableProperty("André", null),
        };
    }
}

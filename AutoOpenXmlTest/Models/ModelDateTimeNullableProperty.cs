using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelDateTimeProperty.WorksheetName)]
    public class ModelDateTimeNullableProperty
    {
        public string Name { get; set; }
        [ExportColumn(VariablesModelDateTimeProperty.FieldName, 4)]
        public DateTime? BirthDay { get; set; }

        public ModelDateTimeNullableProperty() { }

        public ModelDateTimeNullableProperty(string name, DateTime? birthDay)
        {
            Name = name;
            BirthDay = birthDay;
        }
    }

    public static class VariablesModelDateTimeNullableProperty
    {
        public const string WorksheetName = "DateTimeNullableProperty";
        public const string FieldName = "Aniversário da Pessoa";
        public const string FieldFormat = "dd/mm/yyyy hh:mm:ss";
        public static readonly ModelDateTimeNullableProperty[] Data = {
            new ModelDateTimeNullableProperty("Joao", new DateTime(1991, 8, 28, 10, 30, 50)),
            new ModelDateTimeNullableProperty("Carlos", new DateTime(1992, 7, 30, 10, 30, 50)),
            new ModelDateTimeNullableProperty("André", null),
        };
    }
}

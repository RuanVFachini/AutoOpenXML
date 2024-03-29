﻿using System;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelDateTimeProperty.WorksheetName)]
    public class ModelDateTimeProperty
    {
        public string Name { get; set; }
        [ExportColumn(VariablesModelDateTimeProperty.FieldName, 4)]
        public DateTime BirthDay { get; set; }

        public ModelDateTimeProperty() { }

        public ModelDateTimeProperty(string name, DateTime birthDay)
        {
            Name = name;
            BirthDay = birthDay;
        }
    }

    public static class VariablesModelDateTimeProperty
    {
        public const string WorksheetName = "DateTimeProperty";
        public const string FieldName = "Aniversário da Pessoa";
        public const string FieldFormat = "dd/mm/yyyy hh:mm:ss";
        public static readonly ModelDateTimeProperty[] Data = {
            new ModelDateTimeProperty("Joao", new DateTime(1991, 8, 28, 10, 30, 50)),
            new ModelDateTimeProperty("Carlos", new DateTime(1992, 7, 30, 10, 30, 50)),
        };
    }
}

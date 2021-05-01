using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelBoolProperty.WorksheetName)]
    public class ModelBoolNullableProperty
    {
        public string Name { get; set; }

        [ExportColumn(VariablesModelBoolProperty.FieldName, 2)]
        public bool? Active { get; set; }

        public ModelBoolNullableProperty() { }

        public ModelBoolNullableProperty(string name, bool active)
        {
            Name = name;
            Active = active;
        }
    }

    public static class VariablesModelBoolNullableProperty
    {
        public const string WorksheetName = "Nova Aba";
        public const string FieldName = "Ativo";
        public const string FieldFormat = "";
        public static readonly ModelBoolNullableProperty[] Data = {
        new ModelBoolNullableProperty("Joao", true),
        new ModelBoolNullableProperty("Carlos", false),
        };
    }
}

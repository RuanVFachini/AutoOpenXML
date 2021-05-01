using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelBoolProperty.WorksheetName)]
    public class ModelBoolProperty
    {
        public string Name { get; set; }

        [ExportColumn(VariablesModelBoolProperty.FieldName, 5)]
        public bool Active { get; set; }

        public ModelBoolProperty() { }

        public ModelBoolProperty(string name, bool active)
        {
            Name = name;
            Active = active;
        }
    }

    public static class VariablesModelBoolProperty
    {
        public const string WorksheetName = "Nova Aba";
        public const string FieldName = "Ativo";
        public const string FieldFormat = "";
        public static readonly ModelBoolProperty[] Data = {
        new ModelBoolProperty("Joao", true),
        new ModelBoolProperty("Carlos", false),
        };
    }
}

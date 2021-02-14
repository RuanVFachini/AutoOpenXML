using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelOrderedProperties.WorksheetName)]
    public class ModelOrderedProperties
    {
        [ExportColumn(VariablesModelOrderedProperties.IdLabel, 2)]
        public int Id { get; set; }
        [ExportColumn(VariablesModelOrderedProperties.NameLabel, 1)]
        public string Name { get; set; }
        [ExportColumn(VariablesModelOrderedProperties.HeightLabel, 4)]
        public decimal Height { get; set; }
        [ExportColumn(VariablesModelOrderedProperties.BirthDateLabel, 3)]
        public DateTime BrithDate { get; set; }

        public ModelOrderedProperties() { }
        public ModelOrderedProperties(int id, string name, decimal height, DateTime brithDate)
        {
            Id = id;
            Name = name;
            Height = height;
            BrithDate = brithDate;
        }
    }

    public static class VariablesModelOrderedProperties
    {
        public const string WorksheetName = "Nova Aba";
        public const string NameLabel = "Nome";
        public const string IdLabel = "Código";
        public const string HeightLabel = "Altura";
        public const string BirthDateLabel = "Aniversário";
        public static readonly ModelOrderedProperties[] Data = {
            new ModelOrderedProperties(1,"Carlos", (decimal)19.85, new DateTime(1991, 8, 25))
        };
    }
}

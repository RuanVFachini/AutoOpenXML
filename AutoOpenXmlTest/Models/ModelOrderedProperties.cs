using System;
using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelOrderedProperties.WorksheetName, "TableStyleMedium26")]
    public class ModelOrderedProperties
    {
        [ExportColumn(VariablesModelOrderedProperties.IdLabel, 2)]
        public int Id { get; set; }
        [ExportColumn(VariablesModelOrderedProperties.NameLabel, 1)]
        public string Name { get; set; }
        [ExportColumn(VariablesModelOrderedProperties.HeightLabel, 4)]
        public decimal Height { get; set; }
        [ExportColumn(VariablesModelOrderedProperties.BirthDateLabel, 3)]
        public DateTime BirthDate { get; set; }

        public ModelOrderedProperties() { }
        public ModelOrderedProperties(int id, string name, decimal height, DateTime birthDate)
        {
            Id = id;
            Name = name;
            Height = height;
            BirthDate = birthDate;
        }
    }

    public static class VariablesModelOrderedProperties
    {
        public const string WorksheetName = "OrderedProperties";
        public const string NameLabel = "Nome";
        public const string IdLabel = "Código";
        public const string HeightLabel = "Altura";
        public const string BirthDateLabel = "Aniversário";
        public static readonly ModelOrderedProperties[] Data = {
            new ModelOrderedProperties(1,"Carlos", (decimal)19.85, new DateTime(1991, 8, 25))
        };
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;

namespace AutoOpenXmlTest.Cases.ExportOrderedProperties
{
    [ExportWorkSheet(Variables.WorksheetName)]
    internal class ModelOrderedProperties
    {
        [ExportColumn(Variables.IdLabel, 2, ColumnTypes.Int)]
        public int Id { get; set; }
        [ExportColumn(Variables.NameLabel, 1, ColumnTypes.Text)]
        public string Name { get; set; }
        [ExportColumn(Variables.HeightLabel, 4, ColumnTypes.Decimal)]
        public decimal Height { get; set; }
        [ExportColumn(Variables.BirthDateLabel, 3, ColumnTypes.DateTime)]
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

    internal static class Variables
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

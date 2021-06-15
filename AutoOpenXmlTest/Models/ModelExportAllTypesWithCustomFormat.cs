using AutoOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet(VariablesModelExportAllTypesWithCustomFormat.WorksheetName)]
    public class ModelExportAllTypesWithCustomFormat
    {
        [ExportColumn(
            VariablesModelExportAllTypesWithCustomFormat.IdFieldLabel,
            1,
            VariablesModelExportAllTypesWithCustomFormat.IdFieldFormat)]
        public int Id { get; set; }

        [ExportColumn(
            VariablesModelExportAllTypesWithCustomFormat.PointsFieldLabel,
            2,
            VariablesModelExportAllTypesWithCustomFormat.PointsFieldFormat)]
        public decimal Points { get; set; }

        [ExportColumn(
            VariablesModelExportAllTypesWithCustomFormat.DateFieldLabel,
            3,
            VariablesModelExportAllTypesWithCustomFormat.DateFieldFormat)]
        public DateTime Date { get; set; }

        public ModelExportAllTypesWithCustomFormat() {}

        public ModelExportAllTypesWithCustomFormat(int id, decimal points, DateTime date)
        {
            Id = id;
            Points = points;
            Date = date;
        }
    }

    public static class VariablesModelExportAllTypesWithCustomFormat
    {
        public const string WorksheetName = "AllTypesWithCustomFormat";
        public const string IdFieldLabel = "Código";
        public const string IdFieldFormat = "00000";
        public const string PointsFieldLabel = "Pontos";
        public const string PointsFieldFormat = "00000.00";
        public const string DateFieldLabel = "Data Emissão";
        public const string DateFieldFormat = "dd/mm/yyyy";

        public static readonly ModelExportAllTypesWithCustomFormat[] Data = {
            new ModelExportAllTypesWithCustomFormat(
                1,
                (decimal)10.25,
                new DateTime(2010, 2, 5, 4, 2, 3))
        };
    }
}

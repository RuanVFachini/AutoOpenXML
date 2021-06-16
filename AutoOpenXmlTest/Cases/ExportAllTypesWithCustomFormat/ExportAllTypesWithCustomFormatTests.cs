using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;
using System.Linq;

namespace AutoOpenXmlTest.Cases.ExportAllTypesWithCustomFormat
{
    public class ExportAllTypesWithCustomFormatTests
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportAllTypesWithCustomFormat()
        {
            var stream = new ExportManagerBuilder<ModelExportAllTypesWithCustomFormat>()
                            .Init()
                            .SetData(VariablesModelExportAllTypesWithCustomFormat.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelExportAllTypesWithCustomFormat.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(2, 1).Style.NumberFormat.Format.Should()
                .Be(VariablesModelExportAllTypesWithCustomFormat.IdFieldFormat);
            worksheet.Cell(2, 2).Style.NumberFormat.Format.Should()
                .Be(VariablesModelExportAllTypesWithCustomFormat.PointsFieldFormat);
            worksheet.Cell(2, 3).Style.NumberFormat.Format.Should()
                .Be(VariablesModelExportAllTypesWithCustomFormat.DateFieldFormat);
        }
    }
}

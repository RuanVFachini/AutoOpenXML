using System.Linq;
using AutoOpenXml;
using ClosedXML.Excel;
using NUnit.Framework;
using FluentAssertions;
using AutoOpenXmlTest.Models;

namespace AutoOpenXmlTest.Cases.ExportStringProperty
{
    public class ExportStringPropertyTest
    {
        [SetUp]
        public void Setup() {}

        [Test]
        public void ShouldExportExcelFileWithTextColumn()
        {
            var stream = new ExportManagerBuilder<ModelStringProperty>()
                            .Init()
                            .SetData(VariablesModelStringProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelStringProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 1).Value.ToString().Should().BeEquivalentTo(VariablesModelStringProperty.FieldName);
            worksheet.Cell(2, 1).Value.ToString().Should().BeEquivalentTo(VariablesModelStringProperty.Data[0].Name);
            worksheet.Cell(3, 1).Value.ToString().Should().BeEquivalentTo(VariablesModelStringProperty.Data[1].Name);
        }
    }
}
using System.Linq;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ExportDateTimeProperty
{
    public class ExportDateTimePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithDateTimeColumn()
        {
            var stream = new ExportManagerBuilder<ModelDateTimeProperty>()
                            .Init()
                            .SetData(VariablesModelDateTimeProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelDateTimeProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 4).Value.ToString().Should().BeEquivalentTo(VariablesModelDateTimeProperty.FieldName);
            worksheet.Cell(2, 4).Value.Should().BeEquivalentTo(VariablesModelDateTimeProperty.Data[0].BirthDay);
            worksheet.Cell(3, 4).Value.Should().BeEquivalentTo(VariablesModelDateTimeProperty.Data[1].BirthDay);
        }
    }
}

using System.Linq;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ExportDateTimeProperty
{
    public class ExportDateTimeNullablePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithDateTimeColumn()
        {
            var stream = new ExportManagerBuilder<ModelDateTimeNullableProperty>()
                            .Init()
                            .SetData(VariablesModelDateTimeNullableProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelDateTimeNullableProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 4).Value.ToString().Should()
                .BeEquivalentTo(VariablesModelDateTimeNullableProperty.FieldName);
            worksheet.Cell(2, 4).Value.Should()
                .BeEquivalentTo(VariablesModelDateTimeNullableProperty.Data[0].BirthDay);
            worksheet.Cell(3, 4).Value.Should()
                .BeEquivalentTo(VariablesModelDateTimeNullableProperty.Data[1].BirthDay);
            worksheet.Cell(4, 4).Value.Should()
                .BeEquivalentTo("");
        }
    }
}

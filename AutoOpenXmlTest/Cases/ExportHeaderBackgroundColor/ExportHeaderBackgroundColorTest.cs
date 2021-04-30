using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;
using System.Collections.Generic;

namespace AutoOpenXmlTest.Cases.ExportStringProperty
{
    public class ExportHeaderBackgroundColorTest
    {
        [SetUp]
        public void Setup() {}

        [Test]
        public void ShouldExportExcelFileWithHeaderBackgroundColor()
        {
            var data = new List<ModelHeaderBackgroundColor>()
            {
                new ModelHeaderBackgroundColor("Luiz", 32),
                new ModelHeaderBackgroundColor("Frodo", 11)
            };

            var stream = new ExportManagerBuilder<ModelHeaderBackgroundColor>()
                            .Init()
                            .SetData(data)
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet("new sheet", out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 1).Style.Fill.BackgroundColor.Color.Should().BeEquivalentTo(XLColor.Transparent.Color);
            worksheet.Cell(1, 2).Style.Fill.BackgroundColor.Color.Should().BeEquivalentTo(XLColor.FromArgb(0, 255, 0).Color);
            worksheet.Cell(2, 1).Style.Fill.BackgroundColor.Color.Should().BeEquivalentTo(XLColor.Transparent.Color);
            worksheet.Cell(2, 2).Style.Fill.BackgroundColor.Color.Should().BeEquivalentTo(XLColor.Transparent.Color);
        }
    }
}
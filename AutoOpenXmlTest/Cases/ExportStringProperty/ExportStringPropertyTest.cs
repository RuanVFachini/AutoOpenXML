using System;
using System.Collections.Generic;
using System.Linq;
using AutoOpenXml;
using ClosedXML.Excel;
using NUnit.Framework;
using FluentAssertions;

namespace AutoOpenXmlTest.Cases.ExportStringProperty
{
    public class ExportStringPropertyTest
    {
        [SetUp]
        public void Setup() {}

        [Test]
        public void ShouldExportExcelFileWithTextColumn()
        {
            var stream = new AutoOpenXmlManager<ModelStringProperty>()
                            .Init()
                            .SetData(Variables.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(Variables.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 1).Value.ToString().Should().BeEquivalentTo(Variables.FieldName);
            worksheet.Cell(2, 1).Value.ToString().Should().BeEquivalentTo(Variables.Data[0].Name);
            worksheet.Cell(3, 1).Value.ToString().Should().BeEquivalentTo(Variables.Data[1].Name);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ExportIntProperty
{
    public class ExportIntPropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithNumberColumn()
        {
            var stream = new ExportManager<ModelIntProperty>()
                            .Init()
                            .SetData(VariablesModelIntProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelIntProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 1).Value.ToString().Should().BeEquivalentTo(VariablesModelIntProperty.FieldName);
            worksheet.Cell(2, 1).Value.Should().BeEquivalentTo(VariablesModelIntProperty.Data[0].Age);
            worksheet.Cell(3, 1).Value.Should().BeEquivalentTo(VariablesModelIntProperty.Data[1].Age);
        }
    }
}

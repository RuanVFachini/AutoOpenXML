using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ExportDecilmalProperty
{
    public class ExportDecimalPropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithNumberColumn()
        {
            var stream = new ExportManager<ModelDecimalProperty>()
                            .Init()
                            .SetData(VariablesModelDecimalProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelDecimalProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 1).Value.ToString().Should().BeEquivalentTo(VariablesModelDecimalProperty.FieldName);
            worksheet.Cell(2, 1).Value.Should().BeEquivalentTo(VariablesModelDecimalProperty.Data[0].Salary);
            worksheet.Cell(3, 1).Value.Should().BeEquivalentTo(VariablesModelDecimalProperty.Data[1].Salary);
        }
    }
}

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
    public class ExportDecimalNullablePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithNumberColumn()
        {
            var stream = new ExportManagerBuilder<ModelDecimalNullableProperty>()
                            .Init()
                            .SetData(VariablesModelDecimalNullableProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelDecimalNullableProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 2).Value.ToString().Should()
                .BeEquivalentTo(VariablesModelDecimalNullableProperty.FieldName);
            worksheet.Cell(2, 2).Value.Should()
                .BeEquivalentTo(VariablesModelDecimalNullableProperty.Data[0].Salary);
            worksheet.Cell(3, 2).Value.Should()
                .BeEquivalentTo(VariablesModelDecimalNullableProperty.Data[1].Salary);
            worksheet.Cell(4, 2).Value.Should().Be("");
        }
    }
}

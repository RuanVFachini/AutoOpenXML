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
    public class ExportIntNullablePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithNumberColumn()
        {
            var stream = new ExportManagerBuilder<ModelIntNullableProperty>()
                            .Init()
                            .SetData(VariablesModelIntNullableProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelIntNullableProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 3).Value.ToString().Should()
                .BeEquivalentTo(VariablesModelIntNullableProperty.FieldName);
            worksheet.Cell(2, 3).Value.Should()
                .BeEquivalentTo(VariablesModelIntNullableProperty.Data[0].Age);
            worksheet.Cell(3, 3).Value.Should()
                .BeEquivalentTo(VariablesModelIntNullableProperty.Data[1].Age);
        }
    }
}

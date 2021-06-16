using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ExportDateTimeProperty
{
    public class ExportBoolNullablePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithBooleanColumn()
        {
            var stream = new ExportManagerBuilder<ModelBoolNullableProperty>()
                            .Init()
                            .SetData(VariablesModelBoolNullableProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelBoolNullableProperty.WorksheetName, out var worksheet);

            StreamTestFile.SaveExternalFile(stream);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 5).Value.ToString().Should()
                .BeEquivalentTo(VariablesModelBoolNullableProperty.FieldName);
            worksheet.Cell(2, 5).Value.Should()
                .BeEquivalentTo(VariablesModelBoolNullableProperty.Data[0].Active);
            worksheet.Cell(3, 5).Value.Should()
                .BeEquivalentTo(VariablesModelBoolNullableProperty.Data[1].Active);
            worksheet.Cell(4, 5).Value.Should().Be("");
        }
    }
}

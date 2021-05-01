using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
    public class ExportBoolPropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithBooleanColumn()
        {
            var stream = new ExportManagerBuilder<ModelBoolProperty>()
                            .Init()
                            .SetData(VariablesModelBoolProperty.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelBoolProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 5).Value.ToString().Should().BeEquivalentTo(VariablesModelBoolProperty.FieldName);
            worksheet.Cell(2, 5).Value.Should().BeEquivalentTo(VariablesModelBoolProperty.Data[0].Active);
            worksheet.Cell(3, 5).Value.Should().BeEquivalentTo(VariablesModelBoolProperty.Data[1].Active);
        }
    }
}

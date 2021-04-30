using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ExportOrderedProperties
{
    public class ExportOrderedPropertiesTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithNumberColumn()
        {
            var stream = new ExportManagerBuilder<ModelOrderedProperties>()
                            .Init()
                            .SetData(VariablesModelOrderedProperties.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelOrderedProperties.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 1).Value.ToString().Should().BeEquivalentTo(VariablesModelOrderedProperties.NameLabel);
            worksheet.Cell(1, 2).Value.ToString().Should().BeEquivalentTo(VariablesModelOrderedProperties.IdLabel);
            worksheet.Cell(1, 3).Value.ToString().Should().BeEquivalentTo(VariablesModelOrderedProperties.BirthDateLabel);
            worksheet.Cell(1, 4).Value.ToString().Should().BeEquivalentTo(VariablesModelOrderedProperties.HeightLabel);
        }
    }
}

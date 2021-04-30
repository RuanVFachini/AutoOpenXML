using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AutoOpenXmlTest.Cases.ExportAllTypesWithCustomFormat
{
    public class ExportAllTypesWithCustomFormatTests
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportAllTypesWithCustomFormat()
        {
            var stream = new ExportManagerBuilder<ModelExportAllTypesWithCustomFormat>()
                            .Init()
                            .SetData(VariablesModelExportAllTypesWithCustomFormat.Data.ToList())
                            .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(VariablesModelDateTimeProperty.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            using (FileStream fs = new FileStream("c:\\temp\\teste.xls", FileMode.OpenOrCreate))
            {
                stream.WriteTo(fs);
            }
        }
    }
}

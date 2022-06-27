using AutoOpenXml;
using AutoOpenXmlTest.Models;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.MultiExport
{
    public class MultiExport
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportManyWorksheet()
        {
            var stream = new MultiExportManagerBuilder()
                .Init()
                .SetData(VariablesModelDateTimeProperty.Data)
                .SetData(VariablesModelDecimalProperty.Data)
                .StartExportProcess();

            var workbook = new XLWorkbook(stream);
            
            workbook.TryGetWorksheet(VariablesModelDateTimeProperty.WorksheetName, out var worksheet1);
            worksheet1.Should().NotBeNull();
            
            workbook.TryGetWorksheet(VariablesModelDecimalProperty.WorksheetName, out var worksheet2);
            worksheet2.Should().NotBeNull();
        }
    }
}
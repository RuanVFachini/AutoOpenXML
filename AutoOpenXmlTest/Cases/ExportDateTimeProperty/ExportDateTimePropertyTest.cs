﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using AutoOpenXml;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ExportDateTimeProperty
{
    public class ExportDateTimePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithNumberColumn()
        {
            var stream = new AutoOpenXmlManager<ModelDateTimeProperty>()
                            .Init()
                            .SetData(Variables.Data.ToList())
                            .StartProcess();

            var workbook = new XLWorkbook(stream);
            workbook.TryGetWorksheet(Variables.WorksheetName, out var worksheet);

            worksheet.Should().NotBeNull();

            worksheet.Cell(1, 1).Value.ToString().Should().BeEquivalentTo(Variables.FieldName);
            worksheet.Cell(2, 1).Value.Should().BeEquivalentTo(Variables.Data[0].BirthDay);
            worksheet.Cell(3, 1).Value.Should().BeEquivalentTo(Variables.Data[1].BirthDay);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ImportDecimalProperty
{
    public class ImportDecimalProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportDecimalFieldFromNumberColumn()
        {
            var result = new ImportManagerBuilder<ModelDecimalProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelDecimalProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].Salary.Should().Be((decimal) 12233.45);
            result[1].Salary.Should().Be(0);

            StreamTestFile.CloseFile();
        }
    }
}

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
    public class ImportDecimalNullableProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportDecimalFieldFromNumberColumn()
        {
            var result = new ImportManagerBuilder<ModelDecimalNullableProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), "Planilha1")
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(1);
            result[0].Name.Should().BeNull();
            result[0].Salary.Should().Be((decimal) 12233.45);
        }
    }
}

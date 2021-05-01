using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ImportIntProperty
{
    public class ImportIntNullableProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportIntFieldFromNumberColumn()
        {
            var result = new ImportManagerBuilder<ModelIntNullableProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), "Planilha1")
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(1);
            result[0].Name.Should().BeNull();
            result[0].Age.Should().Be(18);

            StreamTestFile.CloseFile();
        }
    }
}

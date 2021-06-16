using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ImportStringProperty
{
    public class ImportStringProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportIntFieldFromNumberColumn()
        {
            var result = new ImportManagerBuilder<ModelStringProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelStringProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].Name.Should().Be("Ruan");
            result[1].Name.Should().Be("");

            StreamTestFile.CloseFile();
        }
    }
}

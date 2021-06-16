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
    public class ImportBoolNullableProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportIntFieldFromBooleanColumn()
        {
            var result = new ImportManagerBuilder<ModelBoolNullableProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelBoolNullableProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].Active.Should().BeTrue();
            result[1].Active.Should().BeNull();

            StreamTestFile.CloseFile();
        }
    }
}

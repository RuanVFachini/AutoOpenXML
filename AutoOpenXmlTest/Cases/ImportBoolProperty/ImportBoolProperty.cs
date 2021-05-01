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
    public class ImportBoolProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportIntFieldFromBooleanColumn()
        {
            var result = new ImportManagerBuilder<ModelBoolProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), "Planilha1")
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].Name.Should().BeNull();
            result[0].Active.Should().BeTrue();

            StreamTestFile.CloseFile();
        }
    }
}

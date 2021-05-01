using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using FluentAssertions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace AutoOpenXmlTest.Cases.ImportDateTimeProperty
{
    public class ImportDateTimeNullablePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportDateTimeFieldFromDateColumn()
        {
            var result = new ImportManagerBuilder<ModelDateTimeNullableProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), "Planilha1")
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(1);
            result[0].Name.Should().BeNull();
            result[0].BirthDay.Should().BeSameDateAs(new DateTime(1991, 08, 28));
        }
    }
}

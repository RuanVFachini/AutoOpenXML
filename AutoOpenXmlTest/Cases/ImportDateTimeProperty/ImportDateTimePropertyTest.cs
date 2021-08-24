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
    public class ImportDateTimePropertyTest
    {
        [SetUp]
        public void Setup() { }

        [Test]
        [Parallelizable(ParallelScope.None)]
        public void ShouldImportDateTimeFieldFromDateColumn()
        {
            var result = new ImportManagerBuilder<ModelDateTimeProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelDateTimeProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].BirthDay.Should().BeSameDateAs(new DateTime(1991, 08, 28));
            result[1].BirthDay.Should().BeSameDateAs(new DateTime(2021, 1, 1));

            StreamTestFile.CloseFile();
        }
    }
}

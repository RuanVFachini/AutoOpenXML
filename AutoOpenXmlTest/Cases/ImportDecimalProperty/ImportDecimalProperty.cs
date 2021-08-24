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
        [Parallelizable(ParallelScope.None)]
        public void ShouldImportDecimalFieldFromNumberColumn()
        {
            var result = new ImportManagerBuilder<ModelDecimalProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelDecimalProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].Salary.Should().Be(12233.45M);
            result[1].Salary.Should().Be(125.24M);

            StreamTestFile.CloseFile();
        }
    }
}

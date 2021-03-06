﻿using System;
using System.Collections.Generic;
using System.Text;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ImportIntProperty
{
    public class ImportIntProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldImportIntFieldFromNumberColumn()
        {
            var result = new ImportManagerBuilder<ModelIntProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelIntProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].Age.Should().Be(18);
            result[1].Age.Should().Be(0);

            StreamTestFile.CloseFile();
        }
    }
}

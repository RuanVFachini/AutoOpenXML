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
        [Parallelizable(ParallelScope.None)]
        public void ShouldImportIntFieldFromBooleanColumn()
        {
            var result = new ImportManagerBuilder<ModelBoolProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelBoolProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(2);
            result[0].Active.Should().BeTrue();
            result[1].Active.Should().BeFalse();

            StreamTestFile.CloseFile();
        }
    }
}

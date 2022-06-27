using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using FluentAssertions;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.ImportFormulaProperty
{
    public class ImportFormulaProperty
    {
        [SetUp]
        public void Setup() { }

        [Test]
        [Parallelizable(ParallelScope.None)]
        public void ShouldImportFormulaFieldFromNumberColumn()
        {
            var result = new ImportManagerBuilder<ModelFormulaProperty>()
                .OpenFile(StreamTestFile.GetStreamTestFile(), VariablesModelFormulaProperty.WorksheetName)
                .Init()
                .StartImportProcess();

            result.Should().NotBeNullOrEmpty();
            result.Count.Should().Be(4);
            result[0].Formula.Should().Be(5);
            result[1].Formula.Should().Be(8);
            result[2].Formula.Should().Be(13);
            result[3].Formula.Should().Be(6);

            StreamTestFile.CloseFile();
        }
    }
}

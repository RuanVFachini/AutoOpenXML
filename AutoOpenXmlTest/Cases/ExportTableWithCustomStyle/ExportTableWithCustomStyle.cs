using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using NUnit.Framework;
using System.Linq;

namespace AutoOpenXmlTest.Cases.ExportTableWithDefaultStyle
{
    public class ExportTableWithCustomStyle
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithDefaultTableStyle()
        {
            var stream = new ExportManagerBuilder<ModelTableWithCustomTableStyle>()
                            .Init()
                            .SetData(VariablesModelTableWithCustomTableStyle.Data.ToList())
                            .StartExportProcess();

            StreamTestFile.SaveExternalFile(stream);
        }
    }
}

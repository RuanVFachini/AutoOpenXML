using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using NUnit.Framework;
using System.Linq;

namespace AutoOpenXmlTest.Cases.ExportTableWithDefaultStyle
{
    public class ExportTableWithoutStyle
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExportExcelFileWithDefaultTableStyle()
        {
            var stream = new ExportManagerBuilder<ModelTableWithOutTableStyle>()
                            .Init()
                            .SetData(VariablesModelTableWithOutTableStyle.Data.ToList())
                            .StartExportProcess();

            StreamTestFile.SaveExternalFile(stream);
        }
    }
}

using AutoOpenXml;
using AutoOpenXmlTest.Models;
using AutoOpenXmlTest.Utils;
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

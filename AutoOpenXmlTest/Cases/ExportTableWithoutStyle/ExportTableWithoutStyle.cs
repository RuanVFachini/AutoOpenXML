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

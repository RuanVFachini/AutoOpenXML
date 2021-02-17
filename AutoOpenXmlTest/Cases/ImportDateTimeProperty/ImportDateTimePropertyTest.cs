using AutoOpenXml;
using AutoOpenXmlTest.Models;
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
        public void ShouldImportDateTimeFieldFromDateColumn()
        {
            var file = File.Open(@"C:\WorkSpace\AutoOpenXML\AutoOpenXmlTest\Resources\input.xlsx", FileMode.Open);
            var stream = new MemoryStream();
            file.CopyTo(stream);

            var result = new ImportManager<ModelDateTimeProperty>()
                .OpenFile(stream, "Planilha1")
                .Init()
                .StartImportProcess();


        }
    }
}

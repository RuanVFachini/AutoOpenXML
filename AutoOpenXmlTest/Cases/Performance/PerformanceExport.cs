using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using AutoOpenXml;
using AutoOpenXmlTest.Models;
using NUnit.Framework;

namespace AutoOpenXmlTest.Cases.Performance
{
    public class PerformanceExport_5k
    {
        [SetUp]
        public void Setup() { }

        [Test]
        public void ShouldExport()
        {
            //Arrange
            var data = MockRamdomData();

            var watch = new Stopwatch();

            watch.Start();
            var stream = new ExportManagerBuilder<ModelOrderedProperties>()
                .Init()
                .SetData(data.ToList())
                .StartExportProcess();

            watch.Stop();
            var time = watch.Elapsed.Seconds;
        }

        private IEnumerable<ModelOrderedProperties> MockRamdomData()
        {
            for (int i = 1; i < 5000; i++)
            {
                yield return new ModelOrderedProperties()
                {
                    Height = (decimal) (new Random().Next(10) * 0.75 * i),
                    Id = i,
                    Name = "Super Mario World",
                    BirthDate = new DateTime(1990, 11, 21)
                };
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace AutoOpenXmlTest.Utils
{
    public static class StreamTestFile
    {
        public static MemoryStream GetStreamTestFile()
        {
            var file = File.Open(@"C:\Users\ruan.fachini\source\repos\ObjectToExcelFile\AutoOpenXmlTest\Resources\input.xlsx", FileMode.Open);
            var stream = new MemoryStream();
            file.CopyTo(stream);
            return stream;
        }
    }
}

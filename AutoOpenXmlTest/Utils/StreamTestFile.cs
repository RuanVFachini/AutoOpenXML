﻿using System.IO;

namespace AutoOpenXmlTest.Utils
{
    public static class StreamTestFile
    {
        public static FileStream FileStream;

        public static MemoryStream GetStreamTestFile()
        {
            FileStream = File.Open(@"C:\workspace\projets - Pessoais\AutoOpenXML\AutoOpenXmlTest\Resources\input.xlsx", FileMode.Open);
            var stream = new MemoryStream();
            FileStream.CopyTo(stream);
            return stream;
        }

        public static void CloseFile()
        {
            FileStream.Close();
        }
    }
}

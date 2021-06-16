using System.IO;

namespace AutoOpenXmlTest.Utils
{
    public static class StreamTestFile
    {
        public static FileStream FileStream;

        public static MemoryStream GetStreamTestFile()
        {
            FileStream = File.Open(@$"{Directory.GetCurrentDirectory()}\Resources\input.xlsx", FileMode.Open);
            var stream = new MemoryStream();
            FileStream.CopyTo(stream);
            return stream;
        }

        public static void SaveExternalFile(MemoryStream stream)
        {
            FileStream = File.Open(@$"{Directory.GetCurrentDirectory()}\Resources\output.xlsx", FileMode.OpenOrCreate);
            FileStream.Write(stream.ToArray());
            FileStream.Close();
        }

        public static void CloseFile()
        {
            FileStream.Close();
        }
    }
}

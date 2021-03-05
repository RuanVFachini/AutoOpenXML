using AutoOpenXml;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet("new sheet")]
    public class ModelHeaderBackgroundColor
    {
        [ExportColumn("name label", 1)]
        public string Name { get; set; }

        [ExportColumn("age label", 2)]
        [ExportColumnHeaderBackgoundColor(0, 255, 0)]
        public int Age { get; set; }

        public ModelHeaderBackgroundColor()
        { }

        public ModelHeaderBackgroundColor(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }
}
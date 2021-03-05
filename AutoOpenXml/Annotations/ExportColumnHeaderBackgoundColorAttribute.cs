using System.Drawing;

namespace AutoOpenXml
{
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ExportColumnHeaderBackgoundColorAttribute : System.Attribute
    {
        public readonly Color HeaderBackgroundColor;

        public ExportColumnHeaderBackgoundColorAttribute(int red = 255, int green = 255, int blue = 255)
        {
            HeaderBackgroundColor = Color.FromArgb(red, green, blue);
        }
    }
}
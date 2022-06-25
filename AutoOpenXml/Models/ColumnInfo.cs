using System.Drawing;
using Fasterflect;

namespace AutoOpenXml.Models
{
    internal class ColumnInfo<T>
    {
        public string Label { get; set; }
        public string Name { get; set; }
        public string Mask { get; set; }
        public int Index { get; set; }
        public TypesEnum Type { get; set; }
        public MemberGetter GetValueFunc { get; set; }
        public MemberSetter SetValueFunc { get; set; }
        public Color HeaderBackgroundColor { get; set; }
    }
}

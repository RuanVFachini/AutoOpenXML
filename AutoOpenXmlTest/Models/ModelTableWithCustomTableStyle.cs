using AutoOpenXml;
using System.Collections.Generic;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet("TableWithCustomTableStyle", "TableStyleDark7")]
    public class ModelTableWithCustomTableStyle
    {
        [ExportColumn("Id", 1)]
        public int Id { get; set; }

        [ExportColumn("Name", 2)]
        public string Name { get; set; }

        [ExportColumn("Last Name", 3)]
        public string LastName { get; set; }

        [ExportColumn("Age", 4)]
        public int Age { get; set; }

        public ModelTableWithCustomTableStyle() { }

        public ModelTableWithCustomTableStyle(int id, string name, string lastName, int age)
        {
            Id = id;
            Name = name;
            LastName = lastName;
            Age = age;
        }
    }

    public static class VariablesModelTableWithCustomTableStyle
    {
        public static List<ModelTableWithCustomTableStyle> Data = new List<ModelTableWithCustomTableStyle>()
        {
            new ModelTableWithCustomTableStyle(1, "Carlos", "Fancisco", 17),
            new ModelTableWithCustomTableStyle(2, "André", "Silva", 36),
            new ModelTableWithCustomTableStyle(3, "Aroldo", "Santos", 45),
            new ModelTableWithCustomTableStyle(4, "Marina", "Fernander", 73),
            new ModelTableWithCustomTableStyle(5, "Gabriel", "Rover", 13),
            new ModelTableWithCustomTableStyle(6, "Wilian", "Fachini", 12),
            new ModelTableWithCustomTableStyle(7, "Mario", "Muller", 82),
            new ModelTableWithCustomTableStyle(8, "Luide", "Barbosa", 23),
            new ModelTableWithCustomTableStyle(9, "Eduardo", "Nascimento", 32)
        };
    }
}

using AutoOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace AutoOpenXmlTest.Models
{
    [ExportWorkSheet("TableWithOutTableStyle")]
    public class ModelTableWithOutTableStyle
    {
        [ExportColumn("Id", 1)]
        public int Id { get; set; }

        [ExportColumn("Name", 2)]
        public string Name { get; set; }

        [ExportColumn("Last Name", 3)]
        public string LastName { get; set; }

        [ExportColumn("Age", 4)]
        public int Age { get; set; }

        public ModelTableWithOutTableStyle() { }

        public ModelTableWithOutTableStyle(int id, string name, string lastName, int age)
        {
            Id = id;
            Name = name;
            LastName = lastName;
            Age = age;
        }
    }

    public static class VariablesModelTableWithOutTableStyle
    {
        public static List<ModelTableWithOutTableStyle> Data = new List<ModelTableWithOutTableStyle>()
        {
            new ModelTableWithOutTableStyle(1, "Carlos", "Fancisco", 17),
            new ModelTableWithOutTableStyle(2, "André", "Silva", 36),
            new ModelTableWithOutTableStyle(3, "Aroldo", "Santos", 45),
            new ModelTableWithOutTableStyle(4, "Marina", "Fernander", 73),
            new ModelTableWithOutTableStyle(5, "Gabriel", "Rover", 13),
            new ModelTableWithOutTableStyle(6, "Wilian", "Fachini", 12),
            new ModelTableWithOutTableStyle(7, "Mario", "Muller", 82),
            new ModelTableWithOutTableStyle(8, "Luide", "Barbosa", 23),
            new ModelTableWithOutTableStyle(9, "Eduardo", "Nascimento", 32)
        };
    }
}

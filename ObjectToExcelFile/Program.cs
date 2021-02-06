using System;
using System.Collections.Generic;
using System.IO;
using AutoOpenXml;

namespace ObjectToExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
           

        }
    }

    [ExportWorkSheet("Nova Aba")]
    class Teste
    {
        [ExportColumn("Nome da Pessoa", 1, ColumnTypes.Text)]
        public string Nome { get; set; }
        [ExportColumn("Idade da Pessoa", 0, ColumnTypes.Number)]
        public int Idade { get; set; }

        public Teste()
        { }

        public Teste(string nome, int idade)
        {
            Nome = nome;
            Idade = idade;
        }
    }
}

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
            var dados = new List<Teste>()
            {
                new Teste("Joao", "15"),
                new Teste("Carlos", "32"),
            };

            var stream = new AutoOpenXmlManager<Teste>()
                            .Init()
                            .SetData(dados)
                            .StartProcess();

            FileStream file = new FileStream("c:\\temp\\export.xlsx", FileMode.Create, FileAccess.Write);
            stream.WriteTo(file);
            file.Close();
            stream.Close();

        }
    }

    [ExportWorkSheet("Nova Aba")]
    class Teste
    {
        [ExportColumn("Nome da Pessoa", 1)]
        public string Nome { get; set; }
        [ExportColumn("Idade da Pessoa", 0)]
        public string Idade { get; set; }

        public Teste()
        { }

        public Teste(string nome, string idade)
        {
            Nome = nome;
            Idade = idade;
        }
    }
}

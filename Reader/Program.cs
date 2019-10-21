using System;
using ClosedXML.Excel;

namespace Reader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("".PadRight(60, '-'));
            Console.WriteLine("Id" + "Nome".PadLeft(16, ' ') + "Endereco".PadLeft(19, ' ') + "Nascimento".PadLeft(23, ' '));

            XLWorkbook cliente = new XLWorkbook(@"C: \Users\Dell\Desktop\arquivosTeste\cliente.xlsx");
            var plan1 = cliente.Worksheet(1);

            int linha = 2;
        
            while (true)
            {
                string id = plan1.Cell("A" + linha.ToString()).Value.ToString();
                string nome = plan1.Cell("B" + linha.ToString()).Value.ToString();
                string endereco = plan1.Cell("C" + linha.ToString()).Value.ToString();
                string nascimento = plan1.Cell("D" + linha.ToString()).Value.ToString();

                if (string.IsNullOrEmpty(id) && string.IsNullOrEmpty(nome) && 
                    string.IsNullOrEmpty(endereco) && string.IsNullOrEmpty(nascimento)) break;

                Console.WriteLine(id + nome.PadLeft(16, ' ') + endereco.PadLeft(19, ' ') + nascimento.PadLeft(23, ' ').ToString());
                ++linha;
            }

            Console.ReadKey();
        }
    }
}

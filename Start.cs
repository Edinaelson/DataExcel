using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataExcel
{
    internal class Start
    {
        public static void Startting() 
        {
            var xlsx = new XLWorkbook(@"C:\Users\sedin\Desktop\Teste.xlsx");
            var planilha = xlsx.Worksheets.First(x => x.Name == "Planilha1");
            var totalL = 40;
            int linhaAtual = 45;

            Console.Title = "Excell VENDAS";
            //Console.BackgroundColor = ConsoleColor.Red;
            int quantMemoriaDisponivel = 0;
            int quantMemoriaNaoDisponivel = 0;
            float maiorValorCompra = 0;

            //inicializar variaveis
            int result = 0;
            float resultCompra = 0f;

            for (int i = 2; i < totalL; i++)
            {
                var marca = planilha.Cell($"A{linhaAtual}").Value.ToString();
                var capacidade = planilha.Cell($"B{linhaAtual}").Value.ToString();
                var frequencia = planilha.Cell($"C{linhaAtual}").Value.ToString();
                var valorCompra = planilha.Cell($"D{linhaAtual}").Value.ToString();
                var valorVenda = planilha.Cell($"E{linhaAtual}").Value.ToString();
                var lucro = planilha.Cell($"F{linhaAtual}").Value.ToString();
                var disponivel = planilha.Cell($"G{linhaAtual}").Value.ToString();

                if (linhaAtual == 45)
                {
                    Console.BackgroundColor = ConsoleColor.Blue;
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else
                {
                    Console.ResetColor();
                }
                if (linhaAtual > 45)
                {
                    if (WhatType.IsFloat(valorCompra) == true)
                    {
                        resultCompra = float.Parse(valorCompra);
                    }

                    if (resultCompra > maiorValorCompra)
                    {
                        maiorValorCompra = resultCompra;
                    }

                    if (WhatType.IsInt(disponivel) == true)
                    {
                        result = int.Parse(disponivel);
                    }

                    if (result == 1)
                    {
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.BackgroundColor = ConsoleColor.DarkMagenta;
                        quantMemoriaDisponivel += 1;   
                    }
                    if (result == 0)
                    {
                        //Console.ResetColor();
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.BackgroundColor = ConsoleColor.DarkBlue;
                        quantMemoriaNaoDisponivel += 1;
                    }
                    //Console.ResetColor();
                }

               linhaAtual++;
               Console.WriteLine($"{marca,-12} | {capacidade,-13} | {frequencia,-13} | {valorCompra,-12} | {valorVenda,-11} | {lucro,-6} | {disponivel,-10}");

            }
            Console.ResetColor();
            Console.WriteLine("-----------------------------------------|");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Quantidades vendidas: {quantMemoriaDisponivel}");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"Quantidade a chegar: {quantMemoriaNaoDisponivel}");
            Console.ResetColor();
            Console.WriteLine($"Compra com maior valor: {maiorValorCompra}");
            Console.WriteLine("-----------------------------------------|");
        }
    }
}

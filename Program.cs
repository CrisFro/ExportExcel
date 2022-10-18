using System;
using System.Diagnostics;
using ClosedXML.Excel;

namespace ExportExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            using(var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Planilha1");
                worksheet.Cell("A1").Value = "Olá mundo";
                worksheet.Cell("A2").Value = 1;
                worksheet.Cell("A3").Value = 2;
                worksheet.Cell("A4").Value = 3;

                //utilizando fórmula
                worksheet.Cell("A5").FormulaA1 = "=SUM(A2:A4)";

                //utilizando imagens
                var caminhoImagem = @"C:\Users\Cristiane\Desktop\Exercícios\ExportExcel\img\csharp.png";
                worksheet.AddPicture(caminhoImagem).MoveTo(worksheet.Cell("A10")).Scale(0.5);

                //utilizando bordas
                worksheet.Cell("A1").Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                worksheet.Cell("A1").Style.Border.BottomBorderColor = XLColor.Red; 
                worksheet.Cell("A1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                worksheet.Cell("A1").Style.Border.LeftBorderColor = XLColor.Red;
                worksheet.Cell("A1").Style.Border.RightBorder = XLBorderStyleValues.Thick;
                worksheet.Cell("A1").Style.Border.RightBorderColor = XLColor.Red;
                worksheet.Cell("A1").Style.Border.TopBorder = XLBorderStyleValues.Thick;
                worksheet.Cell("A1").Style.Border.TopBorderColor = XLColor.Red;

                //Calculando com fórmula
                Console.WriteLine("Valor da soma: {0}", worksheet.Cell("A5").Value);

                workbook.SaveAs(@"c:\Users\Cristiane\Desktop\Exercícios\exportExcel.xlsx");
            }

            Process.Start(new ProcessStartInfo(@"C:\Users\Cristiane\Desktop\Exercícios\ExportExcel\exportExcel.xlsx") { UseShellExecute = true });
        }
    }
}

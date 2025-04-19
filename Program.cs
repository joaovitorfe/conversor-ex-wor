using System;
using System.IO;
using OfficeOpenXml;
using Xceed.Words.NET;

namespace ExcelToWordConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string excelPath = "dados.xlsx";
            string wordPath = "resultado.docx";

            if (!File.Exists(excelPath))
            {
                Console.WriteLine("Arquivo Excel não encontrado.");
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                var doc = DocX.Create(wordPath);
                doc.InsertParagraph("Relatório Gerado do Excel:
");

                for (int row = 2; row <= rowCount; row++)
                {
                    string linha = "";
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        linha += worksheet.Cells[row, col].Text + " ";
                    }
                    doc.InsertParagraph(linha);
                }

                doc.Save();
                Console.WriteLine("Documento Word gerado com sucesso!");
            }
        }
    }
}

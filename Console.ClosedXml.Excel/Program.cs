using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("INICIO");
            List<Estudiante> objListEstudaintes = new List<Estudiante>()
            {
                new Estudiante()
                {
                    Id = "30789573",
                    CodEstudiante = "00000078731816",
                    Nombres = "ARIAS ESCOBAR JESUS ENRIQUE"
                },
                new Estudiante()
                {
                    Id = "31380891",
                    CodEstudiante = "00000078737991",
                    Nombres = "ASCONA REYES JUAN EDGAR"
                },
                new Estudiante()
                {
                    Id = "30942964",
                    CodEstudiante = "00000081250381",
                    Nombres = "CANALES AVENDAÑO CELESTE SARAI"
                },
                new Estudiante()
                {
                    Id = "30942937",
                    CodEstudiante = "00000078731619",
                    Nombres = "PISCO MEJIA ADERLY THIAGO"
                }
            };

            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.Worksheets.Add("Sample Sheet");

                // Generamos la cabecera parte 1: datos sin notas
                worksheet.Cell("A1").Value = "ID";
                worksheet.Range("A1:A3").Merge();
                worksheet.Cell("B1").Value = "CodEstudiante";
                worksheet.Range("B1:B3").Merge();
                worksheet.Cell("C1").Value = "Nombres";
                worksheet.Range("C1:C3").Merge();

                // Generamos la cabecera parte 2: datos con notas
                worksheet.Cell("D1").Value = "C01";
                worksheet.Range("D1:D2").Merge();
                worksheet.Cell("D3").Value = "CAC";
                worksheet.Cell("E1").Value = "C02";
                worksheet.Range("E1:E2").Merge();
                worksheet.Cell("E3").Value = "CAC";
                worksheet.Cell("F1").Value = "C03";
                worksheet.Range("F1:F2").Merge();
                worksheet.Cell("F3").Value = "CAC";
                worksheet.Cell("G1").Value = "C04";
                worksheet.Range("G1:G2").Merge();
                worksheet.Cell("G3").Value = "CAC";

                worksheet.Cell("H1").Value = "CAA";
                worksheet.Range("H1:H3").Merge();

                worksheet.Cell("I1").Value = "Conclusión descriptiva de final del periodo";
                worksheet.Range("I1:I3").Merge();
                

                //---------------- le damos el formato a la cabacera -------------------------
                var rango = worksheet.Range("A1:I3");
                rango.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rango.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rango.Style.Border.BottomBorderColor = XLColor.Black;             
                rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rango.Style.Font.Bold = true;
                rango.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 180, 180);
                //rango.Style.Font.FontName = "Courier New";                
                //rango.Style.Alignment.WrapText = true;                

                //------------------- Generar la tabla de estudiantes con sus notas ----------------------
                int nroRow = 4;
                foreach (var estudiante in objListEstudaintes)
                {
                    worksheet.Cell(nroRow, 1).Value = estudiante.Id;
                    worksheet.Cell(nroRow, 2).Style.NumberFormat.Format = "@";
                    worksheet.Cell(nroRow, 2).Value = estudiante.CodEstudiante.ToString();
                    worksheet.Cell(nroRow, 3).Value = estudiante.Nombres;
                    nroRow++;
                }

                //------------------------ Data Validation -------------------------------------------------------
                var options = new List<string>() { "AD","A","B","C" };
                var validOptions = $"\"{String.Join(",", options)}\"";
                nroRow = 4;
                foreach (var estudiante in objListEstudaintes)
                {
                    worksheet.Cell(nroRow, 4).DataValidation.List(validOptions, true);
                    worksheet.Cell(nroRow, 5).DataValidation.List(validOptions, true);
                    worksheet.Cell(nroRow, 6).DataValidation.List(validOptions, true);
                    worksheet.Cell(nroRow, 7).DataValidation.List(validOptions, true);
                    worksheet.Cell(nroRow, 8).DataValidation.List(validOptions, true);
                    nroRow++;
                }                               

                //------------------------- Leyenda --------------------------------------------------------
                worksheet.Cell(nroRow + 2, 2).Value = "LEYENDA";
                worksheet.Cell(nroRow + 3, 3).Value = new[] 
                {
                    "01 = Se comunica oralmente en su lengua materna",
                    "02 = Lee diversos tipos de texto en su lengua materna",
                    "03 = Escribe diversos tipos de textos en su lengua materna",
                    "04 = Crea Proyectos desde los lenguajes del arte"
                };

                //---------------------------- Le damos formato a la tabla general --------------------------
                worksheet.Columns(1, 9).AdjustToContents(); // Ajustamos el ancho de las columnas para que se muestren todos los contenidos
                //worksheet.Column(9).AdjustToContents(); // Ajustamos el ancho de una columna de acuerdo a su contenido
                worksheet.Column(9).Width = 100;
                worksheet.Columns(4, 8).Width = 6;

                Console.WriteLine("ANTES DE GUARDAR");
                workbook.SaveAs("HelloWorld.xlsx");
                Console.WriteLine("DESPUES DE GUARDAR");
            }
            Console.WriteLine("FIN");
            Console.ReadLine();
        }

       public class Estudiante
        {
            public string Id { get; set; }
            public string CodEstudiante { get; set; }
            public string Nombres { get; set; }

        }      
    }
}

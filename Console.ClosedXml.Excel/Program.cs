using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Console.ClosedXml.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "ID";
                worksheet.Range("A1:A3").Merge();
                worksheet.Cell("B1").Value = "CodEstudiante";
                worksheet.Range("B1:B3").Merge();
                worksheet.Cell("C1").Value = "Nombres";
                worksheet.Range("C1:C3").Merge();

                worksheet.Range("A1:C3").Style
                    .Border.SetBottomBorder(XLBorderStyleValues.Thick)
                    .Border.SetBottomBorderColor(XLColor.Red)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Font.SetBold(true)
                    .Fill.SetBackgroundColor(XLColor.Gray);


                workbook.SaveAs("HelloWorld.xlsx");
            }
        }
    }
}

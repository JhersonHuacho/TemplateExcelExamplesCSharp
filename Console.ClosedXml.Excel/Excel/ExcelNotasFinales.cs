using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel
{
    public class ExcelNotasFinales
    {
        public string GetPlantilla()
        {
            //---------------- crear u obtener la plantilla que con la que se va a trabajar --------------
            string _serverTemp;
            string _FileName;
            string plantilla;
            string pathPlantilla;

            Random random = new Random();
            int numRandom = random.Next(1000, 99999);

            //FileName = $"RegNotasFinales_{entity.Perfil.IdInstitucion + entity.Perfil.Anexo}" +
            //           $"_{entity.DisenioCurricular}" +
            //           $"_{entity.Perfil.IdNivelInstitucion.Trim() + entity.AnioAcademico + entity.GradoIe + entity.SecccionIe}" +
            //           $"_{numAleatorio}.xlsx";

            _FileName = $"RegNotasFinales_{"05674200"}" +
                        $"_{"20"}" +
                        $"_{"B0200180401"}" +
                        $"_{numRandom}.xlsx";

            _serverTemp = Path.Combine("D:\\FHUACHO\\LabGitHub\\TemplateExcelExamplesCSharp\\Console.ClosedXml.Excel", Path.Combine("temp", _FileName));
            plantilla = "Plantilla_RegistroPorNotasFinales.xlsx";
            pathPlantilla = Path.Combine("D:\\FHUACHO\\LabGitHub\\TemplateExcelExamplesCSharp\\Console.ClosedXml.Excel", Path.Combine("Plantillas", plantilla));
            File.Copy(pathPlantilla, _serverTemp, true);

            return _serverTemp;
        }
        public void GenerarExcel_Cero_a_DosAnios()
        {
            Console.WriteLine("INICIO");
            Estudiante objEstudiante = new Estudiante();
            List<Estudiante> objListEstudaintes = objEstudiante.GetEstudiantes();
            Areas objAreas = new Areas();
            List<Areas> objListaAreas = objAreas.GetAreas();

            //---------------- crear u obtener la plantilla que con la que se va a trabajar --------------
            string pathServerTempPlantilla = GetPlantilla();

            //----------------- Trabajar con la plantilla seleccionada --------------------------------
            using (XLWorkbook workbook = new XLWorkbook(pathServerTempPlantilla))
            {
                //------- GENERALIDADES: Cargando los datos a la pestaña Generalidades
                IXLWorksheet workSheetGeneral = workbook.Worksheets.Where(x => x.Name == "Generalidades").FirstOrDefault();
                workSheetGeneral.Cell("E5").Value = "0567420" + "-" + "0"; //entity.Perfil.IdInstitucion + entity.Perfil.Anexo;
                workSheetGeneral.Cell("H5").Value = "Primaria";//entity.Perfil.IdNivelInstitucion;
                workSheetGeneral.Cell("C6").Value = "JESUS";//entity.Perfil.NombreInstitucion.Replace("'", "");
                workSheetGeneral.Cell("D8").Value = "2018";// entity.AnioAcademico;
                workSheetGeneral.Cell("D9").Value = "CURRICULA NACIONAL 2017";//entity.DisenioCurricular;
                workSheetGeneral.Cell("C10").Value = "PRIMERO";//entity.DescGradoIe;
                workSheetGeneral.Cell("F10").Value = "A";//entity.DescSeccionIe;

                //-------- GENERALIDADES:Cargando las areas en la pestaña Generalidades
                workSheetGeneral.Cell("B12").Value = "AREAS";
                int nroRowArea = 14;
                foreach (var area in objListaAreas)
                {
                    workSheetGeneral.Cell(nroRowArea, 2).Value = area.AbrArea;
                    workSheetGeneral.Cell(nroRowArea, 3).Value = area.DescArea;
                    nroRowArea++;
                }

                //-------- Agregando hojas de trabajo por Area
                foreach (var area in objListaAreas)
                {
                    //--------Agregamos la hoja de trabajo
                    IXLWorksheet worksheet = workbook.Worksheets.Add(area.AbrArea);

                    //-------- HEAD: Generamos la cabecera parte 1: datos sin notas
                    worksheet.Cell("A1").Value = "ID";                    
                    worksheet.Cell("B1").Value = "CodEstudiante";                    
                    worksheet.Cell("C1").Value = "Nombres";                    

                    //-------- HEAD: Generamos la cabecera parte 2: datos con notas
                    worksheet.Cell("D1").Value = "01 - Descripción del nivel de logro alcanzado por el niño(a)";                                        
                    worksheet.Cell("E1").Value = "02 - Descripción del nivel de logro alcanzado por el niño(a)";                                        
                    worksheet.Cell("F1").Value = "03 - Descripción del nivel de logro alcanzado por el niño(a)";

                    //--------HEAD: Le damos el formato a la cabacera -------------------------
                    IXLRange rango = worksheet.Range("A1:F1");
                    rango.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    rango.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    rango.Style.Border.BottomBorderColor = XLColor.Black;
                    rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    rango.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 180, 180);
                    //rango.Style.Font.FontName = "Courier New";                
                    //rango.Style.Alignment.WrapText = true;                

                    //--------BODY: Generar la tabla de estudiantes con sus notas ----------------------
                    int nroRow = 2;
                    foreach (var estudiante in objListEstudaintes)
                    {
                        worksheet.Cell(nroRow, 1).Value = estudiante.Id;
                        worksheet.Cell(nroRow, 2).Style.NumberFormat.Format = "@";
                        worksheet.Cell(nroRow, 2).Value = estudiante.CodEstudiante.ToString();
                        worksheet.Cell(nroRow, 3).Value = estudiante.Nombres;
                        nroRow++;
                    }

                    ////--------BODY: Data Validation -------------------------------------------------------
                    //var options = new List<string>() { "AD", "A", "B", "C" };
                    //var validOptions = $"\"{String.Join(",", options)}\"";
                    //nroRow = 2;
                    //foreach (var estudiante in objListEstudaintes)
                    //{
                    //    worksheet.Cell(nroRow, 4).DataValidation.List(validOptions, true);
                    //    worksheet.Cell(nroRow, 5).DataValidation.List(validOptions, true);
                    //    worksheet.Cell(nroRow, 6).DataValidation.List(validOptions, true);
                    //    worksheet.Cell(nroRow, 7).DataValidation.List(validOptions, true);
                    //    worksheet.Cell(nroRow, 8).DataValidation.List(validOptions, true);
                    //    nroRow++;
                    //}

                    //-------FOOTER: Leyenda --------------------------------------------------------
                    worksheet.Cell(nroRow + 2, 2).Value = "LEYENDA";
                    //worksheet.Cell(nroRow + 3, 3).Value = new[]
                    //{
                    //    "01 = Se comunica oralmente en su lengua materna",
                    //    "02 = Lee diversos tipos de texto en su lengua materna lengua materna",
                    //    "03 = Escribe diversos tipos de textos en su lengua materna",
                    //    "04 = Crea Proyectos desde los lenguajes del arte"
                    //};      
                    worksheet.Cell(nroRow + 3, 2).Value = "01 = Se comunica oralmente en su lengua materna";
                    worksheet.Range(nroRow + 3, 2, nroRow + 3, 3).Merge();
                    worksheet.Cell(nroRow + 4, 2).Value = "02 = Lee diversos tipos de texto en su lengua materna lengua materna";
                    worksheet.Range(nroRow + 4, 2, nroRow + 4, 3).Merge();
                    worksheet.Cell(nroRow + 5, 2).Value = "03 = Escribe diversos tipos de textos en su lengua materna";
                    worksheet.Range(nroRow + 5, 2, nroRow + 5, 3).Merge();

                    //---------------------------- Le damos formato a la tabla en general --------------------------
                    worksheet.Columns(1, 9).AdjustToContents(); // Ajustamos el ancho de las columnas para que se muestren todos los contenidos
                                                                //worksheet.Column(9).AdjustToContents(); // Ajustamos el ancho de una columna de acuerdo a su contenido
                    worksheet.Column(9).Width = 100;
                    worksheet.Columns(4, 8).Width = 6;
                }

                Console.WriteLine("ANTES DE GUARDAR");
                workbook.SaveAs(pathServerTempPlantilla);
                Console.WriteLine("DESPUES DE GUARDAR");
            }
            Console.WriteLine("FIN");
            Console.ReadLine();
        }
        public void GenerarExcel_TresAnios_Hasta_PrimeroSecundaria()
        {
            Estudiante objEstudiante = new Estudiante();
            List<Estudiante> objListEstudaintes = objEstudiante.GetEstudiantes();
            Areas objAreas = new Areas();
            List<Areas> objListaAreas = objAreas.GetAreas();

            //---------------- crear u obtener la plantilla que con la que se va a trabajar --------------
            string pathServerTempPlantilla = GetPlantilla();

            //----------------- Trabajar con la plantilla seleccionada --------------------------------
            using (XLWorkbook workbook = new XLWorkbook(pathServerTempPlantilla))
            {
                //------- GENERALIDADES: Cargando los datos a la pestaña Generalidades
                IXLWorksheet workSheetGeneral = workbook.Worksheets.Where(x => x.Name == "Generalidades").FirstOrDefault();
                workSheetGeneral.Cell("E5").Value = "0567420" + "-" + "0"; //entity.Perfil.IdInstitucion + entity.Perfil.Anexo;
                workSheetGeneral.Cell("H5").Value = "Primaria";//entity.Perfil.IdNivelInstitucion;
                workSheetGeneral.Cell("C6").Value = "JESUS";//entity.Perfil.NombreInstitucion.Replace("'", "");
                workSheetGeneral.Cell("D8").Value = "2018";// entity.AnioAcademico;
                workSheetGeneral.Cell("D9").Value = "CURRICULA NACIONAL 2017";//entity.DisenioCurricular;
                workSheetGeneral.Cell("C10").Value = "PRIMERO";//entity.DescGradoIe;
                workSheetGeneral.Cell("F10").Value = "A";//entity.DescSeccionIe;

                //-------- GENERALIDADES:Cargando las areas en la pestaña Generalidades
                workSheetGeneral.Cell("B12").Value = "AREAS";
                int nroRowArea = 14;
                foreach (var area in objListaAreas)
                {
                    workSheetGeneral.Cell(nroRowArea, 2).Value = area.AbrArea;
                    workSheetGeneral.Cell(nroRowArea, 3).Value = area.DescArea;
                    nroRowArea++;
                }

                //-------- Agregando hojas de trabajo por Area
                foreach (var area in objListaAreas)
                {
                    //--------Agregamos la hoja de trabajo
                    IXLWorksheet worksheet = workbook.Worksheets.Add(area.AbrArea);

                    //-------- HEAD: Generamos la cabecera parte 1: datos sin notas
                    worksheet.Cell("A1").Value = "ID";
                    worksheet.Range("A1:A3").Merge();
                    worksheet.Cell("B1").Value = "CodEstudiante";
                    worksheet.Range("B1:B3").Merge();
                    worksheet.Cell("C1").Value = "Nombres";
                    worksheet.Range("C1:C3").Merge();

                    //-------- HEAD: Generamos la cabecera parte 2: datos con notas
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


                    //--------HEAD: Le damos el formato a la cabacera -------------------------
                    IXLRange rango = worksheet.Range("A1:I3");
                    rango.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    rango.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    rango.Style.Border.BottomBorderColor = XLColor.Black;
                    rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    rango.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 180, 180);
                    //rango.Style.Font.FontName = "Courier New";                
                    //rango.Style.Alignment.WrapText = true;                

                    //--------BODY: Generar la tabla de estudiantes con sus notas ----------------------
                    int nroRow = 4;
                    foreach (var estudiante in objListEstudaintes)
                    {
                        worksheet.Cell(nroRow, 1).Value = estudiante.Id;
                        worksheet.Cell(nroRow, 2).Style.NumberFormat.Format = "@";
                        worksheet.Cell(nroRow, 2).Value = estudiante.CodEstudiante.ToString();
                        worksheet.Cell(nroRow, 3).Value = estudiante.Nombres;
                        nroRow++;
                    }

                    //--------BODY: Data Validation -------------------------------------------------------
                    List<string> options = new List<string>() { "AD", "A", "B", "C" };
                    string validOptions = $"\"{String.Join(",", options)}\"";
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

                    //-------FOOTER: Leyenda --------------------------------------------------------
                    worksheet.Cell(nroRow + 2, 2).Value = "LEYENDA";
                    //worksheet.Cell(nroRow + 3, 3).Value = new[]
                    //{
                    //    "01 = Se comunica oralmente en su lengua materna",
                    //    "02 = Lee diversos tipos de texto en su lengua materna lengua materna",
                    //    "03 = Escribe diversos tipos de textos en su lengua materna",
                    //    "04 = Crea Proyectos desde los lenguajes del arte"
                    //};      
                    worksheet.Cell(nroRow + 3, 2).Value = "01 = Se comunica oralmente en su lengua materna";
                    worksheet.Range(nroRow + 3, 2, nroRow + 3, 3).Merge();
                    worksheet.Cell(nroRow + 4, 2).Value = "02 = Lee diversos tipos de texto en su lengua materna lengua materna";
                    worksheet.Range(nroRow + 4, 2, nroRow + 4, 3).Merge();
                    worksheet.Cell(nroRow + 5, 2).Value = "03 = Escribe diversos tipos de textos en su lengua materna";
                    worksheet.Range(nroRow + 5, 2, nroRow + 5, 3).Merge();
                    worksheet.Cell(nroRow + 6, 2).Value = "04 = Crea Proyectos desde los lenguajes del arte";
                    worksheet.Range(nroRow + 6, 2, nroRow + 6, 3).Merge();

                    //---------------------------- Le damos formato a la tabla en general --------------------------
                    worksheet.Columns(1, 9).AdjustToContents(); // Ajustamos el ancho de las columnas para que se muestren todos los contenidos
                                                                //worksheet.Column(9).AdjustToContents(); // Ajustamos el ancho de una columna de acuerdo a su contenido
                    worksheet.Column(9).Width = 100;
                    worksheet.Columns(4, 8).Width = 6;
                }

                Console.WriteLine("ANTES DE GUARDAR");
                workbook.SaveAs(pathServerTempPlantilla);
                Console.WriteLine("DESPUES DE GUARDAR");
            }
            Console.WriteLine("FIN");
            Console.ReadLine();
        }
        public void GenerarExcel_Segundo_a_QuintoDeSecundaria()
        {
            Console.WriteLine("INICIO");
            Estudiante objEstudiante = new Estudiante();
            List<Estudiante> objListEstudaintes = objEstudiante.GetEstudiantes();
            Areas objAreas = new Areas();
            List<Areas> objListaAreas = objAreas.GetAreas();
            //---------------- crear u obtener la plantilla que con la que se va a trabajar --------------
            string pathServerTempPlantilla = GetPlantilla();

            //----------------- Trabajar con la plantilla seleccionada --------------------------------
            using (XLWorkbook workbook = new XLWorkbook(pathServerTempPlantilla))
            {
                //------- GENERALIDADES: Cargando los datos a la pestaña Generalidades
                IXLWorksheet workSheetGeneral = workbook.Worksheets.Where(x => x.Name == "Generalidades").FirstOrDefault();
                workSheetGeneral.Cell("E5").Value = "0567420" + "-" + "0"; //entity.Perfil.IdInstitucion + entity.Perfil.Anexo;
                workSheetGeneral.Cell("H5").Value = "Primaria";//entity.Perfil.IdNivelInstitucion;
                workSheetGeneral.Cell("C6").Value = "JESUS";//entity.Perfil.NombreInstitucion.Replace("'", "");
                workSheetGeneral.Cell("D8").Value = "2018";// entity.AnioAcademico;
                workSheetGeneral.Cell("D9").Value = "CURRICULA NACIONAL 2017";//entity.DisenioCurricular;
                workSheetGeneral.Cell("C10").Value = "PRIMERO";//entity.DescGradoIe;
                workSheetGeneral.Cell("F10").Value = "A";//entity.DescSeccionIe;

                //-------- GENERALIDADES:Cargando las areas en la pestaña Generalidades
                workSheetGeneral.Cell("B12").Value = "AREAS";
                int nroRowArea = 14;
                foreach (var area in objListaAreas)
                {
                    workSheetGeneral.Cell(nroRowArea, 2).Value = area.AbrArea;
                    workSheetGeneral.Cell(nroRowArea, 3).Value = area.DescArea;
                    nroRowArea++;
                }
                
                //--------Agregamos la hoja de trabajo
                IXLWorksheet worksheet = workbook.Worksheets.Add("NF");

                //-------- HEAD: Generar la tabla de estudiantes con sus notas --------------------
                int nroRowHead = 1;
                int nroColHead = 3;

                worksheet.Cell(nroRowHead, 1).Value = "ID";
                worksheet.Cell(nroRowHead, 2).Value = "CodEstudiante";
                worksheet.Cell(nroRowHead, 3).Value = "Nombres";
                
                foreach (var area in objListaAreas)
                {
                    nroColHead++;
                    worksheet.Cell(nroRowHead, nroColHead).Value = area.AbrArea;                    
                }
                //--------HEAD: Le damos el formato a la cabecera ---------------------------------
                IXLRange rango = worksheet.Range(nroRowHead, 1, nroRowHead, nroColHead);
                rango.Style.Border.OutsideBorder     = XLBorderStyleValues.Thin;
                rango.Style.Border.InsideBorder      = XLBorderStyleValues.Thin;
                rango.Style.Border.BottomBorderColor = XLColor.Black;
                rango.Style.Alignment.Horizontal     = XLAlignmentHorizontalValues.Center;
                rango.Style.Alignment.Vertical       = XLAlignmentVerticalValues.Center;
                rango.Style.Fill.BackgroundColor     = XLColor.FromArgb(180, 180, 180);                      

                //--------BODY: Generar la tabla de estudiantes con sus notas ----------------------
                int nroRowBody = 2;
                int nroColBody;

                foreach (Estudiante estudiante in objListEstudaintes)
                {
                    worksheet.Cell(nroRowBody, 1).Value = estudiante.Id;
                    worksheet.Cell(nroRowBody, 2).Style.NumberFormat.Format = "@";
                    worksheet.Cell(nroRowBody, 2).Value = estudiante.CodEstudiante.ToString();
                    worksheet.Cell(nroRowBody, 3).Value = estudiante.Nombres;

                    nroColBody = 4;
                    foreach (var area in objListaAreas)
                    {
                        worksheet.Cell(nroRowBody, nroColBody).Value = "";
                        //--------BODY: Data Validation -------------------------------------------------------  
                        worksheet.Cell(nroRowBody, nroColBody).DataValidation.Decimal.Between(0, 20);

                        nroColBody++;
                    }
                    nroRowBody++;
                }                                

                worksheet.Columns(nroRowHead, nroColHead).AdjustToContents();              

                Console.WriteLine("ANTES DE GUARDAR");
                workbook.SaveAs(pathServerTempPlantilla);
                Console.WriteLine("DESPUES DE GUARDAR");
            }
            Console.WriteLine("FIN");
            Console.ReadLine();
        }
    }
}

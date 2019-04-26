using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel.ExcelData
{
    public class ExcelNotasFinales //: IDisposable
    {
        public string BasePath { get; set; }
        public BEPerfil Perfil { get; set; }
        public string FileName { get; set; }
        public string ServerTemp { get; set; }
        public string IdDisenioCurricular { get; set; }
        public string DescDisenioCurricular { get; set; }
        public string IdGrado { get; set; }
        public string DescGrado { get; set; }
        public string IdSeccion { get; set; }
        public string DescSeccion { get; set; }
        public string PlantillaExcel { get; set; }
        public string TipoPlantillaExcel { get; set; }
        public string PathPlantillaExcel { get; set; }
        public int TipoPlantilla { get; set; }

        public ExcelNotasFinales(string basePath, BEPerfil perfil, string idDisenioCurricular, string descDisenioCurricular, string idGrado, string idSeccion, string descGrado, string descSeccion, int tipoPlantilla)
        {
            BasePath = basePath;
            Perfil = perfil;
            IdDisenioCurricular = idDisenioCurricular;
            DescDisenioCurricular = descDisenioCurricular;
            IdGrado = idGrado;
            IdSeccion = idSeccion;
            PlantillaExcel = "Plantilla_RegistroPorNotasFinales.xlsx";
            TipoPlantilla = tipoPlantilla; // "Plantilla excel notas finales (de 0 a 2 años)"
            DescGrado = descGrado;
            DescSeccion = descSeccion;
        }
        private void AsignarPlantilla()
        {
            //---------------- crear u obtener la plantilla que con la que se va a trabajar --------------                                       
            Random random = new Random();
            int numRandom = random.Next(1000, 99999);

            FileName = $"RegNotasFinales_{Perfil.IdInstitucion + Perfil.Anexo}" +
                       $"_{IdDisenioCurricular}" +
                       $"_{Perfil.IdNivelInstitucion.Trim() + Perfil.AnioAcademico + IdGrado + IdSeccion}" +
                       $"_{numRandom}.xlsx";

            ServerTemp = Path.Combine(BasePath, Path.Combine("temp", FileName));
            PathPlantillaExcel = Path.Combine(BasePath, Path.Combine("Plantillas", PlantillaExcel));
            File.Copy(PathPlantillaExcel, ServerTemp, true);
        }
        //public byte[] CrearExcelNotas(List<BEPlantillaNotasExcel> notasEstudiantes)
        public void CrearExcelNotas(List<BEPlantillaNotasExcel> notasEstudiantes)
        {
            AsignarPlantilla();
            switch (TipoPlantilla)
            {
                case 1:
                    GenerarExcelNF_Cero_a_DosAnios();
                    break;
                case 2:
                    GenerarExcelNF_TresAnios_Hasta_PrimeroSecundaria(notasEstudiantes);
                    break;
                case 3:
                    GenerarExcelNF_Segundo_a_QuintoDeSecundaria(notasEstudiantes);
                    break;
                default:
                    break;
            }
            //return FuncionesComun.ReadFileToBytes(ServerTemp);
        }
        private void GenerarExcelNF_Cero_a_DosAnios()
        {

        }
        private void GenerarExcelNF_TresAnios_Hasta_PrimeroSecundaria(List<BEPlantillaNotasExcel> notasEstudiantes)
        {
            using (XLWorkbook workbook = new XLWorkbook(ServerTemp))
            {
                BEAreasPorGradoPorIe[] areas = GetAreasxGradoxIE();
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
                foreach (var area in areas)
                {
                    workSheetGeneral.Cell(nroRowArea, 2).Value = area.abr_area;
                    workSheetGeneral.Cell(nroRowArea, 3).Value = area.dsc_area;
                    nroRowArea++;
                }

                //-------- Agregando hojas de trabajo por Area
                foreach (var area in areas)
                {
                    //--------Agregamos la hoja de trabajo
                    IXLWorksheet worksheet = workbook.Worksheets.Add(area.abr_area);

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

                    //var notasEstudiantesPivot = notasEstudiantes.GroupBy(c => new {
                    //    c.PersonaId,
                    //    c.EstudianteApellidoPaterno,
                    //    c.EstudianteApellidoMaterno,
                    //    c.EstudianteNombre
                    //}).Select(g => new BEPlantillaNotasExcel {
                    //    PersonaId = g.Key.PersonaId,
                    //    EstudianteApellidoPaterno = g.Key.EstudianteApellidoPaterno,
                    //    EstudianteApellidoMaterno = g.Key.EstudianteApellidoMaterno,
                    //    EstudianteNombre = g.Key.EstudianteNombre,
                    //    subject = g.GroupBy(f => f)
                    //});
                }

            }
        }
        private void GenerarExcelNF_Segundo_a_QuintoDeSecundaria(List<BEPlantillaNotasExcel> notasEstudiantes)
        {
            using (XLWorkbook workbook = new XLWorkbook(ServerTemp))
            {
                BEAreasPorGradoPorIe[] areas = GetAreasxGradoxIE();
                //------- GENERALIDADES: Cargando los datos a la pestaña Generalidades
                IXLWorksheet workSheetGeneral = workbook.Worksheets.Where(x => x.Name == "Generalidades").FirstOrDefault();
                workSheetGeneral.Cell("E5").Value = Perfil.IdInstitucion + "-" + Perfil.Anexo;
                workSheetGeneral.Cell("H5").Value = Perfil.DesNivelInstitucion;
                workSheetGeneral.Cell("C6").Value = Perfil.NombreInstitucion.Replace("'", "");
                workSheetGeneral.Cell("D8").Value = Perfil.AnioAcademico;
                workSheetGeneral.Cell("D8").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                workSheetGeneral.Cell("D9").Value = DescDisenioCurricular;
                workSheetGeneral.Cell("C10").Value = DescGrado;
                workSheetGeneral.Cell("F10").Value = DescSeccion;

                //-------- GENERALIDADES:Cargando las areas en la pestaña Generalidades
                workSheetGeneral.Cell("B12").Value = "AREAS";
                int nroRowArea = 14;

                foreach (var area in areas)
                {
                    workSheetGeneral.Cell(nroRowArea, 2).Value = area.abr_area;
                    workSheetGeneral.Cell(nroRowArea, 3).Value = area.dsc_area;
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

                foreach (var area in areas)
                {
                    nroColHead++;
                    worksheet.Cell(nroRowHead, nroColHead).Value = area.abr_area;
                }

                //--------HEAD: Le damos el formato a la cabecera ---------------------------------
                IXLRange rango = worksheet.Range(nroRowHead, 1, nroRowHead, nroColHead);
                rango.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rango.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rango.Style.Border.BottomBorderColor = XLColor.Black;
                rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rango.Style.Fill.BackgroundColor = XLColor.FromArgb(180, 180, 180);

                //--------BODY: Generar la tabla de estudiantes con sus notas ----------------------
                int nroRowBody = 2;
                int nroColBody;

                List<BEPlantillaNotasExcel> soloEstuadintes = notasEstudiantes.Select(item => new BEPlantillaNotasExcel {
                    PersonaId = item.PersonaId,
                    EstudianteCodigo = item.EstudianteCodigo,
                    EstudianteApellidoPaterno = item.EstudianteApellidoPaterno,
                    EstudianteApellidoMaterno = item.EstudianteApellidoMaterno,
                    EstudianteNombre = item.EstudianteNombre
                }).Distinct().ToList();

                foreach (BEPlantillaNotasExcel estudiante in soloEstuadintes)
                {
                    worksheet.Cell(nroRowBody, 1).Value = estudiante.PersonaId;
                    worksheet.Cell(nroRowBody, 2).Style.NumberFormat.Format = "@";
                    worksheet.Cell(nroRowBody, 2).Value = estudiante.EstudianteCodigo.ToString();
                    worksheet.Cell(nroRowBody, 3).Value = $"{estudiante.EstudianteApellidoPaterno.Trim()} {estudiante.EstudianteApellidoMaterno.Trim()} {estudiante.EstudianteNombre.Trim()}";

                    nroColBody = 4;
                    foreach (BEAreasPorGradoPorIe area in areas)
                    {
                        BEPlantillaNotasExcel notaEstudiantePorArea = notasEstudiantes.Where(x => x.AreaId == area.id_area && x.PersonaId == estudiante.PersonaId).FirstOrDefault();
                        worksheet.Cell(nroRowBody, nroColBody).Value = (notaEstudiantePorArea == null) ? "" : notaEstudiantePorArea.Nota ?? "";                        
                        worksheet.Cell(nroRowBody, nroColBody).DataValidation.Decimal.Between(0, 20);//--> BODY: Data Validation
                        nroColBody++;
                    }
                    nroRowBody++;
                }

                ////// otra forma                                
                ////int idPersonaCell;
                ////foreach (var area in areas)
                ////{
                ////    nroColBody = 4;
                ////    nroRowBody = 2;
                ////    var estudiantesPorArea = notasEstudiantes.Where(x => x.IdArea == area.id_area).ToList();

                ////    foreach (BEPlantillaNotasExcel notaEstudiante in estudiantesPorArea)
                ////    {
                ////        idPersonaCell = int.Parse(worksheet.Cell(nroRowBody, 1).GetString());
                ////        if (idPersonaCell == notaEstudiante.IdPersona)
                ////        {
                ////            worksheet.Cell(nroRowBody, nroColBody).Value = (notaEstudiante.Nota ?? "");                            
                ////            worksheet.Cell(nroRowBody, nroColBody).DataValidation.Decimal.Between(0, 20);//--> Data Validation  
                ////            nroRowBody++;
                ////        }                        
                ////    }                                                                          
                ////    nroColBody++;
                ////}
                worksheet.Columns(nroRowHead, nroColHead).AdjustToContents();
                workbook.SaveAs(ServerTemp);

            }
        }
        private BEAreasPorGradoPorIe[] GetAreasxGradoxIE()
        {
            var filtro = new BEAreasPorGradoPorIe
            {
                cod_mod = Perfil.IdInstitucion,
                anexo = Perfil.Anexo,
                id_disenio = IdDisenioCurricular,
                id_anio = short.Parse(Perfil.AnioAcademico),
                id_nivel = Perfil.IdNivelInstitucion,
                id_grado = IdGrado,
                es_conducta = 0
            };

            BEAreasPorGradoPorIe[] arrAreas = null;
            var siagieService = new NotasFinalesDA();
            
            arrAreas = siagieService.AreasPorGradoPorIeListarSoloAreasDynamic(filtro)
                            .Where(p => p.es_area_agrupadora.GetValueOrDefault(0) == 0
                                    && p.es_tutoria.GetValueOrDefault(0) == 0)
                            .ToArray();
            

            //BUG8124-MHUAYTA-07/07/2017-Se realiza limpieza de los IDs areas eliminando espacios en blanco para evitar errores en el armado de la plantilla
            if (arrAreas.Any())
            {
                foreach (var area in arrAreas)
                {
                    area.id_area = string.IsNullOrWhiteSpace(area.id_area) ? "" : area.id_area.Trim();
                }
            }
            return arrAreas;
        }

        //#region IDisposable Members

        //private bool _disposed = false;

        //public void Dispose()
        //{
        //    Dispose(true);
        //    GC.SuppressFinalize(this);
        //}

        //protected virtual void Dispose(bool disposing)
        //{
        //    if (_disposed)
        //    {
        //        return;
        //    }

        //    if (disposing)
        //    {
        //        if (!string.IsNullOrEmpty(ServerTemp))
        //        {
        //            File.Delete(ServerTemp);
        //        }
        //    }
        //    _disposed = true;
        //}

        //#endregion
    }
}

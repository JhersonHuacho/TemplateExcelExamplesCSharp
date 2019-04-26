using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            pivotLinq();
            //ExcelNotasFinales objExcelNotasFinales = new ExcelNotasFinales();
            //objExcelNotasFinales.GenerarExcel_TresAnios_Hasta_PrimeroSecundaria();
            //objExcelNotasFinales.GenerarExcel_Cero_a_DosAnios();
            //objExcelNotasFinales.GenerarExcel_Segundo_a_QuintoDeSecundaria();

            ClosedXml.Excel.ExcelData.BEPerfil perfil = new ClosedXml.Excel.ExcelData.BEPerfil()
            {
                IdInstitucion = "0238725",
                Anexo = "0",
                IdNivelInstitucion = "F0",
                DesNivelInstitucion = "Secundaria",
                NombreInstitucion = "JOSE R. TUESTA RUIZ",
                AnioAcademico = "2018"
            };

            var basePath = Directory.GetCurrentDirectory(); //"D:\\FHUACHO\\proyectos59\\v3.19.2\\Web\\Release 1.0\\Web.Siagie2\\";
            var CodigoModular = "0238725";
            var Anexo = "0";
            var DisenioCurricular = "15";
            var AnioAcademico = 2018;
            var Nivel = "F0";
            var GradoIe = "11";
            var SecccionIe = "01";
            var EsCicloAvanzado = false;
            var DescDisenioCurricular = "DISEÑO CURRICULAR NACIONAL 2009";
            var DescGradoIe = "PRIMERO";
            var DescSeccionIe = "A";
            var PermiteCompetenciaAlcanzada = false;            

            //using (ClosedXml.Excel.ExcelData.ExcelNotasFinales objExcelNotasFinalesData =
            //    new ClosedXml.Excel.ExcelData.ExcelNotasFinales(basePath, perfil, DisenioCurricular, GradoIe, SecccionIe, DescGradoIe, DescSeccionIe, 3))
            //{
            //    var notasExcel = new NotasFinalesDA();
            //    var resultado = new List<ClosedXml.Excel.ExcelData.BEPlantillaNotasExcel>();
            //    resultado = notasExcel.GetEvaluacionNotasPorEstudianteGeneral(CodigoModular, Anexo, AnioAcademico, Nivel, DisenioCurricular, "", GradoIe, SecccionIe, "");
            //    objExcelNotasFinalesData.CrearExcelNotas(resultado);
            //}
            ClosedXml.Excel.ExcelData.ExcelNotasFinales objExcelNotasFinalesData =
                new ClosedXml.Excel.ExcelData.ExcelNotasFinales(basePath, perfil, DisenioCurricular, DescDisenioCurricular, GradoIe, SecccionIe, DescGradoIe, DescSeccionIe, 3);
            var notasExcel = new NotasFinalesDA();
            var resultado = new List<ClosedXml.Excel.ExcelData.BEPlantillaNotasExcel>();
            resultado = notasExcel.GetEvaluacionNotasPorEstudianteGeneral(CodigoModular, Anexo, AnioAcademico, Nivel, DisenioCurricular, "", GradoIe, SecccionIe, "");
            objExcelNotasFinalesData.CrearExcelNotas(resultado);
            Console.ReadLine();

        }   
        
        private static void pivotLinq()
        {
            List<Visit> Visits = new List<Visit>
            {
                new Visit(1, new DateTime(2015,2,24), "A"),
                new Visit(2, new DateTime(2015,2,23), "S"),
                new Visit(2, new DateTime(2015,2,24), "D"),
                new Visit(4, new DateTime(2015,2,22), "S"),
                new Visit(2, new DateTime(2015,2,22), "A"),
                new Visit(2, new DateTime(2015,2,22), "B"),
                new Visit(3, new DateTime(2015,2,23), "A"),
                new Visit(1, new DateTime(2015,2,23), "A"),
                new Visit(1, new DateTime(2015,2,24), "D"),
                new Visit(4, new DateTime(2015,2,24), "S"),
                new Visit(4, new DateTime(2015,2,22), "S"),
                new Visit(2, new DateTime(2015,2,24), "S"),
                new Visit(3, new DateTime(2015,2,24), "D")
            };

            Console.WriteLine("");
            foreach (var visit in Visits)
            {
                Console.WriteLine(visit.PersonelId + " " + visit.VisitDate + " " + visit.VisitTypeId);
            }
            Console.WriteLine("");
            //static headers
            var qry = Visits.GroupBy(v => new { v.VisitDate, v.PersonelId })
                .Select(g => new {
                    VisitDate = g.Key.VisitDate,
                    PersonelId = g.Key.PersonelId,
                    A = g.Where(d => d.VisitTypeId == "A").Count(),
                    B = g.Where(d => d.VisitTypeId == "B").Count(),
                    D = g.Where(d => d.VisitTypeId == "D").Count(),
                    S = g.Where(d => d.VisitTypeId == "S").Count()
                });

            foreach (var q in qry)
            {
                Console.WriteLine(Convert.ToString(q.PersonelId) + " " + Convert.ToString(q.VisitDate) + " " + q.A + " " + q.B + " " + q.D + " " + q.S);
            }
            Console.WriteLine("");
            //dynamic headers
            var qry1 = Visits.GroupBy(v => new { v.VisitDate, v.PersonelId })
                .Select(g => new {
                    PersonelId = g.Key.PersonelId,
                    VisitDate = g.Key.VisitDate,
                    subject = g.GroupBy(f => f.VisitTypeId).Select(m => new { Sub = m.Key, Score = m.Count() })
                });

            var totalHead = Visits.Select(item => item.VisitTypeId ).Distinct().Count();

            foreach (var q in qry1)
            {
                foreach (var item in q.subject)
                {
                    Console.WriteLine(Convert.ToString(q.PersonelId) + " " + Convert.ToString(q.VisitDate) + " " + item.Sub + " " + item.Score);
                }                
            }

        }

        // class definition
        public class Visit
        {
            private int id = 0;
            private DateTime vd;
            private string vt = string.Empty;

            public Visit(int _id, DateTime _vd, string _vt)
            {
                id = _id;
                vd = _vd;
                vt = _vt;
            }

            public int PersonelId
            {
                get { return id; }
                set { id = value; }
            }

            public DateTime VisitDate
            {
                get { return vd; }
                set { vd = value; }
            }

            public string VisitTypeId
            {
                get { return vt; }
                set { vt = value; }
            }
        }

    }
}

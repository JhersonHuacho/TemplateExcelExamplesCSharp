using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel
{
    public class Estudiante
    {
        public string Id { get; set; }
        public string CodEstudiante { get; set; }
        public string Nombres { get; set; }

        public List<Estudiante> GetEstudiantes()
        {
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

            return objListEstudaintes;
        }
    }
}

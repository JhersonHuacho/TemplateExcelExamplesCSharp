using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel
{
    public class Areas
    {
        public string IdArea { get; set; }
        public string AbrArea { get; set; }
        public string DescArea { get; set; }
        public int IdAsignatura { get; set; }

        public List<Areas> GetAreas()
        {
            List<Areas> objListaAreas = new List<Areas>()
                {
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "ART_Y_CULT",
                        DescArea = "ARTE Y CULTURA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "CAST_SEGNL",
                        DescArea = "CASTELLANO SEGUNDA LENGUA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "CIENC_TEC",
                        DescArea = "CIENCIA Y TECNOLOGIA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "COMU",
                        DescArea = "COMUNICACION"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "EFIS",
                        DescArea = "EDUCACION FISICA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "EREL",
                        DescArea = "EDUCACION RELIGIOSA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "INGLES_EXT",
                        DescArea = "INGLES COMO LENGUA EXTERNA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "MATE",
                        DescArea = "MATEMATICA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "PPSS",
                        DescArea = "PERSONAL SOCIAL"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "GEST_AUTO",
                        DescArea = "GESTIONA SU APRENDIZAJE DE MANERA AUTÓMATICA"
                    },
                    new Areas()
                    {
                        IdArea = "30789573",
                        AbrArea = "DESEN_TIC",
                        DescArea = "SE DESENVUELVE EN ENTORNOS VIRTUALES GENERADOS POR LAS TIC"
                    }
                };

            return objListaAreas;
        }
    }
}

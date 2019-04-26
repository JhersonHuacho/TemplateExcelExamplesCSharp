using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel.ExcelData
{
    public class BEPerfil
    {
        public string Accion { get; set; }
        public string IdSistema { get; set; }
        public string LoginUsuario { get; set; }
        public string NombreUsuario { get; set; }
        public string IdInstitucion { get; set; }
        public string Anexo { get; set; }
        public string NombreInstitucion { get; set; }
        public string IdNivelInstitucion { get; set; }
        public string DesNivelInstitucion { get; set; }
        public string IdTipoSede { get; set; }
        public string DesTipoSede { get; set; }
        public string AnioAcademico { get; set; }
        public string DesAnioAcademico { get; set; }
        public string IdRol { get; set; }
        public bool? Descentralizado { get; set; }
        public string IdGestion { get; set; }
        public string IdModalidad { get; set; }
        public string IdGrupoEtiqueta { get; set; }
        //propiedades extendidas
        public string DscGestion { get; set; }
        public string Dre { get; set; }
        public string Ugel { get; set; }
        public int IdPersona { get; set; }
        public short? tiene_responsable_unico { get; set; }
    }
}

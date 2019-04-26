using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel.ExcelData
{
    public class BEAreasPorGradoPorIe
    {
        
        public short? horas_efectivas { get; set; }
        
        public short id_anio { get; set; }
        
        public string id_area { get; set; }
        
        public string id_area_agrupadora_padre { get; set; }
        
        public string id_area_relacionada { get; set; }
        
        public string id_disenio { get; set; }
        
        public short? estado_cierre { get; set; }
        
        public string id_grado { get; set; }
        
        public string id_periodo { get; set; }
        
        public string id_seccion { get; set; }
        
        public string id_tipo_area { get; set; }
        
        public short? idtipocalificacionevaluacion { get; set; }
        
        public string num_asignatura { get; set; }
        
        public short? num_silabo { get; set; }
        
        public string id_nivel { get; set; }
        
        public short? tiene_recuperacion { get; set; }
        
        public short? es_tutoria { get; set; }
        
        public short? es_opcional { get; set; }
        
        public string abr_area { get; set; }
        
        public string anexo { get; set; }
        
        //public BEAsignaturaPorGradoPorIe[] asignaturas_notas_libreta { get; set; }
        
        public string cod_mod { get; set; }
        
        public string dsc_area { get; set; }
        
        public string dsc_area_agrupadora { get; set; }
        
        public short? es_taller { get; set; }
        
        public string dsc_disenio { get; set; }
        
        public short? es_area_agrupadora { get; set; }
        
        public short? es_area_relacionada { get; set; }
        
        public short? es_competencia_transversal { get; set; }
        
        public short? es_comunicacion { get; set; }
        
        public short? es_conducta { get; set; }
        
        public int es_exonerada { get; set; }
        
        public string dsc_grado { get; set; }
        
        public int tipo_area { get; set; }
    }
}

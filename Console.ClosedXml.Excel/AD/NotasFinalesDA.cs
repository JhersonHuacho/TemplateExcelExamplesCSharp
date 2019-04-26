using ClosedXml.Excel.ExcelData;
using Dapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel
{
    public class NotasFinalesDA : BaseConnection
    {
        public List<BEPlantillaNotasExcel> GetEvaluacionNotasPorEstudianteGeneral(string codMod, string anexo, int idAnio, string idNivel, string idDisenio,
                                                                                     string idPeriodo, string idGrado, string idSeccion, string idArea)
        {
            var result = new List<BEPlantillaNotasExcel>();
            var sql = "USP_SEL_EVALUACION_NOTAS_POR_ESTUDIANTE_GENERAL";
            using (IDbConnection cn = new SqlConnection(GetConnection()))
            {
                result = cn.Query<BEPlantillaNotasExcel>(sql, new {
                                                                    cod_mod = codMod,
                                                                    anexo,
                                                                    id_anio = idAnio,
                                                                    id_nivel = idNivel,
                                                                    id_disenio = idDisenio,
                                                                    id_periodo = idPeriodo,
                                                                    id_grado = idGrado,
                                                                    id_seccion = idSeccion,
                                                                    id_area = idArea
                                                                  }
                                                            , commandType: CommandType.StoredProcedure).ToList();
            }
            return result;
        }
        public List<BEAreasPorGradoPorIe> AreasPorGradoPorIeListarSoloAreasDynamic(BEAreasPorGradoPorIe objParams)
        {
            var result = new List<BEAreasPorGradoPorIe>();
            var sql = "USP_AREAS_POR_GRADO_POR_IE_SEL_LISTAR_SOLO_AREAS";
            using (IDbConnection cn = new SqlConnection(GetConnection()))
            {
                result = cn.Query<BEAreasPorGradoPorIe>(sql, new {
                    objParams.cod_mod,
                    objParams.anexo,
                    objParams.id_disenio,
                    objParams.id_anio,
                    objParams.id_nivel,
                    objParams.id_grado,
                    objParams.es_conducta
                }, commandType: CommandType.StoredProcedure ).ToList();
            }
            return result;
        }
    }
}

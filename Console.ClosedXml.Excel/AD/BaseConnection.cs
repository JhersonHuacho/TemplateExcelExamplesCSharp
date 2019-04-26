using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXml.Excel
{
    public class BaseConnection
    {
        public static string GetConnection()
        {
            string cadenaConexion = $"Data Source=10.1.1.186;Initial Catalog=siagie2_20;User ID=user_siagie;Password=siagie2017";
            return cadenaConexion;
        }
    }
}

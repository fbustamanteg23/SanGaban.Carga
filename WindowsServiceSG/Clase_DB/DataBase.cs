using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Clase_DB
{
    public class DataBase
    {
        public String Conexion() { 
            String connection = System.Configuration.ConfigurationManager.ConnectionStrings["bd"].ConnectionString;

            return connection;
        }
    }
}

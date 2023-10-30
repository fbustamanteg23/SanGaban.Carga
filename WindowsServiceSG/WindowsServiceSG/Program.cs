using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace WindowsServiceSG
{
    internal static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main()
        {


            if (Environment.UserInteractive)
            {
                Service1 servicio = new Service1();
                servicio.run();
            }
            else //Ejecutamos el código por defecto de un servicio
            {
                //Flujo Normal de un servicio Windows
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                new Service1()
                };
                ServiceBase.Run(ServicesToRun);
            }

        }
    }
}

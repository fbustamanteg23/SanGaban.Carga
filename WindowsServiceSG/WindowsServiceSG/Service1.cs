using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Clase_DB;
 

namespace WindowsServiceSG
{
    public partial class Service1 : ServiceBase
    {
        Timer timerRD = new Timer(); // name space(using System.Timers;) 
        Timer timerCH = new Timer(); // name space(using System.Timers;) 
        String Ruta_Excel = "";
        string[] files;
        String Nombre_Archivo = "";
        String Nombre_Archivo_ch = "";
        int indice = 0;
        int Hoja = 0;
        String hora;

        #region Codigo_Prueba
        public void run()
        {
            DataTable dt1 = new DataTable();
            ClassData obj = new ClassData();
            obj.Obtener_Parametros();
            dt1 = obj.Obtener_Parametros();
            Int32 Tiempo_RD = Convert.ToInt32(dt1.Rows[0][0].ToString());
            Int32 Tiempo_CH = Convert.ToInt32(dt1.Rows[1][0].ToString());
             Ruta_Excel = (dt1.Rows[2][0].ToString());

            //--------------obteniendo el nombre de archivo RD--------------------------//
            files = Directory.GetFiles(Ruta_Excel); // Obtener archivo de una carpeta
            indice = files[0].IndexOf("\\" , 4);
          
            Nombre_Archivo = files[1].Substring(indice, files[1].Length- indice);
            Nombre_Archivo = Nombre_Archivo.Substring(1, Nombre_Archivo.Length-1);

            //--------------obteniendo el nombre de archivo CH--------------------------//

            indice = files[0].IndexOf("\\", 4);

            Nombre_Archivo_ch = files[0].Substring(indice, files[0].Length - indice);
            Nombre_Archivo_ch = Nombre_Archivo_ch.Substring(1, Nombre_Archivo_ch.Length - 1);



            //-------------------------------------------------------------------------//

            string Date = DateTime.Now.ToString("dd");
            Hoja = Convert.ToInt32(Date);//obtiene el dia actual en numero

            hora = DateTime.Now.ToShortTimeString();//obtiene la hora actual

            if(hora=="24:00" || hora == "00:00")
            {
                Leer_CH(); //permite leer el metodo de CH
            }


          Leer_RD();
        }
        #endregion

        public Service1()
        {
            InitializeComponent();
        }
     
        protected override void OnStart(string[] args)
        {
            DataTable dt1 = new DataTable();
            ClassData obj = new ClassData();
            obj.Obtener_Parametros();
            dt1 = obj.Obtener_Parametros();
            Int32 Tiempo_RD = Convert.ToInt32(dt1.Rows[0][0].ToString());
            Int32 Tiempo_CH = Convert.ToInt32(dt1.Rows[1][0].ToString());
            Ruta_Excel = (dt1.Rows[2][0].ToString());
            files = Directory.GetFiles(Ruta_Excel); // Obtener archivo de una carpeta
            indice = files[0].IndexOf("\\", 4);

            Nombre_Archivo = files[0].Substring(indice, files[0].Length - indice);
            Nombre_Archivo = Nombre_Archivo.Substring(1, Nombre_Archivo.Length - 1);
            string Date = DateTime.Now.ToString("dd");
            Hoja = Convert.ToInt32(Date);


            hora = DateTime.Now.ToShortTimeString();//obtiene la hora actual

            if (hora == "24:00" || hora == "00:00")
            {
                Leer_CH(); //permite leer el metodo de CH
            }


            timerRD.Elapsed += new ElapsedEventHandler(OnElapsedTimeRD);
            this.timerRD.AutoReset = true;
            timerRD.Interval = Tiempo_RD;  // Tiempo_RD; 
            timerRD.Enabled = true;
            

        }



        protected override void OnStop()
        {
            this.timerRD.Stop();
            this.timerRD = null;

          
            
        }
      
        private void OnElapsedTimeRD(object source, ElapsedEventArgs e)
        {
            Leer_RD();
           

        }
        public void Leer_RD()
        {
            ClassData _Obj_Clase_Data = new ClassData();
            _Obj_Clase_Data.Leer_Excel_RD(Ruta_Excel, Nombre_Archivo, Hoja);

            

        }

        public void Leer_CH()
        {
            ClassData _Obj_Clase_Data = new ClassData();
            _Obj_Clase_Data.Leer_Excel_CH(Ruta_Excel, Nombre_Archivo_ch, Hoja);



        }


        public void Log(String mensaje)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(mensaje);


            // flush every 20 seconds as you do it
            File.AppendAllText(@"c:\\log.txt", sb.ToString());
            sb.Clear();
        }

    }
}

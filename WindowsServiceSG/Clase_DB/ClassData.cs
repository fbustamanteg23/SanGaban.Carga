using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

using Excel = Microsoft.Office.Interop.Excel;

//using Microsoft.Office.Interop.Excel;


namespace Clase_DB
{
    public class ClassData
    {


        public void Log(String mensaje)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(mensaje);


            // flush every 20 seconds as you do it
            File.AppendAllText(@"c:\\log.txt", sb.ToString());
            sb.Clear();
        }

        public DataTable Obtener_Parametros()
        {

            DataBase connection = new DataBase();
            SqlConnection conn1 = new SqlConnection(connection.Conexion());
            conn1.Open();
            SqlCommand cmd1 = new SqlCommand();
            cmd1.Connection = conn1;
            cmd1.CommandType = CommandType.StoredProcedure;


            cmd1.CommandText = "PRI_SP_EXTRAER_CONFIGURACION";
            SqlDataReader dr = cmd1.ExecuteReader();
            var tb = new DataTable();
            tb.Load(dr);

            return tb;

        }

        public Boolean Leer_Excel_RD(String ruta, String Nombre_Archivo, int Hoja)
        {


            var connection = System.Configuration.ConfigurationManager.ConnectionStrings["bd"].ConnectionString;
            SqlConnection conn = new SqlConnection(connection);
            conn.Open();




            int _fila = 7;

            String Reporte_Cabecera = "";
            String Tipo_Medida = "";
            String dato = "";

            string filepath = ruta + Nombre_Archivo;



            try
            {
                //for (int _num_hoja = 1; _num_hoja <= 10; _num_hoja++)
                //{
                //    String _n_hoja = "";
                //    if (_num_hoja < 10)
                //    {
                //        _n_hoja = "0" + _num_hoja;
                //    }

                    Excel.Application excelApp = new Excel.Application();
                if (excelApp != null)
                {


                   

                        Log("inicio ruta");

                        Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filepath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    //Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[Hoja];

                   


                        Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[09];
                        Log("abrio ruta" + filepath);
                        Excel.Range excelRange = excelWorksheet.UsedRange;
                        //-----------------ELIMINA LOS REGISTRSOS DE LA TABLA MEDIDA DETALLE----------
                        SqlCommand cmd5 = new SqlCommand();
                        cmd5.Connection = conn;
                        cmd5.CommandType = CommandType.StoredProcedure;

                        cmd5.CommandText = "LIMPIAR_REPORTE_MEDIDA_DETALLE";

                        cmd5.ExecuteNonQuery();

                        //--------grabando en la tabla - reporte_medida_detalle-----------------

                        String fecha_registro = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                        SqlCommand cmd2 = new SqlCommand();
                        cmd2.Connection = conn;
                        cmd2.CommandType = CommandType.StoredProcedure;

                        cmd2.Parameters.Add(new SqlParameter("@fecha", SqlDbType.VarChar)).Value = fecha_registro;
                        cmd2.Parameters.Add(new SqlParameter("@tip_reporte", SqlDbType.VarChar)).Value = "RD";

                        cmd2.CommandText = "PRI_SP_INSERTAR_REPORTE_DETALLE";

                        cmd2.ExecuteNonQuery();




                        //-------------- Capturando los datos de la cabecera -------------------
                        //--------------recorriendo las columnas--------------------------------------
                        for (int col = 3; col <= 34; col++)
                        {
                            for (int fil = 4; fil == 4; fil++)
                            {

                                dato = (String)(excelWorksheet.Cells[fil, col] as Excel.Range).Text;

                                if (col >= 7 && col <= 19 && dato != "")
                                {
                                    fil = fil + 1;
                                    Reporte_Cabecera = (String)(excelWorksheet.Cells[fil, col] as Excel.Range).Text;
                                    fil = fil - 1;
                                }
                                else
                                {
                                    if (dato != "")
                                        Reporte_Cabecera = dato;

                                }

                                if (col == 13 || col == 11 || col == 12 || col == 16 || col == 18 || col == 29 || col == 31)
                                {
                                    fil = fil + 1;
                                    Reporte_Cabecera = (String)(excelWorksheet.Cells[fil, col] as Excel.Range).Text;
                                    fil = fil - 1;
                                }

                                Tipo_Medida = (String)(excelWorksheet.Cells[fil + 2, col] as Excel.Range).Text;

                                //--------------------------Validando la cabecera y tipo_medida--------------------------
                                DataTable dt_datos = new DataTable();
                                int respuesta = 0;


                                SqlCommand cmd = new SqlCommand();
                                cmd.Connection = conn;
                                cmd.CommandType = CommandType.StoredProcedure;

                                cmd.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = Regex.Replace(Reporte_Cabecera, @"\s", "");
                                cmd.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = Tipo_Medida;

                                cmd.CommandText = "PRI_SP_VALIDAR_DATOS";

                                cmd.ExecuteNonQuery();
                                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                                sda.Fill(dt_datos);


                                respuesta = Convert.ToInt32(dt_datos.Rows[0][0].ToString());
                                //---------------------------recorriendo las filas a grabar--------------- 
                                if (respuesta == 1)
                                {


                                    if (col >= 29 && col <= 33)

                                    {


                                        for (_fila = 7; _fila <= 103; _fila++)
                                        {
                                            String Hora = "";
                                            Decimal _valor_medida;
                                            //String prueba = (excelWorksheet.Cells[_fila, col] as Excel.Range).Text;

                                            if ((excelWorksheet.Cells[_fila, col] as Excel.Range).Text != null)
                                            {
                                                try
                                                {


                                                    Hora = (String)(excelWorksheet.Cells[_fila, 28] as Excel.Range).Text;
                                                    _valor_medida = Convert.ToDecimal((excelWorksheet.Cells[_fila, col] as Excel.Range).Text);
                                                    SqlCommand cmd3 = new SqlCommand();
                                                    cmd3.Connection = conn;

                                                    cmd3.CommandType = CommandType.StoredProcedure;
                                                    if (Hora == "24:00")
                                                    {
                                                        Hora = "00:00";
                                                    }
                                                    cmd3.Parameters.Add(new SqlParameter("@hora", SqlDbType.DateTime)).Value = Hora;
                                                    cmd3.Parameters.Add(new SqlParameter("@valor", SqlDbType.Decimal)).Value = _valor_medida;
                                                    cmd3.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = Regex.Replace(Reporte_Cabecera, @"\s", "");
                                                    cmd3.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = Tipo_Medida;
                                                    cmd3.CommandText = "PRI_SP_INSERTAR_REPORTE_MEDIDA_DETALLE";

                                                    cmd3.ExecuteNonQuery();

                                                }
                                                catch (Exception ex) { }



                                            }
                                            else
                                            {
                                                break;

                                            }
                                        }



                                    }
                                    else
                                    {
                                        if (col == 7 || col == 8 || col == 9 || col == 10) { }
                                        else
                                        {
                                            for (_fila = 7; _fila <= 54; _fila++)
                                            {
                                                String Hora = "";
                                                Decimal _valor_medida;
                                                //  String prueba = (excelWorksheet.Cells[_fila, col] as Excel.Range).Text;

                                                if ((excelWorksheet.Cells[_fila, col] as Excel.Range).Text != null)
                                                {
                                                    try
                                                    {


                                                        Hora = (String)(excelWorksheet.Cells[_fila, 2] as Excel.Range).Text;
                                                        _valor_medida = Convert.ToDecimal((excelWorksheet.Cells[_fila, col] as Excel.Range).Text);
                                                        SqlCommand cmd3 = new SqlCommand();
                                                        cmd3.Connection = conn;

                                                        cmd3.CommandType = CommandType.StoredProcedure;

                                                        cmd3.Parameters.Add(new SqlParameter("@hora", SqlDbType.DateTime)).Value = Hora;
                                                        cmd3.Parameters.Add(new SqlParameter("@valor", SqlDbType.Decimal)).Value = _valor_medida;
                                                        cmd3.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = Regex.Replace(Reporte_Cabecera, @"\s", "");
                                                        cmd3.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = Tipo_Medida;
                                                        cmd3.CommandText = "PRI_SP_INSERTAR_REPORTE_MEDIDA_DETALLE";

                                                        cmd3.ExecuteNonQuery();


                                                    }
                                                    catch (Exception ex) { }



                                                }
                                                else
                                                {
                                                    break;

                                                }
                                            }
                                        }
                                    }
                                }
                                //-----------------------------------------------------------------------



                            }
                        }



                        //excelWorkbook.Close();
                        //excelApp.Quit();


                        excelWorkbook.Close(false, filepath, System.Type.Missing);
                        excelApp.Quit();
                        excelWorkbook = null;
                        excelApp = null;



                    }



                //}

            }
            catch (Exception error89)
            {
                Log(error89.ToString());
            }

            return true;
        }



        public Boolean Leer_Excel_CH(String ruta, String Nombre_Archivo, int Hoja)
        {

            var connection = System.Configuration.ConfigurationManager.ConnectionStrings["bd"].ConnectionString;
            SqlConnection conn = new SqlConnection(connection);
            conn.Open();





            string filepath = ruta + Nombre_Archivo;



            try
            {

                Excel.Application excelApp = new Excel.Application();
                if (excelApp != null)
                {

                    String _Cabecera = "";
                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filepath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    //Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[Hoja];
                    Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[09]; //numero de hoja prueba

                    Excel.Range excelRange = excelWorksheet.UsedRange;

                    String fecha_registro_mod = DateTime.Now.ToString("yyyy-MM-dd");
                    //--------grabando en la tabla - reporte_medida_detalle-----------------

                    String fecha_registro = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    //DateTime nuevaFecha = Convert.ToDateTime(fecha_registro);
                    //nuevaFecha = nuevaFecha.AddDays(-1);
                    //fecha_registro = nuevaFecha.ToString("yyyy-MM-dd hh:mm:ss");
                    SqlCommand cmd10 = new SqlCommand();
                    cmd10.Connection = conn;
                    cmd10.CommandType = CommandType.StoredProcedure;

                    cmd10.Parameters.Add(new SqlParameter("@fecha", SqlDbType.VarChar)).Value = fecha_registro;
                    cmd10.Parameters.Add(new SqlParameter("@tip_reporte", SqlDbType.VarChar)).Value = "CH";

                    cmd10.CommandText = "PRI_SP_INSERTAR_REPORTE_DETALLE";

                    cmd10.ExecuteNonQuery();
                    //--------------VALIDANDO LA CABECERA Y LA MEDIDA----------------------------
                    int _fil1 = 9;
                    int _col_ava = 2;
                    for (int _col1 = 4; _col1 <= 20; _col1 = _col1 + _col_ava)
                    {

                        if (_col1 == 12)
                        {
                            _col_ava = 1;
                        }
                        
                        if (_col1 == 14 || _col1 == 13 || _col1 == 12)
                        {
                            _Cabecera = (String)(excelWorksheet.Cells[_fil1, 12] as Excel.Range).Text;
                        }
                        else
                        {
                            _Cabecera = (String)(excelWorksheet.Cells[_fil1, _col1] as Excel.Range).Text;
                        }

                        if (_col1 == 16)
                        {
                            _Cabecera = (String)(excelWorksheet.Cells[_fil1, 15] as Excel.Range).Text;
                        }

                        String _Medida = (String)(excelWorksheet.Cells[_fil1 + 1, _col1] as Excel.Range).Text;

                        DataTable dt_datos_v = new DataTable();
                        int respuesta = 0;

                        if (_Cabecera != "")
                        {
                            SqlCommand cmd = new SqlCommand();
                            cmd.Connection = conn;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _Cabecera;
                            cmd.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = _Medida;

                            cmd.CommandText = "PRI_SP_VALIDAR_CH";

                            cmd.ExecuteNonQuery();
                            SqlDataAdapter sda = new SqlDataAdapter(cmd);
                            sda.Fill(dt_datos_v);


                            respuesta = Convert.ToInt32(dt_datos_v.Rows[0][0].ToString());
                        }


                        if (respuesta == 1)
                        {
                            for (int _fil_2 = 11; _fil_2 <= 28; _fil_2++)
                            {
                                try
                                {



                                    String _cab_1 = (String)(excelWorksheet.Cells[_fil_2, 2] as Excel.Range).Text;
                                    if (_fil_2 == 16 || _fil_2 == 17 || _fil_2 == 18)
                                    {
                                        _cab_1 = (String)(excelWorksheet.Cells[15, 2] as Excel.Range).Text;
                                    }
                                    if (_fil_2 == 20)
                                    {
                                        _cab_1 = (String)(excelWorksheet.Cells[19, 2] as Excel.Range).Text;
                                    }
                                    if (_fil_2 == 23)
                                    {
                                        _cab_1 = (String)(excelWorksheet.Cells[22, 2] as Excel.Range).Text;
                                    }
                                    if (_fil_2 == 26)
                                    {
                                        _cab_1 = (String)(excelWorksheet.Cells[25, 2] as Excel.Range).Text;
                                    }

                                    String _Reporte_fila = _cab_1 + " " + (String)(excelWorksheet.Cells[_fil_2, 3] as Excel.Range).Text;

                                    String _valor_fila = (String)(excelWorksheet.Cells[_fil_2, _col1] as Excel.Range).Text;
                                    String _Hora_fila = (String)(excelWorksheet.Cells[_fil_2, _col1 + 1] as Excel.Range).Text;

                                    if (_col1 >= 12 && _col1 <= 19)
                                    {
                                        _Hora_fila = "00:00";
                                    }




                                    SqlCommand cmd90 = new SqlCommand();
                                    cmd90.Connection = conn;
                                    cmd90.CommandType = CommandType.StoredProcedure;






                                    if (_Cabecera == "HORAS DE OPERACIÓN")
                                    {
                                        cmd90.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = "libre";
                                        cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = "0.00";
                                        if (_valor_fila=="24:00:00")
                                        {
                                            cmd90.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " +  "00:00";
                                        }
                                        else
                                        {
                                            cmd90.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " + _valor_fila.Replace("%", "").Trim();
                                        }

                                        
                                        cmd90.Parameters.Add(new SqlParameter("@FILA", SqlDbType.VarChar)).Value = Regex.Replace(_Reporte_fila, @"\s", "");
                                        cmd90.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _Cabecera;
                                        cmd90.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";

                                        cmd90.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        if (_Cabecera == "DISPONIBILIDAD")
                                        {
                                            cmd90.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = "libre";


                                            if (_valor_fila == "" || _valor_fila is null)
                                            {
                                                cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = null;
                                            }
                                            else
                                            {
                                                cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _valor_fila.Replace("%", "").Trim();

                                            }

                                            //cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _valor_fila.Replace("%", "").Trim();
                                            cmd90.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " +  "00:00";
                                            cmd90.Parameters.Add(new SqlParameter("@FILA", SqlDbType.VarChar)).Value = Regex.Replace(_Reporte_fila, @"\s", "");
                                            cmd90.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _Cabecera;
                                            cmd90.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";

                                            cmd90.ExecuteNonQuery();
                                        }
                                        else
                                        {
                                            if (_Cabecera == "OBSERVACIONES")
                                            {
                                                cmd90.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = "libre";

                                                if (_valor_fila=="" || _valor_fila is null)
                                                {
                                                    cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = null;
                                                    cmd90.Parameters.Add(new SqlParameter("@OBS", SqlDbType.VarChar)).Value = null;
                                                }
                                                else
                                                {
                                                    cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _valor_fila.Replace("%", "").Trim();
                                                    cmd90.Parameters.Add(new SqlParameter("@OBS", SqlDbType.VarChar)).Value = _valor_fila.Replace("%", "").Trim();
                                                }

                                               
                                                cmd90.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " +  "00:00";
                                                cmd90.Parameters.Add(new SqlParameter("@FILA", SqlDbType.VarChar)).Value = Regex.Replace(_Reporte_fila, @"\s", "");
                                                cmd90.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _Cabecera;
                                                //cmd90.Parameters.Add(new SqlParameter("@OBS", SqlDbType.VarChar)).Value = _valor_fila.Replace("%", "").Trim();
                                                cmd90.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";

                                                cmd90.ExecuteNonQuery();
                                            }
                                            else
                                            {
                                                
                                                cmd90.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = _Medida.Trim();

                                                if(_valor_fila=="" || _valor_fila is null)
                                                {
                                                    cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = null;
                                                }
                                                else
                                                {
                                                    cmd90.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _valor_fila.Replace("%", "").Replace("--","-").Trim();

                                                }
                                              
                                                if(_Hora_fila== "#¡REF!")
                                                {
                                                    _Hora_fila = "00:00";
                                                }
                                                
                                                
                                                cmd90.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " "+_Hora_fila;
                                                cmd90.Parameters.Add(new SqlParameter("@FILA", SqlDbType.VarChar)).Value = Regex.Replace(_Reporte_fila, @"\s", "");
                                                cmd90.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _Cabecera;

                                                cmd90.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";

                                                cmd90.ExecuteNonQuery();
                                            }
                                        }
                                    }

                                }
                                catch (Exception error) {

                                    


                                }
                            }

                        }



                    }



                    //------------------------------------------------------------------------//

                    for (int _ff_2 = 31; _ff_2 <= 31; _ff_2++)
                    {
                        String _cc_cabecera_3 = (String)(excelWorksheet.Cells[_ff_2, 2] as Excel.Range).Text;
                        String _ff_valor_3 = (String)(excelWorksheet.Cells[_ff_2, 4] as Excel.Range).Text;
                        SqlCommand cmd91 = new SqlCommand();
                        cmd91.Connection = conn;
                        cmd91.CommandType = CommandType.StoredProcedure;
                        try
                        {

                            cmd91.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = "libre";
                            cmd91.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _ff_valor_3.Replace("%", "").Trim();
                            cmd91.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " + "00:00";                         
                            cmd91.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _cc_cabecera_3.Replace(" ", "").Trim();
                
                            cmd91.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";

                            cmd91.ExecuteNonQuery();
                           
                        }
                        catch (Exception err99) { }

                    }


                    //------------------carga cuadros----------------------------------------
                    for (int __col = 6; __col <= 9; __col++)
                    {
                        for (int _ff_3 = 30; _ff_3 <= 30; _ff_3++)
                        {
                            String _cc_cabecera_3 = (String)(excelWorksheet.Cells[30, 6] as Excel.Range).Text;
                            String _tipo_medida = (String)(excelWorksheet.Cells[_ff_3+1, 7] as Excel.Range).Text;
                            String _fila = (String)(excelWorksheet.Cells[_ff_3 + 2, __col + 1] as Excel.Range).Text;
                            String _Hora = (String)(excelWorksheet.Cells[32, 6] as Excel.Range).Text;
                            SqlCommand cmd96 = new SqlCommand();
                            cmd96.Connection = conn;
                            cmd96.CommandType = CommandType.StoredProcedure;
                            try
                            {

                                cmd96.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = _tipo_medida;
                                cmd96.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _fila.Replace("%", "").Trim();
                                cmd96.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " + _Hora;
                                cmd96.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _cc_cabecera_3.Replace(" ", "").Trim();

                           
                                cmd96.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";
                                cmd96.ExecuteNonQuery();
                            }
                            catch (Exception err96) { }

                        }
                    }

                    for (int __col = 6; __col <= 9; __col++)
                    {
                        for (int _ff_3 = 30; _ff_3 <= 30; _ff_3++)
                        {
                            String _cc_cabecera_3 = (String)(excelWorksheet.Cells[30, 6] as Excel.Range).Text;
                            String _tipo_medida = (String)(excelWorksheet.Cells[_ff_3+1, 7] as Excel.Range).Text;
                            String _fila = (String)(excelWorksheet.Cells[_ff_3 + 3, __col + 1] as Excel.Range).Text;
                            String _Hora = (String)(excelWorksheet.Cells[32, 6] as Excel.Range).Text;
                            SqlCommand cmd98 = new SqlCommand();
                            cmd98.Connection = conn;
                            cmd98.CommandType = CommandType.StoredProcedure;
                            try
                            {

                                cmd98.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = _tipo_medida;
                                cmd98.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _fila.Replace("%", "").Trim();
                                cmd98.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " + _Hora;
                                cmd98.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _cc_cabecera_3.Replace(" ", "").Trim();


                            
                                cmd98.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";
                                cmd98.ExecuteNonQuery();
                            }
                            catch (Exception err98) { }

                        }
                    }
                    //----------------------------------------------------------------------------------


                    for (int __col = 15; __col <= 16; __col++)
                    {
                        for (int _ff_3 = 31; _ff_3 <= 31; _ff_3++)
                        {
                            String _cc_cabecera_3 = (String)(excelWorksheet.Cells[30, 14] as Excel.Range).Text;
                            String _tipo_medida = (String)(excelWorksheet.Cells[_ff_3, __col] as Excel.Range).Text;
                        
                            String _fila = (String)(excelWorksheet.Cells[_ff_3 + 1, __col] as Excel.Range).Text;

                            SqlCommand cmd96 = new SqlCommand();
                            cmd96.Connection = conn;
                            cmd96.CommandType = CommandType.StoredProcedure;
                            try
                            {
                                cmd96.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = _tipo_medida;
                                cmd96.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _fila.Replace("%", "").Trim();
                                cmd96.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " + "23:00";
                                cmd96.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _cc_cabecera_3.Replace(" ", "").Trim();
                                
                                cmd96.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";
                                cmd96.ExecuteNonQuery();
                            }
                            catch (Exception err96) { }

                        }
                    }
                    //----------------------------------------------------------------------------------


                    for (int __col = 15; __col <= 16; __col++)
                    {
                        for (int _ff_3 = 31; _ff_3 <= 31; _ff_3++)
                        {
                            String _cc_cabecera_3 = (String)(excelWorksheet.Cells[30, 14] as Excel.Range).Text;
                            String _tipo_medida = (String)(excelWorksheet.Cells[_ff_3, __col] as Excel.Range).Text;
                            String _fila = (String)(excelWorksheet.Cells[_ff_3 + 2, __col] as Excel.Range).Text;
                            SqlCommand cmd99 = new SqlCommand();
                            cmd99.Connection = conn;
                            cmd99.CommandType = CommandType.StoredProcedure;
                            try
                            {
                                cmd99.Parameters.Add(new SqlParameter("@MEDIDA", SqlDbType.VarChar)).Value = _tipo_medida;
                                cmd99.Parameters.Add(new SqlParameter("@VALOR", SqlDbType.VarChar)).Value = _fila.Replace("%", "").Trim();
                                cmd99.Parameters.Add(new SqlParameter("@HORA", SqlDbType.VarChar)).Value = fecha_registro_mod + " " + "00:00";
                                cmd99.Parameters.Add(new SqlParameter("@CABECERA", SqlDbType.VarChar)).Value = _cc_cabecera_3.Replace(" ", "").Trim();
                                cmd99.CommandText = "PRI_SP_REPORTE_MEDIDA_DETALLE_CH";
                                cmd99.ExecuteNonQuery();
                            }
                            catch (Exception err99) { }

                        }
                    }



                    //-------------------------------------------------------------------------
                    excelWorkbook.Close(false, filepath, System.Type.Missing);
                    excelApp.Quit();
                    excelWorkbook = null;
                    excelApp = null;



                }

            }
            catch (Exception error89)
            {
                Log(error89.ToString());
            }

            return true;
        }
    }
}

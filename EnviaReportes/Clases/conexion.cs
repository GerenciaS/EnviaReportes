using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
namespace EnviaReportes
{
    class conexion
    {
        public static SqlConnection conectar(string bd) 
        {
            try
            {
                SqlConnection cn = new SqlConnection();

                cn.ConnectionString = "Server=server-cln;Database=" + bd + ";User Id=prgrmusr;Password=$M3rc4d0$;";

               // cn.ConnectionString = "Server=sqltest;Database=" + bd + ";User Id=sa;Password=03170754;";
                              
                cn.Open();
                return cn;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
    }
}

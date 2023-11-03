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
                //cn.ConnectionString = "Server=187.216.118.170,14335;Database=" + bd + ";User Id=prgrmusr;Password=$M3rc4d0$;";
                //cn.ConnectionString = "Server=server-cln;Database=" + bd + ";User Id=prgrmusr;Password=$M3rc4d0$;";

                cn.ConnectionString = "Server=sqltest;Database=" + bd + ";User Id=sa;Password=03170754;";

                //cn.ConnectionString = "Server=187.216.118.170,14335;Database=" + bd + ";User Id=eliasb;Password=23032004;";

                //cn.ConnectionString = "Server=server-cln,1433;Database=" + bd + ";User Id=eliasb;Password=23032004;Connection Timeout=0"; 
                //cn.ConnectionString = "Server=server-cln,1433;Database=" + bd + ";User Id=jcbeltran;Password=010285;Connection Timeout=0";  

                //cn.ConnectionString = "Server=(local),1433;Database=" + bd + ";User Id=sa;Password=kasuko!*suc;";                 
                //cn.ConnectionString = "Server=servercorporativo.myvnc.com,14335;Database=" + bd + ";User Id=eliasb;Password=23032004;";                 
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

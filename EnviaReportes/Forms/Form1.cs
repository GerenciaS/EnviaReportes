using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
namespace EnviaReportes
{
    public partial class FrmPrimcipal : Form
    {

        int diferencia = 0;
        DateTime fi_anterior;
        DateTime ff_anterior;
        List<string> emailList = new List<string>();

        public FrmPrimcipal()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(DateTime.Now.ToString("dd-MM-yyyy"));
        }

        private string ReporteTemporadaNavideña()
        {
            SqlConnection cn = conexion.conectar("BDIntegrador");
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandTimeout = 240;
            System.Data.DataTable dt = new System.Data.DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            DateTime dia = DateTime.Now.AddDays(-1);

            DateTime primero = dia.AddDays((dia.Day - 1) * -1);
            //DateTime primero=Convert.ToDateTime("2016-11-01");
            DateTime diabase = dia.AddYears(-1);
            DateTime primerobase = primero.AddYears(-1);
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //
            //                                          VENTA DIARIA
            //
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            cmd.CommandText = "select 	Sucursales.cod_estab,ltrim(rtrim(Sucursales.nombre)) as nombre,isnull(BlisterDiario.unidades,0) as BlisterDiarioUni,isnull(BlisterAcu.unidades,0) as BlisterAcuUni,"
            + " isnull((BlisterDiario.unidades/(select sum(cantidad) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='28' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(BlisterDiario.VentaNeta,0) as BlisterDiarioVta,isnull(BlisterAcu.VentaNeta,0) as BlisterAcuVta,isnull((BlisterDiario.VentaNeta/(select sum(VentaNeta) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='28' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(BlisterDiario.UtilBruta,0) as BlisterDiarioUtil,isnull(BlisterAcu.UtilBruta,0) as BlisterAcuUtil,"
            + " isnull((BlisterDiario.UtilBruta/(select sum(va.VentaNetaSinIva)-(sum(cantidad*isnull(cedis.ultimo_costo,1))) as unidades from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='28' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "') left join (select cod_prod,ultimo_costo from BMSEPM_CEDIS..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod ))*100,0) as '%',"
            + " isnull(BurbujaDiario.unidades,0) as BurbujaDiarioUni,isnull(BurbujaAcu.unidades,0) as BurbujaAcuUni,"
            + " isnull((BurbujaDiario.unidades/(select sum(cantidad) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='29' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(BurbujaDiario.VentaNeta,0) as BurbujaDiarioVta,isnull(BurbujaAcu.VentaNeta,0) as BurbujaAcuVta,"
            + " isnull((BurbujaDiario.VentaNeta/(select sum(VentaNeta) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='29' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(BurbujaDiario.UtilBruta,0) as BurbujaDiarioUtil,	isnull(BurbujaAcu.UtilBruta,0) as BurbujaAcuUtil,"
            + " isnull((BurbujaDiario.UtilBruta/(select sum(va.VentaNetaSinIva)-(sum(cantidad*isnull(cedis.ultimo_costo,1))) as unidades from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='29' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "') left join (select cod_prod,ultimo_costo from BMSEPM_CEDIS..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod ))*100,0) as '%',"
            + " isnull(CajaDiario.unidades,0) as CajaDiarioUni,isnull(CajaAcu.unidades,0) as CajaAcuUni,isnull((CajaDiario.unidades/(select sum(cantidad) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='30' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(CajaDiario.VentaNeta,0) as CajaDiarioVta,	isnull(CajaAcu.VentaNeta,0) as CajaAcuVta,"
            + " isnull((CajaDiario.VentaNeta/(select sum(VentaNeta) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='30' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(CajaDiario.UtilBruta,0) as CajaDiarioUtil,isnull(CajaAcu.UtilBruta,0) as CajaAcuUtil,isnull((CajaDiario.UtilBruta/(select sum(va.VentaNetaSinIva)-(sum(cantidad*isnull(cedis.ultimo_costo,1))) as unidades from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='30' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "') left join (select cod_prod,ultimo_costo from BMSEPM_CEDIS..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod ))*100,0) as '%',"
            + " isnull(LuzDiario.unidades,0) as LuzDiarioUni,isnull(LuzAcu.unidades,0) as LuzAcuUni,isnull((LuzDiario.unidades/(select sum(cantidad) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='33' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(LuzDiario.VentaNeta,0) as LuzDiarioVta,isnull(LuzAcu.VentaNeta,0) as LuzAcuVta,isnull((LuzDiario.VentaNeta/(select sum(VentaNeta) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='33' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(LuzDiario.UtilBruta,0) as LuzDiarioUtil,isnull(LuzAcu.UtilBruta,0) as LuzAcuUtil,isnull((LuzDiario.UtilBruta/(select sum(va.VentaNetaSinIva)-(sum(cantidad*isnull(cedis.ultimo_costo,1))) as unidades from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='33' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "') left join (select cod_prod,ultimo_costo from BMSEPM_CEDIS..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod ))*100,0) as '%',"
            + " isnull(NavideñoDiario.unidades,0) as NaviDiarioUni,isnull(NavideñoAcu.unidades,0) as NaviAcuUni,isnull((NavideñoDiario.unidades/(select sum(cantidad) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='35' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(NavideñoDiario.VentaNeta,0) as NaviDiarioVta,isnull(NavideñoAcu.VentaNeta,0) as NaviAcuVta,isnull((NavideñoDiario.VentaNeta/(select sum(VentaNeta) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='35' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(NavideñoDiario.UtilBruta,0) as NaviDiarioUtil,isnull(NavideñoAcu.UtilBruta,0) as NaviAcuUtil,isnull((NavideñoDiario.UtilBruta/(select sum(va.VentaNetaSinIva)-(sum(cantidad*isnull(cedis.ultimo_costo,1))) as unidades from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='35' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "') left join (select cod_prod,ultimo_costo from BMSEPM_CEDIS..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod ))*100,0) as '%',"
            + " isnull(PinoDiario.unidades,0) as PinoDiarioUni,isnull(PinoAcu.unidades,0) as PinoAcuUni,isnull((PinoDiario.unidades/(select sum(cantidad) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='38' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(PinoDiario.VentaNeta,0) as PinoDiarioVta,isnull(PinoAcu.VentaNeta,0) as PinoAcuVta,isnull((PinoDiario.VentaNeta/(select sum(VentaNeta) as unidades from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='38' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "'))*100,0) as '%',"
            + " isnull(PinoDiario.UtilBruta,0) as PinoDiarioUtil,isnull(PinoAcu.UtilBruta,0) as PinoAcuUtil,isnull((PinoDiario.UtilBruta/(select sum(va.VentaNetaSinIva)-(sum(cantidad*isnull(cedis.ultimo_costo,1))) as unidades from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod and p.familia='38' and cod_estab not in ('1','1001','1002','1003','1004','67') and fecha='" + dia.ToString("yyyyMMdd") + "') left join (select cod_prod,ultimo_costo from BMSEPM_CEDIS..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod ))*100,0) as '%'"
            + " from ((((((((((((select cod_estab,nombre from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','67')) as Sucursales left join"
            + " (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod"
            + " where p.familia='28' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BlisterDiario on Sucursales.cod_estab=BlisterDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod )"
            + " left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='28' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BlisterAcu on Sucursales.cod_estab=BlisterAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='29' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BurbujaDiario on Sucursales.cod_estab=BurbujaDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='29' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BurbujaAcu on Sucursales.cod_estab=BurbujaAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='30' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as CajaDiario on Sucursales.cod_estab=CajaDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='30' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as CajaAcu on Sucursales.cod_estab=CajaAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='33' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as LuzDiario on Sucursales.cod_estab=LuzDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='33' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as LuzAcu on Sucursales.cod_estab=LuzAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='35' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as NavideñoDiario on Sucursales.cod_estab=NavideñoDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='35' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as NavideñoAcu on Sucursales.cod_estab=NavideñoAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='38' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as PinoDiario on Sucursales.cod_estab=PinoDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='38' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as PinoAcu on Sucursales.cod_estab=pinoAcu.cod_estab order by CAST(Sucursales.cod_estab as int) asc";
            da.Fill(dt);
            DG1.DataSource = dt;
            DG1.SelectAll();
            object objeto = DG1.GetClipboardContent();
            Microsoft.Office.Interop.Excel.Application excel;
            excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook libro;
            libro = excel.Workbooks.Add();
            libro.Worksheets.Add();
            libro.Worksheets.Add();
            libro.Worksheets.Add();
            Worksheet hoja = new Worksheet();
            hoja = (Worksheet)libro.Worksheets.get_Item(1);
            hoja.Name = "VENTA DIARIA";
            Microsoft.Office.Interop.Excel.Range rango;
            if (objeto != null)
            {
                Clipboard.SetDataObject(objeto);
                hoja.Cells[1, 2] = "REPORTE DE VENTA DIARIA";
                //ENCABEZADO VENTA DIARIA
                rango = (Range)hoja.get_Range("B1", "BE2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango = (Range)hoja.get_Range("B4", "B6");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "CODIGO";
                rango = (Range)hoja.get_Range("C4", "C6");
                rango.Select();
                rango.Merge();
                //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "SUCURSAL";

                for (int i = 4; i <= 49; i += 9)
                {
                    //rango = (Range)hoja.get_Range("4,4","12,4");
                    rango = (Range)hoja.get_Range(sCol(i) + "4", sCol(i + 8) + "4");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    switch (i)
                    {
                        case 4:
                            rango.Cells[1, 1] = "B L I S T E R";
                            break;
                        case 13:
                            rango.Cells[1, 1] = "B U R B U J A";
                            break;
                        case 22:
                            rango.Cells[1, 1] = "C A J A";
                            break;
                        case 31:
                            rango.Cells[1, 1] = "L U Z";
                            break;
                        case 40:
                            rango.Cells[1, 1] = "N A V I D E Ñ O";
                            break;
                        case 49:
                            rango.Cells[1, 1] = "P I N O";
                            break;
                    }
                    for (int j = i; j <= i + 8; j += 3)
                    {
                        rango = (Range)hoja.get_Range(sCol(j) + "5", sCol(j + 2) + "5");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        if (j == i)
                        {
                            rango.Cells[1, 1] = "UNIDADES";
                        }
                        else if (j == i + 3)
                        {
                            rango.Cells[1, 1] = "VENTA NETA";
                        }
                        else if (j == i + 6)
                        {
                            rango.Cells[1, 1] = "UTILIDAD BRUTA";
                        }
                        for (int k = j; k <= j + 2; k++)
                        {
                            if (k == j)
                            {
                                hoja.Cells[6, k] = "DIARIAS";
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }

                            }
                            else if (k == j + 1)
                            {
                                hoja.Cells[6, k] = "ACUMULADAS";
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }
                            }
                            else if (k == j + 2)
                            {
                                hoja.Cells[6, k] = "%";
                                hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00";
                            }
                            hoja.Cells[6, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[6, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            hoja.Cells[6, k].Font.FontStyle = "Bold";
                            hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=SUM(" + sCol(k) + "7:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 6) + ")";
                        }
                    }
                }

                hoja.Range["B7", "C" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                hoja.Range["D7", "BE" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                hoja.Range["B7", "BE" + Convert.ToString(DG1.Rows.Count + 7)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                hoja.Cells[DG1.Rows.Count + 7, 3] = "T O T A L";
                rango = (Range)hoja.Cells[7, 1];
                rango.Select();
                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                rango = (Range)hoja.get_Range("A1", "BE" + Convert.ToString(DG1.Rows.Count + 10));
                rango.EntireColumn.AutoFit();
            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //
            //                                          COMPARATIVO DIARIO
            //
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Clipboard.Clear();
            objeto = null;
            DG1.DataSource = null;
            DG1.Rows.Clear();
            DG1.Columns.Clear();

            cmd.CommandText = "select Sucursales.cod_estab,Sucursales.nombre, isnull(BlisterDiario.unidades,0) as BlisterDiarioUni,	isnull(BlisterBase.unidades,0) as BlisterBaseUni,[Inc o Dec]=case"
            + " when BlisterBase.unidades=0 then 0 when BlisterBase.unidades is null then 0 when BlisterBase.unidades>0 then ((BlisterDiario.unidades/BlisterBase.unidades)-1)*100 end,isnull(BlisterDiario.VentaNeta,0) as BlisterDiarioVta,"
            + " isnull(BlisterBase.VentaNeta,0) as BlisterBaseVta,[Inc o Dec]=case when BlisterBase.VentaNeta=0 then 0 when BlisterBase.VentaNeta is null then 0	when BlisterBase.VentaNeta>0 then ((BlisterDiario.VentaNeta/BlisterBase.VentaNeta)-1)*100 end,"
            + " isnull(BlisterDiario.UtilBruta,0) as BlisterDiarioUtil,isnull(BlisterBase.UtilBruta,0) as BlisterBaseUtil,[Inc o Dec]=case when BlisterBase.UtilBruta=0 then 0 when BlisterBase.UtilBruta is null then 0 when BlisterBase.UtilBruta>0 then ((BlisterDiario.UtilBruta/BlisterBase.UtilBruta)-1)*100 end,"
            + " isnull(BurbujaDiario.unidades,0) as BurbujaDiarioUni,isnull(BurbujaBase.unidades,0) as BurbujaBaseUni,[Inc o Dec]=case when BurbujaBase.unidades=0 then 0 when BurbujaBase.unidades is null then 0 when BurbujaBase.unidades>0 then ((BurbujaDiario.unidades/BurbujaBase.unidades)-1)*100 end,"
            + " isnull(BurbujaDiario.VentaNeta,0) as BurbujaDiarioVta,isnull(BurbujaBase.VentaNeta,0) as BurbujaBaseVta,[Inc o Dec]=case when BurbujaBase.VentaNeta=0 then 0 when BurbujaBase.VentaNeta is null then 0 when BurbujaBase.VentaNeta>0 then ((BurbujaDiario.VentaNeta/BurbujaBase.VentaNeta)-1)*100	end,"
            + " isnull(BurbujaDiario.UtilBruta,0) as BurbujaDiarioUtil,isnull(BurbujaBase.UtilBruta,0) as BurbujaBaseUtil,[Inc o Dec]=case when BurbujaBase.UtilBruta=0 then 0 when BurbujaBase.UtilBruta is null then 0	when BurbujaBase.UtilBruta>0 then ((BurbujaDiario.UtilBruta/BurbujaBase.UtilBruta)-1)*100 end,"
            + " isnull(CajaDiario.unidades,0) as CajaDiarioUni,isnull(CajaBase.unidades,0) as CajaBaseUni,[Inc o Dec]=case when CajaBase.unidades=0 then 0 when CajaBase.unidades is null then 0 when CajaBase.unidades>0 then ((CajaDiario.unidades/CajaBase.unidades)-1)*100 end,"
            + " isnull(CajaDiario.VentaNeta,0) as CajaDiarioVta,isnull(CajaBase.VentaNeta,0) as CajaBaseVta,[Inc o Dec]=case when CajaBase.VentaNeta=0 then 0 when CajaBase.VentaNeta is null then 0	when CajaBase.VentaNeta>0 then ((CajaDiario.VentaNeta/CajaBase.VentaNeta)-1)*100 end,"
            + " isnull(CajaDiario.UtilBruta,0) as CajaDiarioUtil,isnull(CajaBase.UtilBruta,0) as CajaBaseUtil,[Inc o Dec]=case when CajaBase.UtilBruta=0 then 0 when CajaBase.UtilBruta is null then 0 when CajaBase.UtilBruta>0 then ((CajaDiario.UtilBruta/CajaBase.UtilBruta)-1)*100 end,"
            + " isnull(LuzDiario.unidades,0) as LuzDiarioUni,isnull(LuzBase.unidades,0) as LuzBaseUni,[Inc o Dec]=case when LuzBase.unidades=0 then 0 when LuzBase.unidades is null then 0 when LuzBase.unidades>0 then ((LuzDiario.unidades/LuzBase.unidades)-1)*100 end,"
            + " isnull(LuzDiario.VentaNeta,0) as LuzDiarioVta,isnull(LuzBase.VentaNeta,0) as LuzBaseVta,[Inc o Dec]=case	when LuzBase.VentaNeta=0 then 0	when LuzBase.VentaNeta is null then 0 when LuzBase.VentaNeta>0 then ((LuzDiario.VentaNeta/LuzBase.VentaNeta)-1)*100 end,"
            + " isnull(LuzDiario.UtilBruta,0) as LuzDiarioUtil,isnull(LuzBase.UtilBruta,0) as LuzBaseUtil,[Inc o Dec]=case when LuzBase.UtilBruta=0 then 0 when LuzBase.UtilBruta is null then 0 when LuzBase.UtilBruta>0 then ((LuzDiario.UtilBruta/LuzBase.UtilBruta)-1)*100 end,"
            + " isnull(NavideñoDiario.unidades,0) as NaviDiarioUni,isnull(NavideñoBase.unidades,0) as NaviBaseUni,[Inc o Dec]=case when NavideñoBase.unidades=0 then 0 when NavideñoBase.unidades is null then 0 when NavideñoBase.unidades>0 then ((NavideñoDiario.unidades/NavideñoBase.unidades)-1)*100 end,"
            + " isnull(NavideñoDiario.VentaNeta,0) as NaviDiarioVta,isnull(NavideñoBase.VentaNeta,0) as NaviBaseVta,[Inc o Dec]=case	when NavideñoBase.VentaNeta=0 then 0 when NavideñoBase.VentaNeta is null then 0	when NavideñoBase.VentaNeta>0 then ((NavideñoDiario.VentaNeta/NavideñoBase.VentaNeta)-1)*100 end,"
            + " isnull(NavideñoDiario.UtilBruta,0) as NaviDiarioUtil,isnull(NavideñoBase.UtilBruta,0) as NaviBaseUtil,[Inc o Dec]=case when NavideñoBase.UtilBruta=0 then 0 when NavideñoBase.UtilBruta is null then 0 when NavideñoBase.UtilBruta>0 then ((NavideñoDiario.UtilBruta/NavideñoBase.UtilBruta)-1)*100 end,"
            + " isnull(PinoDiario.unidades,0) as PinoDiarioUni,isnull(PinoBase.unidades,0) as PinoBaseUni,[Inc o Dec]=case when PinoBase.unidades=0 then 0 when PinoBase.unidades is null then 0 when PinoBase.unidades>0 then ((PinoDiario.unidades/PinoBase.unidades)-1)*100 end,"
            + " isnull(PinoDiario.VentaNeta,0) as PinoDiarioVta,isnull(PinoBase.VentaNeta,0) as PinoBaseVta,[Inc o Dec]=case when PinoBase.VentaNeta=0 then 0 when PinoBase.VentaNeta is null then 0	when PinoBase.VentaNeta>0 then ((PinoDiario.VentaNeta/PinoBase.VentaNeta)-1)*100 end,"
            + " isnull(PinoDiario.UtilBruta,0) as PinoDiarioUtil,isnull(PinoBase.UtilBruta,0) as PinoBaseUtil,[Inc o Dec]=case when PinoBase.UtilBruta=0 then 0 when PinoBase.UtilBruta is null then 0 when PinoBase.UtilBruta>0 then ((PinoDiario.UtilBruta/PinoBase.UtilBruta)-1)*100 end"
            + " from ((((((((((((select cod_estab,nombre from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','67')) as Sucursales left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='28' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BlisterDiario on Sucursales.cod_estab=BlisterDiario.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta,SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='28' and fecha='" + diabase.ToString("yyyyMMdd") + "' group by va.cod_estab) as BlisterBase on Sucursales.cod_estab=BlisterBase.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod"
            + " where p.familia='29' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BurbujaDiario on Sucursales.cod_estab=BurbujaDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod"
            + " where p.familia='29' and fecha='" + diabase.ToString("yyyyMMdd") + "' group by va.cod_estab) as BurbujaBase on Sucursales.cod_estab=BurbujaBase.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod )"
            + " left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='30' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as CajaDiario on Sucursales.cod_estab=CajaDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta"
            + " from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='30' and fecha='" + diabase.ToString("yyyyMMdd") + "' group by va.cod_estab) as CajaBase on Sucursales.cod_estab=CajaBase.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='33' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as LuzDiario on Sucursales.cod_estab=LuzDiario.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='33' and fecha='" + diabase.ToString("yyyyMMdd") + "' group by va.cod_estab) as LuzBase on Sucursales.cod_estab=LuzBase.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod"
            + " where p.familia='35' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as NavideñoDiario on Sucursales.cod_estab=NavideñoDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod"
            + " where p.familia='35' and fecha='" + diabase.ToString("yyyyMMdd") + "' group by va.cod_estab) as NavideñoBase on Sucursales.cod_estab=NavideñoBase.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod )"
            + " left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='38' and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as PinoDiario on Sucursales.cod_estab=PinoDiario.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta"
            + " from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='38' and fecha='" + diabase.ToString("yyyyMMdd") + "' group by va.cod_estab) as PinoBase on Sucursales.cod_estab=PinoBase.cod_estab order by CAST(Sucursales.cod_estab as int) asc";
            dt = new System.Data.DataTable();
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            da.Fill(dt);
            System.Windows.Forms.Application.DoEvents();
            DG1.DataSource = dt;
            DG1.SelectAll();
            objeto = DG1.GetClipboardContent();
            if (objeto != null)
            {

                Clipboard.SetDataObject(objeto);
                hoja = (Worksheet)libro.Sheets.get_Item(2);
                hoja.Activate();
                hoja.Name = "COMPARATIVO DIARIO";
                hoja.Cells[1, 2] = "COMPARATIVO DIARIO";
                //ENCABEZADO VENTA DIARIA
                rango = (Range)hoja.get_Range("B1", "BE2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango = (Range)hoja.get_Range("B3", "BE3");
                rango.Select();
                rango.Merge();
                rango.Cells[1.1, Type.Missing] = dia.ToString("d MMM yyyy");
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango = (Range)hoja.get_Range("B4", "B6");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "CODIGO";
                rango = (Range)hoja.get_Range("C4", "C6");
                rango.Select();
                rango.Merge();
                //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "SUCURSAL";
                for (int i = 4; i <= 49; i += 9)
                {
                    //rango = (Range)hoja.get_Range("4,4","12,4");
                    rango = (Range)hoja.get_Range(sCol(i) + "4", sCol(i + 8) + "4");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    switch (i)
                    {
                        case 4:
                            rango.Cells[1, 1] = "COMPARATIVO DIARIO BLISTER";
                            break;
                        case 13:
                            rango.Cells[1, 1] = "COMPARATIVO DIARIO BURBUJA";
                            break;
                        case 22:
                            rango.Cells[1, 1] = "COMPARATIVO DIARIO CAJA";
                            break;
                        case 31:
                            rango.Cells[1, 1] = "COMPARATIVO DIARIO LUZ";
                            break;
                        case 40:
                            rango.Cells[1, 1] = "COMPARATIVO DIARIO NAVIDEÑO";
                            break;
                        case 49:
                            rango.Cells[1, 1] = "COMPARATIVO DIARIO PINO";
                            break;
                    }
                    for (int j = i; j <= i + 8; j += 3)
                    {
                        rango = (Range)hoja.get_Range(sCol(j) + "5", sCol(j + 1) + "5");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        hoja.Cells[5, j + 2] = "%";
                        hoja.Cells[5, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        hoja.Cells[5, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        if (j == i)
                        {
                            rango.Cells[1, 1] = "UNIDADES";
                        }
                        else if (j == i + 3)
                        {
                            rango.Cells[1, 1] = "VENTA NETA";
                        }
                        else if (j == i + 6)
                        {
                            rango.Cells[1, 1] = "UTILIDAD BRUTA";
                        }
                        for (int k = j; k <= j + 2; k++)
                        {
                            if (k == j)
                            {
                                hoja.Cells[6, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                hoja.Cells[6, k].NumberFormat = "@";
                                hoja.Cells[6, k] = dia.Year.ToString();
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=SUM(" + sCol(k) + "7:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 6) + ")";
                            }
                            else if (k == j + 1)
                            {
                                hoja.Cells[6, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                hoja.Cells[6, k].NumberFormat = "@";
                                hoja.Cells[6, k] = diabase.Year.ToString();
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=SUM(" + sCol(k) + "7:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 6) + ")";
                            }
                            else if (k == j + 2)
                            {
                                hoja.Cells[6, k] = "Inc o Dec";
                                hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00";
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 7) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 7) + ")-1)*100";
                            }
                            hoja.Cells[6, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[6, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            hoja.Cells[6, k].Font.FontStyle = "Bold";

                        }
                    }
                }
                hoja.Range["B7", "C" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                hoja.Range["D7", "BE" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                hoja.Range["B7", "BE" + Convert.ToString(DG1.Rows.Count + 7)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                hoja.Cells[DG1.Rows.Count + 7, 3] = "T O T A L";
                rango = (Range)hoja.Cells[7, 1];
                rango.Select();
                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                rango = (Range)hoja.get_Range("A1", "BE" + Convert.ToString(DG1.Rows.Count + 10));
                rango.EntireColumn.AutoFit();
            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //
            //                                          COMPARATIVO ACUMULADO
            //
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Clipboard.Clear();
            objeto = null;
            DG1.DataSource = null;
            DG1.Rows.Clear();
            DG1.Columns.Clear();
            cmd.CommandText = "SELECT Sucursales.cod_estab,Sucursales.nombre,"
            + " isnull(BlisterAcu.unidades,0) as BlisterAcuUni,isnull(BlisterAcuBase.unidades,0) as BlisterBaseUni,[Inc o Dec]=case when BlisterAcuBase.unidades=0 then 0 when BlisterAcuBase.unidades is null then 0 when BlisterAcuBase.unidades>0 then ((BlisterAcu.unidades/BlisterAcuBase.unidades)-1)*100 end,"
            + " isnull(BlisterAcu.VentaNeta,0) as BlisterAcuVta,isnull(BlisterAcuBase.VentaNeta,0) as BlisterBaseVta,[Inc o Dec]=case when BlisterAcuBase.VentaNeta=0 then 0 when BlisterAcuBase.VentaNeta is null then 0 when BlisterAcuBase.VentaNeta>0 then ((BlisterAcu.VentaNeta/BlisterAcuBase.VentaNeta)-1)*100 end,"
            + " isnull(BlisterAcu.UtilBruta,0) as BlisterAcuUtil,isnull(BlisterAcuBase.UtilBruta,0) as BlisterBaseUtil,[Inc o Dec]=case when BlisterAcuBase.UtilBruta=0 then 0 when BlisterAcuBase.UtilBruta is null then 0 when BlisterAcuBase.UtilBruta>0 then ((BlisterAcu.UtilBruta/BlisterAcuBase.UtilBruta)-1)*100 end,"
            + " isnull(BurbujaAcu.unidades,0) as BurbujaAcuUni,isnull(BurbujaAcuBase.unidades,0) as BurbujaBaseUni,[Inc o Dec]=case when BurbujaAcuBase.unidades=0 then 0 when BurbujaAcuBase.unidades is null then 0 when BurbujaAcuBase.unidades>0 then ((BurbujaAcu.unidades/BurbujaAcuBase.unidades)-1)*100 end,"
            + " isnull(BurbujaAcu.VentaNeta,0) as BurbujaAcuVta,isnull(BurbujaAcuBase.VentaNeta,0) as BurbujaBaseVta,[Inc o Dec]=case when BurbujaAcuBase.VentaNeta=0 then 0 when BurbujaAcuBase.VentaNeta is null then 0 when BurbujaAcuBase.VentaNeta>0 then ((BurbujaAcu.VentaNeta/BurbujaAcuBase.VentaNeta)-1)*100 end,"
            + " isnull(BurbujaAcu.UtilBruta,0) as BurbujaAcuUtil,isnull(BurbujaAcuBase.UtilBruta,0) as BurbujaBaseUtil,[Inc o Dec]=case when BurbujaAcuBase.UtilBruta=0 then 0 when BurbujaAcuBase.UtilBruta is null then 0 when BurbujaAcuBase.UtilBruta>0 then ((BurbujaAcu.UtilBruta/BurbujaAcuBase.UtilBruta)-1)*100 end,"
            + " isnull(CajaAcu.unidades,0) as CajaAcuUni,isnull(CajaAcuBase.unidades,0) as CajaBaseUni,[Inc o Dec]=case when CajaAcuBase.unidades=0 then 0 when CajaAcuBase.unidades is null then 0 when CajaAcuBase.unidades>0 then ((CajaAcu.unidades/CajaAcuBase.unidades)-1)*100 end,"
            + " isnull(CajaAcu.VentaNeta,0) as CajaAcuVta,isnull(CajaAcuBase.VentaNeta,0) as CajaBaseVta,[Inc o Dec]=case when CajaAcuBase.VentaNeta=0 then 0 when CajaAcuBase.VentaNeta is null then 0 when CajaAcuBase.VentaNeta>0 then ((CajaAcu.VentaNeta/CajaAcuBase.VentaNeta)-1)*100 end,"
            + " isnull(CajaAcu.UtilBruta,0) as CajaAcuUtil,isnull(CajaAcuBase.UtilBruta,0) as CajaBaseUtil,[Inc o Dec]=case	when CajaAcuBase.UtilBruta=0 then 0	when CajaAcuBase.UtilBruta is null then 0 when CajaAcuBase.UtilBruta>0 then ((CajaAcu.UtilBruta/CajaAcuBase.UtilBruta)-1)*100 end,"
            + " isnull(LuzAcu.unidades,0) as LuzAcuUni,isnull(LuzAcuBase.unidades,0) as LuzBaseUni,[Inc o Dec]=case when LuzAcuBase.unidades=0 then 0 when LuzAcuBase.unidades is null then 0 when LuzAcuBase.unidades>0 then ((LuzAcu.unidades/LuzAcuBase.unidades)-1)*100 end,"
            + " isnull(LuzAcu.VentaNeta,0) as LuzAcuVta,isnull(LuzAcuBase.VentaNeta,0) as LuzBaseVta,[Inc o Dec]=case when LuzAcuBase.VentaNeta=0 then 0 when LuzAcuBase.VentaNeta is null then 0 when LuzAcuBase.VentaNeta>0 then ((LuzAcu.VentaNeta/LuzAcuBase.VentaNeta)-1)*100 end,"
            + " isnull(LuzAcu.UtilBruta,0) as LuzAcuUtil,isnull(LuzAcuBase.UtilBruta,0) as LuzBaseUtil,[Inc o Dec]=case when LuzAcuBase.UtilBruta=0 then 0 when LuzAcuBase.UtilBruta is null then 0 when LuzAcuBase.UtilBruta>0 then ((LuzAcu.UtilBruta/LuzAcuBase.UtilBruta)-1)*100 end,"
            + " isnull(NavideñoAcu.unidades,0) as NaviAcuUni,isnull(NavideñoAcuBase.unidades,0) as NaviBaseUni,[Inc o Dec]=case when NavideñoAcuBase.unidades=0 then 0 when NavideñoAcuBase.unidades is null then 0 when NavideñoAcuBase.unidades>0 then ((NavideñoAcu.unidades/NavideñoAcuBase.unidades)-1)*100 end,"
            + " isnull(NavideñoAcu.VentaNeta,0) as NaviAcuVta,isnull(NavideñoAcuBase.VentaNeta,0) as NaviBaseVta,[Inc o Dec]=case when NavideñoAcuBase.VentaNeta=0 then 0 when NavideñoAcuBase.VentaNeta is null then 0 when NavideñoAcuBase.VentaNeta>0 then ((NavideñoAcu.VentaNeta/NavideñoAcuBase.VentaNeta)-1)*100 end,"
            + " isnull(NavideñoAcu.UtilBruta,0) as NaviAcuUtil,isnull(NavideñoAcuBase.UtilBruta,0) as NaviBaseUtil,[Inc o Dec]=case when NavideñoAcuBase.UtilBruta=0 then 0 when NavideñoAcuBase.UtilBruta is null then 0 when NavideñoAcuBase.UtilBruta>0 then ((NavideñoAcu.UtilBruta/NavideñoAcuBase.UtilBruta)-1)*100 end,"
            + " isnull(PinoAcu.unidades,0) as PinoAcuUni,isnull(PinoAcuBase.unidades,0) as PinoBaseUni,[Inc o Dec]=case when PinoAcuBase.unidades=0 then 0 when PinoAcuBase.unidades is null then 0 when PinoAcuBase.unidades>0 then ((PinoAcu.unidades/PinoAcuBase.unidades)-1)*100 end,"
            + " isnull(PinoAcu.VentaNeta,0) as PinoAcuVta,isnull(PinoAcuBase.VentaNeta,0) as PinoBaseVta,[Inc o Dec]=case when PinoAcuBase.VentaNeta=0 then 0 when PinoAcuBase.VentaNeta is null then 0 when PinoAcuBase.VentaNeta>0 then ((PinoAcu.VentaNeta/PinoAcuBase.VentaNeta)-1)*100 end,"
            + " isnull(PinoAcu.UtilBruta,0) as PinoAcuUtil,isnull(PinoAcuBase.UtilBruta,0) as PinoBaseUtil,[Inc o Dec]=case when PinoAcuBase.UtilBruta=0 then 0 when PinoAcuBase.UtilBruta is null then 0 when PinoAcuBase.UtilBruta>0 then ((PinoAcu.UtilBruta/PinoAcuBase.UtilBruta)-1)*100 End"
            + " from ((((((((((((select cod_estab,nombre from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','67')) as Sucursales left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='28' and fecha between '" + primero.ToString("yyyyMMdd") + "'  and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BlisterAcu on Sucursales.cod_estab=BlisterAcu.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='28' and fecha between '" + primerobase.ToString("yyyyMMdd") + "'  and '" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as BlisterAcuBase on Sucursales.cod_estab=BlisterAcuBase.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod"
            + " where p.familia='29' and fecha between '" + primero.ToString("yyyyMMdd") + "'  and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as BurbujaAcu on Sucursales.cod_estab=BurbujaAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod "
            + " where p.familia='29' and fecha between '" + primerobase.ToString("yyyyMMdd") + "'  and '" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as BurbujaAcuBase on Sucursales.cod_estab=BurbujaAcuBase.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='30' and fecha between '" + primero.ToString("yyyyMMdd") + "'  and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as CajaAcu on Sucursales.cod_estab=CajaAcu.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='30' and fecha between '" + primerobase.ToString("yyyyMMdd") + "'  and '" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as CajaAcuBase on Sucursales.cod_estab=CajaAcuBase.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod"
            + " where p.familia='33' and fecha between '" + primero.ToString("yyyyMMdd") + "'  and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as LuzAcu on Sucursales.cod_estab=LuzAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod"
            + " where p.familia='33' and fecha between '" + primerobase.ToString("yyyyMMdd") + "'  and '" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as LuzAcuBase on Sucursales.cod_estab=LuzAcuBase.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod )"
            + " left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='35' and fecha between '" + primero.ToString("yyyyMMdd") + "'  and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as NavideñoAcu on Sucursales.cod_estab=NavideñoAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta"
            + " from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='35' and fecha between '" + primerobase.ToString("yyyyMMdd") + "'  and '" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as NavideñoAcuBase on Sucursales.cod_estab=NavideñoAcuBase.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia='38' and fecha between '" + primero.ToString("yyyyMMdd") + "'  and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as PinoAcu on Sucursales.cod_estab=PinoAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta"
            + " from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia='38' and fecha between '" + primerobase.ToString("yyyyMMdd") + "'  and '" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as PinoAcuBase on Sucursales.cod_estab=PinoAcuBase.cod_estab order by CAST(Sucursales.cod_estab as int) asc";
            dt = new System.Data.DataTable();
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            da.Fill(dt);
            System.Windows.Forms.Application.DoEvents();
            DG1.DataSource = dt;
            DG1.SelectAll();
            objeto = DG1.GetClipboardContent();
            if (objeto != null)
            {

                Clipboard.SetDataObject(objeto);
                hoja = (Worksheet)libro.Sheets.get_Item(3);
                hoja.Activate();
                hoja.Name = "COMPARATIVO ACUMULADO";
                hoja.Cells[1, 2] = "COMPARATIVO ACUMULADO";
                //ENCABEZADO VENTA DIARIA
                rango = (Range)hoja.get_Range("B1", "BE2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango = (Range)hoja.get_Range("B3", "BE3");
                rango.Select();
                rango.Merge();
                rango.Cells[1.1, Type.Missing] = "Del " + primero.ToString("d MMM yyyy") + " al " + dia.ToString("d MMM yyyy");
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango = (Range)hoja.get_Range("B4", "B6");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "CODIGO";
                rango = (Range)hoja.get_Range("C4", "C6");
                rango.Select();
                rango.Merge();
                //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "SUCURSAL";
                for (int i = 4; i <= 49; i += 9)
                {
                    //rango = (Range)hoja.get_Range("4,4","12,4");
                    rango = (Range)hoja.get_Range(sCol(i) + "4", sCol(i + 8) + "4");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    switch (i)
                    {
                        case 4:
                            rango.Cells[1, 1] = "COMPARATIVO ACUMULADO BLISTER";
                            break;
                        case 13:
                            rango.Cells[1, 1] = "COMPARATIVO ACUMULADO BURBUJA";
                            break;
                        case 22:
                            rango.Cells[1, 1] = "COMPARATIVO ACUMULADO CAJA";
                            break;
                        case 31:
                            rango.Cells[1, 1] = "COMPARATIVO ACUMULADO LUZ";
                            break;
                        case 40:
                            rango.Cells[1, 1] = "COMPARATIVO ACUMULADO NAVIDEÑO";
                            break;
                        case 49:
                            rango.Cells[1, 1] = "COMPARATIVO ACUMULADO PINO";
                            break;
                    }
                    for (int j = i; j <= i + 8; j += 3)
                    {
                        rango = (Range)hoja.get_Range(sCol(j) + "5", sCol(j + 1) + "5");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        hoja.Cells[5, j + 2] = "%";
                        hoja.Cells[5, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        hoja.Cells[5, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        if (j == i)
                        {
                            rango.Cells[1, 1] = "UNIDADES";
                        }
                        else if (j == i + 3)
                        {
                            rango.Cells[1, 1] = "VENTA NETA";
                        }
                        else if (j == i + 6)
                        {
                            rango.Cells[1, 1] = "UTILIDAD BRUTA";
                        }
                        for (int k = j; k <= j + 2; k++)
                        {
                            if (k == j)
                            {
                                hoja.Cells[6, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                hoja.Cells[6, k].NumberFormat = "@";
                                hoja.Cells[6, k] = dia.Year.ToString();
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=SUM(" + sCol(k) + "7:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 6) + ")";
                            }
                            else if (k == j + 1)
                            {
                                hoja.Cells[6, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                hoja.Cells[6, k].NumberFormat = "@";
                                hoja.Cells[6, k] = diabase.Year.ToString();
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=SUM(" + sCol(k) + "7:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 6) + ")";
                            }
                            else if (k == j + 2)
                            {
                                hoja.Cells[6, k] = "Inc o Dec";
                                hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00";
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 7) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 7) + ")-1)*100";
                            }
                            hoja.Cells[6, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[6, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            hoja.Cells[6, k].Font.FontStyle = "Bold";

                        }
                    }
                }
                hoja.Range["B7", "C" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                hoja.Range["D7", "BE" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                hoja.Range["B7", "BE" + Convert.ToString(DG1.Rows.Count + 7)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                hoja.Cells[DG1.Rows.Count + 7, 3] = "T O T A L";
                rango = (Range)hoja.Cells[7, 1];
                rango.Select();
                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                rango = (Range)hoja.get_Range("A1", "BE" + Convert.ToString(DG1.Rows.Count + 10));
                rango.EntireColumn.AutoFit();
            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //
            //                                          COMPARATIVO TOTAL DIARIO Y ACUMULADO
            //
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            Clipboard.Clear();
            objeto = null;
            DG1.DataSource = null;
            DG1.Rows.Clear();
            DG1.Columns.Clear();
            cmd.CommandText = "select Sucursales.cod_estab,Sucursales.nombre,"
            + " isnull(TotalDiario.unidades,0) as TotalDiarioUni,isnull(TotalBase.unidades,0) as TotalDiarioBaseUni,[Inc o Dec]=case when TotalBase.unidades=0 then 0 when TotalBase.unidades is null then 0 when TotalBase.unidades>0 then ((TotalDiario.unidades/TotalBase.unidades)-1)*100 end,"
            + " isnull(TotalDiario.VentaNeta,0) as TotalDiarioVta,isnull(TotalBase.VentaNeta,0) as TotalDiarioBaseVta,[Inc o Dec]=case when TotalBase.VentaNeta=0 then 0 when TotalBase.VentaNeta is null then 0 when TotalBase.VentaNeta>0 then ((TotalDiario.VentaNeta/TotalBase.VentaNeta)-1)*100 end,"
            + " isnull(TotalDiario.UtilBruta,0) as TotalDiarioUtil,isnull(TotalBase.UtilBruta,0) as TotalDiarioBaseUtil,[Inc o Dec]=	case when TotalBase.UtilBruta=0 then 0 when TotalBase.UtilBruta is null then 0 when TotalBase.UtilBruta>0 then ((TotalDiario.UtilBruta/TotalBase.UtilBruta)-1)*100 end,"
            + " isnull(TotalAcu.unidades,0) as TotalAcuUni,isnull(TotalAcuBase.unidades,0) as TotalAcuBaseUni,[Inc o Dec]=case when TotalAcuBase.unidades=0 then 0 when TotalAcuBase.unidades is null then 0 when TotalAcuBase.unidades>0 then ((TotalAcu.unidades/TotalAcuBase.unidades)-1)*100 end,"
            + " isnull(TotalAcu.VentaNeta,0) as TotalAcuVta,isnull(TotalAcuBase.VentaNeta,0) as TotalAcuBaseVta,[Inc o Dec]=case when TotalAcuBase.VentaNeta=0 then 0 when TotalAcuBase.VentaNeta is null then 0 when TotalAcuBase.VentaNeta>0 then ((TotalAcu.VentaNeta/TotalAcuBase.VentaNeta)-1)*100 end,"
            + " isnull(TotalAcu.UtilBruta,0) as TotalAcuUtil,isnull(TotalAcuBase.UtilBruta ,0) TotalAcuBaseUtil,[Inc o Dec]=case when TotalAcuBase.UtilBruta=0 then 0 when TotalAcuBase.UtilBruta is null then 0 when TotalAcuBase.UtilBruta>0 then ((TotalAcu.UtilBruta/TotalAcuBase.UtilBruta)-1)*100 end"
            + " from (((((select cod_estab,nombre from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','67')) as Sucursales left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta"
            + " from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod where p.familia in ('28','29','30','33','38','35') and fecha='" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as TotalDiario on Sucursales.cod_estab=TotalDiario.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod where p.familia in ('28','29','30','33','38','35') and fecha='" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as TotalBase on Sucursales.cod_estab=TotalBase.cod_estab)"
            + " left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva)-SUM(va.cantidad*isnull(cedis.ultimo_costo,1)) as UtilBruta from (vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod ) left join (select cod_prod,ultimo_costo from bmsepm_cedis..prodestab where cod_estab='65') as cedis on va.cod_prod=cedis.cod_prod"
            + " where p.familia in ('28','29','30','33','38','35') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' group by va.cod_estab) as TotalAcu on Sucursales.cod_estab=TotalAcu.cod_estab) left join (select cod_estab,sum(cantidad) as unidades,SUM(va.VentaNeta) as VentaNeta, SUM(va.VentaNetaSinIva-(va.costo/1.0633)) as UtilBruta from vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod"
            + " where p.familia in ('28','29','30','33','38','35') and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd") + "'  group by va.cod_estab) as TotalAcuBase on Sucursales.cod_estab=TotalAcuBase .cod_estab) order by CAST(Sucursales.cod_estab as int) asc";
            dt = new System.Data.DataTable();
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            da.Fill(dt);
            System.Windows.Forms.Application.DoEvents();
            DG1.DataSource = dt;
            DG1.SelectAll();
            objeto = DG1.GetClipboardContent();
            if (objeto != null)
            {
                Clipboard.SetDataObject(objeto);
                hoja = (Worksheet)libro.Sheets.get_Item(4);
                hoja.Activate();
                hoja.Name = "COMPARATIVO TOTAL";
                hoja.Cells[1, 2] = "COMPARATIVO TOTAL";
                //ENCABEZADO VENTA DIARIA
                rango = (Range)hoja.get_Range("B1", "U2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango = (Range)hoja.get_Range("B3", "U3");
                rango.Select();
                rango.Merge();
                rango.Cells[1.1, Type.Missing] = "Del " + primero.ToString("d MMM yyyy") + " al " + dia.ToString("d MMM yyyy");
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango = (Range)hoja.get_Range("B4", "B6");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "CODIGO";
                rango = (Range)hoja.get_Range("C4", "C6");
                rango.Select();
                rango.Merge();
                //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Cells.Font.FontStyle = "Bold";
                rango.Cells[1, 1] = "SUCURSAL";
                for (int i = 4; i <= 21; i += 9)
                {
                    //rango = (Range)hoja.get_Range("4,4","12,4");
                    rango = (Range)hoja.get_Range(sCol(i) + "4", sCol(i + 8) + "4");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    switch (i)
                    {
                        case 4:
                            rango.Cells[1, 1] = "COMPARATIVO TOTAL DIARIO";
                            break;
                        case 13:
                            rango.Cells[1, 1] = "COMPARATIVO TOTAL ACUMULADO";
                            break;

                    }
                    for (int j = i; j <= i + 8; j += 3)
                    {
                        rango = (Range)hoja.get_Range(sCol(j) + "5", sCol(j + 1) + "5");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        hoja.Cells[5, j + 2] = "%";
                        hoja.Cells[5, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        hoja.Cells[5, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        if (j == i)
                        {
                            rango.Cells[1, 1] = "UNIDADES";
                        }
                        else if (j == i + 3)
                        {
                            rango.Cells[1, 1] = "VENTA NETA";
                        }
                        else if (j == i + 6)
                        {
                            rango.Cells[1, 1] = "UTILIDAD BRUTA";
                        }
                        for (int k = j; k <= j + 2; k++)
                        {
                            if (k == j)
                            {
                                hoja.Cells[6, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                hoja.Cells[6, k].NumberFormat = "@";
                                hoja.Cells[6, k] = dia.Year.ToString();
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=SUM(" + sCol(k) + "7:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 6) + ")";
                            }
                            else if (k == j + 1)
                            {
                                hoja.Cells[6, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                hoja.Cells[6, k].NumberFormat = "@";
                                hoja.Cells[6, k] = diabase.Year.ToString();
                                if (j == i)
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0"; }
                                else
                                { hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00"; }
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=SUM(" + sCol(k) + "7:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 6) + ")";
                            }
                            else if (k == j + 2)
                            {
                                hoja.Cells[6, k] = "Inc o Dec";
                                hoja.Range[sCol(k) + "7", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00";
                                hoja.Cells[DG1.Rows.Count + 7, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 7) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 7) + ")-1)*100";
                            }
                            hoja.Cells[6, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[6, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            hoja.Cells[6, k].Font.FontStyle = "Bold";

                        }
                    }
                }
                hoja.Range["B7", "C" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                hoja.Range["D7", "U" + Convert.ToString(DG1.Rows.Count + 6)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                hoja.Range["B7", "U" + Convert.ToString(DG1.Rows.Count + 7)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                hoja.Cells[DG1.Rows.Count + 7, 3] = "T O T A L";
                rango = (Range)hoja.Cells[7, 1];
                rango.Select();
                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                rango = (Range)hoja.get_Range("A1", "U" + Convert.ToString(DG1.Rows.Count + 10));
                rango.EntireColumn.AutoFit();
            }
            hoja = (Worksheet)libro.Sheets.get_Item(1);
            hoja.Activate();
            if (cn.State.ToString() == "Open") { cn.Close(); }
            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas de Temporada al " + dia.ToString("dd MMM yyyy") + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas de Temporada al " + dia.ToString("dd MMM yyyy") + ".xlsx");
            }
            libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas de Temporada al " + dia.ToString("dd MMM yyyy") + ".xlsx");
            libro.Close();
            excel.Quit();
            return System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas de Temporada al " + dia.ToString("dd MMM yyyy") + ".xlsx";
        }

        private void EnviaMail(string sReport, string Destinatarios)
        {
            if (sReport == "" || Destinatarios == "") { return; }
            string[] temp = sReport.Split('\\');
            string asunto = temp.Last();
            temp = asunto.Split('.');
            asunto = temp.First();
            MailMessage mail = new MailMessage();
            mail.Bcc.Add(Destinatarios);
            mail.To.Add("maferperezle01@gmail.com");
            mail.To.Add("programador@mercadodeimportaciones.com");

            mail.From = new MailAddress("sistemas@mercadodeimportaciones.com");
            //mail.From = new MailAddress("programador@mercadodeimportaciones.com");
            mail.Body = "<h2 style=\"font-style:italic;\"><strong>REPORT SERVICE</strong></h2>";
            mail.IsBodyHtml = true;
            mail.Subject = asunto;
            mail.Priority = MailPriority.Normal;
            Attachment adjunto = new Attachment(sReport);
            mail.Attachments.Add(adjunto);
            SmtpClient server = new SmtpClient();
            //server.Host = "mail.mercadodeimportaciones.com";
            server.Host = "187.216.118.171";
            server.Port = 25;
            server.EnableSsl = false;
            server.DeliveryMethod = SmtpDeliveryMethod.Network;
            //server.Port = 465;  // Usar el puerto seguro para que no haya fallas
            //server.EnableSsl = true;
            //server.UseDefaultCredentials = false;
            server.Timeout = 0;
            server.Credentials = new System.Net.NetworkCredential("sistemas@mercadodeimportaciones.com", "Mercado1");
            //server.Credentials = new System.Net.NetworkCredential("programador@mercadodeimportaciones.com", "Mercado1");
            try
            {
                //server.SendAsync(mail, (object)mail);
                server.Send(mail);
            }
            catch (Exception e)
            {
                lblEstado.Text = "Hubo un problema con el envío del Reporte " + e.Message.ToString();
            }
            //server.Send(mail);
            mail.Dispose();
        }

        private void EnviaMailGmail(string sReport, string Destinatarios)
        {
            if (sReport == "" || Destinatarios == "") { return; }
            string[] temp = sReport.Split('\\');
            string asunto = temp.Last();
            temp = asunto.Split('.');
            asunto = temp.First();

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress("sistemas.gmimportaciones@gmail.com");
            mail.Bcc.Add(Destinatarios);
            mail.To.Add("maferperezle01@gmail.com");
            mail.To.Add("programador@mercadodeimportaciones.com");
            mail.Subject = asunto;
            mail.IsBodyHtml = true;
            mail.Body = "<h2 style=\"font-style:italic;\"><strong>REPORT SERVICE</strong></h2>";
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(sReport);
            mail.Attachments.Add(attachment);
            SmtpServer.Port = 587;
            //SmtpServer.Credentials = new System.Net.NetworkCredential("sistemas.gmimportaciones@gmail.com", "$mercadoYladies$");
            // 08/Jun/2022 - Se activó la verificación en 2 pasos en GMAIL y se tiene que usar la siguiente contraseña para aplicaciones: ynpjfjgjwblgupam 
            SmtpServer.Credentials = new System.Net.NetworkCredential("sistemas.gmimportaciones@gmail.com", "ynpjfjgjwblgupam");
            SmtpServer.EnableSsl = true;
            try
            {
                SmtpServer.SendAsync(mail, (object)mail);
            }
            catch (Exception e)
            {
                lblEstado.Text = "Hubo un problema con el envío del Reporte " + e.Message.ToString();
            }
        }

        private string ComparativoLineas()
        {
            SqlConnection cn = conexion.conectar("BDIntegrador");
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            System.Data.DataTable dt = new System.Data.DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            DateTime dia = DateTime.Now.AddDays(-1);
            DateTime primero = dia.AddDays((dia.Day - 1) * -1);
            DateTime diabase = dia.AddYears(-1);
            DateTime primerobase = primero.AddYears(-1);
            cmd.CommandText = "SELECT cast(case when anterior.cod_estab is null then actual.cod_estab else anterior.cod_estab end as int) as cod_estab,case when anterior.nombre is null then actual.nombre else anterior.nombre end as Establecimiento,"
            + " case when anterior.CodLinea is null then actual.CodLinea else anterior.CodLinea end as CodLinea,case when anterior.linea is null then actual.linea else anterior.linea end as Linea,case when anterior.CodFamilia is null then anterior.CodFamilia else anterior.CodFamilia end as CodFamilia,"
            + " case when anterior.Familia is null then anterior.Familia else anterior.Familia end as Familia,case when anterior.CodClasificacion is null then actual.CodClasificacion else anterior.CodClasificacion end as CodClasificacion,case when anterior.Clasificacion is null then actual.Clasificacion else anterior.Clasificacion end  as Clasificacion,"
            + " isnull(actual.unidades,0) as unidades,isnull(actual.importe,0) as importe,isnull(actual.descto,0) as descto,isnull(actual.VentaNeta,0)as VentaNeta,isnull(anterior.unidades,0) as Unidades,"
            + " isnull(anterior.importe,0) as Importe,isnull(anterior.descto,0) as Descto,isnull(anterior.VentaNeta,0) as VentaNeta,ISNULL(actual.ventaneta,0)-ISNULL(anterior.VentaNeta,0) as Diferencia FROM "
            + " (select va.cod_estab,establecimientos.nombre,isnull(lineas_productos.linea_producto,'S/L') as CodLinea,isnull(lineas_productos.nombre,'SIN LINEA') as linea,isnull(familias.familia,'S/F') as CodFamilia,isnull(familias.nombre,'SIN FAMILIA') as Familia,isnull(clasificaciones_productos.clasificacion_productos,'S/C') as CodClasificacion,"
            + " isnull(clasificaciones_productos.nombre,'SIN CLASIFICACION') as Clasificacion,SUM(va.cantidad) as unidades,SUM(va.Importe) as importe,SUM(va.descuento) as descto,SUM(va.VentaNeta) as VentaNeta  from "
            + " ((((vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod) left join lineas_productos on p.linea_producto=lineas_productos.linea_producto) left join familias on familias.familia=p.familia)"
            + " left join clasificaciones_productos on clasificaciones_productos.clasificacion_productos=p.clasificacion_productos) inner join establecimientos on establecimientos.cod_estab=va.cod_estab where va.fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd") + "' and establecimientos.status='V' group by"
            + " va.cod_estab,establecimientos.nombre,lineas_productos.linea_producto,lineas_productos.nombre,familias.familia,familias.nombre,clasificaciones_productos.clasificacion_productos,clasificaciones_productos.nombre) AS anterior"
            + " full join (select va.cod_estab,establecimientos.nombre,isnull(lineas_productos.linea_producto,'S/L') as CodLinea,isnull(lineas_productos.nombre,'SIN LINEA') as linea,isnull(familias.familia,'S/F') as CodFamilia,isnull(familias.nombre,'SIN FAMILIA') as Familia,isnull(clasificaciones_productos.clasificacion_productos,'S/C') as CodClasificacion,"
            + " isnull(clasificaciones_productos.nombre,'SIN CLASIFICACION') as Clasificacion,SUM(va.cantidad) as unidades,SUM(va.Importe) as importe,SUM(va.descuento) as descto,SUM(va.VentaNeta) as VentaNeta  from ((((vta_artic_consolid_real as va inner join productos as p on va.cod_prod=p.cod_prod) left join lineas_productos on p.linea_producto=lineas_productos.linea_producto)"
            + " left join familias on familias.familia=p.familia) left join clasificaciones_productos on clasificaciones_productos.clasificacion_productos=p.clasificacion_productos) inner join establecimientos on establecimientos.cod_estab=va.cod_estab where va.fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + "' and establecimientos.status='V' group by "
            + " va.cod_estab,establecimientos.nombre,lineas_productos.linea_producto,lineas_productos.nombre,familias.familia,familias.nombre,clasificaciones_productos.clasificacion_productos,clasificaciones_productos.nombre) as actual on anterior.CodLinea=actual.CodLinea and anterior.CodFamilia=actual.CodFamilia and anterior.Clasificacion=actual.Clasificacion and "
            + " anterior.cod_estab=actual.cod_estab";
            da.Fill(dt);
            DG1.DataSource = null;
            DG1.Rows.Clear();
            DG1.Columns.Clear();
            DG1.DataSource = dt;
            DG1.SelectAll();
            object objeto = DG1.GetClipboardContent();
            Microsoft.Office.Interop.Excel.Application excel;
            excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook libro;
            libro = excel.Workbooks.Add();
            libro.Worksheets.Add();
            Worksheet hoja = new Worksheet();
            hoja = (Worksheet)libro.Worksheets.get_Item(1);
            hoja.Name = "COMPARATIVO POR CLASIFICACION";
            Microsoft.Office.Interop.Excel.Range rango;
            if (objeto != null)
            {
                Clipboard.SetDataObject(objeto);
                hoja.Cells[2, 10] = "VENTAS DEL " + primero.ToString("dd MMM yyyy") + " AL " + dia.ToString("dd MMM yyyy");
                rango = (Range)hoja.get_Range("J2", "M2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                hoja.Cells[2, 14] = "VENTAS DEL " + primerobase.ToString("dd MMM yyyy") + " AL " + diabase.ToString("dd MMM yyyy");
                rango = (Range)hoja.get_Range("N2", "Q2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                for (int i = 1; i <= DG1.Columns.Count; i++)
                {
                    hoja.Cells[3, i + 1] = DG1.Columns[i - 1].Name.ToString();
                    if (i == 9 || i == 13)
                    {
                        hoja.Range[sCol(i + 1) + "3", sCol(i + 1) + Convert.ToString(DG1.Rows.Count + 3)].EntireColumn.NumberFormat = "###,###,##0";
                    }
                    else if (i > 9)
                    {
                        hoja.Range[sCol(i + 1) + "3", sCol(i + 1) + Convert.ToString(DG1.Rows.Count + 3)].EntireColumn.NumberFormat = "###,###,##0.00";
                    }
                }
                rango = (Range)hoja.get_Range("B2", "S3");
                rango.Select();
                rango.Cells.Font.FontStyle = "Bold";

                rango = (Range)hoja.Cells[4, 1];
                rango.Select();
                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                rango = (Range)hoja.get_Range("A1", "S" + Convert.ToString(DG1.Rows.Count + 10));
                rango.EntireColumn.AutoFit();
                rango = (Range)hoja.get_Range("B4", "R" + Convert.ToString(DG1.Rows.Count + 3));
                rango.Sort(rango.Columns[1], XlSortOrder.xlAscending, rango.Columns[3], Type.Missing, XlSortOrder.xlAscending, rango.Columns[5], XlSortOrder.xlAscending);
                rango = (Range)hoja.get_Range("B3", "R" + Convert.ToString(DG1.Rows.Count + 3));
                rango.Select();
                int[] cols = new int[] { 9, 10, 11, 12, 13, 14, 15, 16, 17 };
                rango.Subtotal(2, XlConsolidationFunction.xlSum, cols, true, Type.Missing, XlSummaryRow.xlSummaryBelow);
                rango.Subtotal(4, XlConsolidationFunction.xlSum, cols, false, Type.Missing);
                rango.Subtotal(6, XlConsolidationFunction.xlSum, cols, false, Type.Missing);

            }
            if (cn.State.ToString() == "Open") { cn.Close(); }
            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas por Estab-Clasificacion al " + dia.ToString("dd MMM yyyy") + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas por Estab-Clasificacion al " + dia.ToString("dd MMM yyyy") + ".xlsx");
            }
            libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas por Estab-Clasificacion al " + dia.ToString("dd MMM yyyy") + ".xlsx");
            libro.Close();
            excel.Quit();
            return System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas por Estab-Clasificacion al " + dia.ToString("dd MMM yyyy") + ".xlsx";
        }
        private string ReporteVentasDiarias(DateTime FechaHora)
        {
            try
            {
                SqlConnection cn = conexion.conectar("BMSNayar");
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandTimeout = 240;
                System.Data.DataTable dt = new System.Data.DataTable();
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                DateTime dia = FechaHora;
                //DateTime dia = Convert.ToDateTime("2017-05-31 23:59");
                DateTime primero = dia.AddDays((dia.Day - 1) * -1);
                //DateTime primero = Convert.ToDateTime("2017-01-01");
                DateTime diabase = dia.AddYears(-1);
                DateTime primerobase = primero.AddYears(-1);
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //  PREPARA TABLA TEMPORAL DE TRAAJO
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "delete from entysalVentas";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into entysalVentas( folio,transaccion,fecha,cod_prod,unidad,cantidad,precio_lista,tipo_precio_venta,descuento_porcentual,importe_descuento,importe,"
                + "iva,ieps,costo,peso,volumen,cod_estab,cod_cte,cod_prv,vendedor,status,id_origen,abreviatura_unidad,fecha_proceso) "
                + "select entysal.folio,entysal.transaccion,entysal.fecha,entysal.cod_prod,entysal.unidad,entysal.cantidad,entysal.precio_lista,entysal.tipo_precio_venta,entysal.descuento_porcentual,entysal.importe_descuento,entysal.importe,"
                + "entysal.iva,entysal.ieps,entysal.costo,entysal.peso,entysal.volumen,entysal.cod_estab,entysal.cod_cte,entysal.cod_prv,entysal.vendedor,entysal.status,entysal.id_origen,entysal.abreviatura_unidad,entysal.fecha_proceso "
                + "from BMSMayoristas..entysal with(nolock) inner join productos with(nolock) on entysal.cod_prod=productos.cod_prod where productos.tipo_producto<>'7' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' and transaccion in ('36','37','38','308','34','68')";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into entysalVentas( folio,transaccion,fecha,cod_prod,unidad,cantidad,precio_lista,tipo_precio_venta,descuento_porcentual,importe_descuento,importe,"
                + "iva,ieps,costo,peso,volumen,cod_estab,cod_cte,cod_prv,vendedor,status,id_origen,abreviatura_unidad,fecha_proceso)"
                + "select entysal.folio,entysal.transaccion,entysal.fecha,entysal.cod_prod,entysal.unidad,entysal.cantidad,entysal.precio_lista,entysal.tipo_precio_venta,entysal.descuento_porcentual,entysal.importe_descuento,entysal.importe,"
                + "entysal.iva,entysal.ieps,entysal.costo,entysal.peso,entysal.volumen,entysal.cod_estab,entysal.cod_cte,entysal.cod_prv,entysal.vendedor,entysal.status,entysal.id_origen,entysal.abreviatura_unidad,entysal.fecha_proceso "
                + "from BMSMayoristas..entysal with(nolock) inner join productos with(nolock) on entysal.cod_prod=productos.cod_prod where productos.tipo_producto<>'7' and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' and transaccion in ('36','37','38','308','34','68')";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into entysalVentas( folio,transaccion,fecha,cod_prod,unidad,cantidad,precio_lista,tipo_precio_venta,descuento_porcentual,importe_descuento,importe,"
                + "iva,ieps,costo,peso,volumen,cod_estab,cod_cte,cod_prv,vendedor,status,id_origen,abreviatura_unidad,fecha_proceso) "
                + "select entysal.folio,entysal.transaccion,entysal.fecha,entysal.cod_prod,entysal.unidad,entysal.cantidad,entysal.precio_lista,entysal.tipo_precio_venta,entysal.descuento_porcentual,entysal.importe_descuento,entysal.importe,"
                + "entysal.iva,entysal.ieps,entysal.costo,entysal.peso,entysal.volumen,entysal.cod_estab,entysal.cod_cte,entysal.cod_prv,entysal.vendedor,entysal.status,entysal.id_origen,entysal.abreviatura_unidad,entysal.fecha_proceso "
                + "from BMSCajeme..entysal with(nolock) inner join productos with(nolock) on entysal.cod_prod=productos.cod_prod where productos.tipo_producto<>'7' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' and transaccion in ('36','37','38','308','34','68')";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into entysalVentas( folio,transaccion,fecha,cod_prod,unidad,cantidad,precio_lista,tipo_precio_venta,descuento_porcentual,importe_descuento,importe,"
                + "iva,ieps,costo,peso,volumen,cod_estab,cod_cte,cod_prv,vendedor,status,id_origen,abreviatura_unidad,fecha_proceso)"
                + "select entysal.folio,entysal.transaccion,entysal.fecha,entysal.cod_prod,entysal.unidad,entysal.cantidad,entysal.precio_lista,entysal.tipo_precio_venta,entysal.descuento_porcentual,entysal.importe_descuento,entysal.importe,"
                + "entysal.iva,entysal.ieps,entysal.costo,entysal.peso,entysal.volumen,entysal.cod_estab,entysal.cod_cte,entysal.cod_prv,entysal.vendedor,entysal.status,entysal.id_origen,entysal.abreviatura_unidad,entysal.fecha_proceso "
                + "from BMSCajeme..entysal with(nolock) inner join productos with(nolock) on entysal.cod_prod=productos.cod_prod where productos.tipo_producto<>'7' and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' and transaccion in ('36','37','38','308','34','68')";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into entysalVentas( folio,transaccion,fecha,cod_prod,unidad,cantidad,precio_lista,tipo_precio_venta,descuento_porcentual,importe_descuento,importe,"
                + "iva,ieps,costo,peso,volumen,cod_estab,cod_cte,cod_prv,vendedor,status,id_origen,abreviatura_unidad,fecha_proceso) "
                + "select entysal.folio,entysal.transaccion,entysal.fecha,entysal.cod_prod,entysal.unidad,entysal.cantidad,entysal.precio_lista,entysal.tipo_precio_venta,entysal.descuento_porcentual,entysal.importe_descuento,entysal.importe,"
                + "entysal.iva,entysal.ieps,entysal.costo,entysal.peso,entysal.volumen,entysal.cod_estab,entysal.cod_cte,entysal.cod_prv,entysal.vendedor,entysal.status,entysal.id_origen,entysal.abreviatura_unidad,entysal.fecha_proceso "
                + "from BMSNayar..entysal with(nolock) inner join productos with(nolock) on entysal.cod_prod=productos.cod_prod where productos.tipo_producto<>'7' and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' and transaccion in ('36','37','38','308','34','68')";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into entysalVentas( folio,transaccion,fecha,cod_prod,unidad,cantidad,precio_lista,tipo_precio_venta,descuento_porcentual,importe_descuento,importe,"
                + "iva,ieps,costo,peso,volumen,cod_estab,cod_cte,cod_prv,vendedor,status,id_origen,abreviatura_unidad,fecha_proceso)"
                + "select entysal.folio,entysal.transaccion,entysal.fecha,entysal.cod_prod,entysal.unidad,entysal.cantidad,entysal.precio_lista,entysal.tipo_precio_venta,entysal.descuento_porcentual,entysal.importe_descuento,entysal.importe,"
                + "entysal.iva,entysal.ieps,entysal.costo,entysal.peso,entysal.volumen,entysal.cod_estab,entysal.cod_cte,entysal.cod_prv,entysal.vendedor,entysal.status,entysal.id_origen,entysal.abreviatura_unidad,entysal.fecha_proceso "
                + "from BMSNayar..entysal with(nolock) inner join productos with(nolock) on entysal.cod_prod=productos.cod_prod where productos.tipo_producto<>'7' and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' and transaccion in ('36','37','38','308','34','68')";
                cmd.ExecuteNonQuery();
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //  INSERTA NOTAS DE CREDITO DEL PERIODO PARA CUADRAR LA CONTRIBUCION
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "insert into entysalventas(fecha,folio,transaccion,cod_estab,cod_prod,importe,iva) select c.fecha,e.folio,e.transaccion,e.cod_estab,e.cod_prod,(e.importe * f.tipo_cambio * -1) * ac.importe / f.total AS ImporteSiva,"
                + " ( (e.total * f.tipo_cambio * -1.0) * ac.importe / f.total)-((e.importe * f.tipo_cambio * -1.0) * ac.importe / f.total) as iva "
                + "from BMSNayar..entysal as e inner join BMSNayar..facremtick as f on e.folio=f.folio and e.transaccion=f.transaccion "
                + "inner join BMSNayar..abonos_clientes as ac on ac.folio_aplicado=f.folio and ac.transaccion_folio_aplicado=f.transaccion "
                + "inner join BMSNayar..creditos_clientes as c on c.folio=ac.folio and c.transaccion=ac.transaccion where c.fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyMMdd HH:mm") + "'";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into entysalventas(fecha,folio,transaccion,cod_estab,cod_prod,importe,iva) select c.fecha,e.folio,e.transaccion,e.cod_estab,e.cod_prod,(e.importe * f.tipo_cambio * -1) * ac.importe / f.total AS ImporteSiva,"
                + "( (e.total * f.tipo_cambio * -1.0) * ac.importe / f.total)-((e.importe * f.tipo_cambio * -1.0) * ac.importe / f.total) as iva "
                + "from BMSNayar..entysal as e inner join BMSNayar..facremtick as f on e.folio=f.folio and e.transaccion=f.transaccion "
                + "inner join BMSNayar..abonos_clientes as ac on ac.folio_aplicado=f.folio and ac.transaccion_folio_aplicado=f.transaccion "
                + "inner join BMSNayar..creditos_clientes as c on c.folio=ac.folio and c.transaccion=ac.transaccion where c.fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyMMdd HH:mm") + "'";
                cmd.ExecuteNonQuery();
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //
                //                                          VENTA DIARIA
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                #region ConsultaVentaDiaria
                cmd.CommandText = "delete from _TempVentas";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "insert into _TempVentas select Sucursales.cod_estab,Sucursales.nombre,isnull(HogarDiario.unidades,0) as HogarDiarioUni,isnull(HogarAcu.unidades,0) as HogarAcuUni,cast(0.0 as float) as '%1',"
                + "isnull(HogarDiario.Total,0) as HogarDiarioVta,isnull(HogarAcu.Total,0) as HogarAcuVta,cast(0.0 as float) as '%2',isnull(HogarDiario.UtilBruta,0) as HogarDiarioUtil,isnull(HogarAcu.UtilBruta,0) as HogarAcuUtil,"
                + "cast(0.0 as float) as '%3',isnull(BisuteriaDiario.unidades,0) as BisuteriaDiarioUni,isnull(BisuteriaAcu.unidades,0) as BisuteriaAcuUni,cast(0.0 as float) as '%4',isnull(BisuteriaDiario.Total,0) as BisuteriaDiarioVta,"
                + "isnull(BisuteriaAcu.Total,0) as BisuteriaAcuVta,cast(0.0 as float) as '%5',isnull(BisuteriaDiario.UtilBruta,0) as BisuteriaDiarioUtil,isnull(BisuteriaAcu.UtilBruta,0) as BisuteriaAcuUtil,cast(0.0 as float) as '%6',"
                + "isnull(BellezaDiario.unidades,0) as BellezaDiarioUni,isnull(BellezaAcu.unidades,0) as BellezaAcuUni,cast(0.0 as float) as '%7',isnull(BellezaDiario.Total,0) as BellezaDiarioVta,isnull(BellezaAcu.Total,0) as BellezaAcuVta,"
                + "cast(0.0 as float) as '%8',isnull(BellezaDiario.UtilBruta,0) as BellezaDiarioUtil,isnull(BellezaAcu.UtilBruta,0) as BellezaAcuUtil,cast(0.0 as float) as '%9',isnull(CalzadoDiario.unidades,0) as CalzadoDiarioUni,"
                + "isnull(CalzadoAcu.unidades,0) as CalzadoAcuUni,cast(0.0 as float) as '%10',isnull(CalzadoDiario.Total,0) as CalzadoDiarioVta,isnull(CalzadoAcu.Total,0) as CalzadoAcuVta,cast(0.0 as float) as '%11',"
                + "isnull(CalzadoDiario.UtilBruta,0) as CalzadoDiarioUtil,isnull(CalzadoAcu.UtilBruta,0) as CalzadoAcuUtil,cast(0.0 as float) as '%12',isnull(RopaDiario.unidades,0) as RopaDiarioUni,isnull(RopaAcu.unidades,0) as RopaAcuUni,"
                + "cast(0.0 as float) as '%13',isnull(RopaDiario.Total,0) as RopaDiarioVta,isnull(RopaAcu.Total,0) as RopaAcuVta,cast(0.0 as float) as '%14',isnull(RopaDiario.UtilBruta,0) as RopaDiarioUtil,isnull(RopaAcu.UtilBruta,0) as RopaAcuUtil,"
                + "cast(0.0 as float) as '%15',isnull(ServiciosDiario.unidades,0) as ServiciosDiarioUni,isnull(ServiciosAcu.unidades,0) as ServiciosAcuUni,cast(0.0 as float) as '%16',isnull(ServiciosDiario.Total,0) as ServiciosDiarioVta,"
                + "isnull(ServiciosAcu.Total,0) as ServiciosAcuVta,cast(0.0 as float) as '%17',isnull(ServiciosDiario.UtilBruta,0) as ServiciosDiarioUtil,isnull(ServiciosAcu.UtilBruta,0) as ServiciosAcuUtil,cast(0.0 as float) as '%18',"
                + "isnull(BotanasDiario.unidades,0) as BotanasDiarioUni,isnull(BotanasAcu.unidades,0) as BotanasAcuUni,cast(0.0 as float) as '%19',isnull(BotanasDiario.Total,0) as BotanasDiarioVta,isnull(BotanasAcu.Total,0) as BotanasAcuVta,"
                + "cast(0.0 as float) as '%20',isnull(BotanasDiario.UtilBruta,0) as BotanasDiarioUtil,isnull(BotanasAcu.UtilBruta,0) as BotanasAcuUtil,cast(0.0 as float) as '%21',isnull(TotalDiario.unidades,0) as TotalDiarioUni,"
                + "isnull(TotalAcu.unidades,0) as TotalAcuUni,Cast(0.0 as float) as '%22',isnull(TotalDiario.Total,0) as TotalDiarioVta,isnull(TotalAcu.Total,0) as TotalAcuVta,cast(0.0 as float) as '%23',isnull(TotalDiario.UtilBruta,0) as TotalDiarioUtil,"
                + "isnull(TotalAcu.UtilBruta,0) as TotalAcuUtil,cast(0.0 as float) as '%24'  from "
                + "(((((((((((((((select cod_estab,nombre from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) as Sucursales "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BisuteriaDiario on Sucursales.cod_estab=BisuteriaDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BisuteriaAcu on Sucursales.cod_estab=BisuteriaAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as RopaDiario on Sucursales.cod_estab=RopaDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as RopaAcu on Sucursales.cod_estab=RopaAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as CalzadoDiario on Sucursales.cod_estab=CalzadoDiario.cod_estab) "
                + "left join  "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as CalzadoAcu on Sucursales.cod_estab=CalzadoAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as HogarDiario on Sucursales.cod_estab=HogarDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as HogarAcu on Sucursales.cod_estab=HogarAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ServiciosDiario on Sucursales.cod_estab=ServiciosDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ServiciosAcu on Sucursales.cod_estab=ServiciosAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BotanasDiario on Sucursales.cod_estab=BotanasDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BotanasAcu on Sucursales.cod_estab=BotanasAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BellezaDiario on Sucursales.cod_estab=BellezaDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BellezaAcu on Sucursales.cod_estab=BellezaAcu.cod_estab "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto in ('1','2','3','4','6','7','8') and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TotalDiario on Sucursales.cod_estab=TotalDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto in ('1','2','3','4','6','7','8') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TotalAcu on Sucursales.cod_estab=TotalAcu.cod_estab "
                + "order by CAST(Sucursales.cod_estab as int) asc";
                #endregion
                cmd.ExecuteNonQuery();
                cmd.CommandText = "update _TempVentas "
                + "set [%1]=(HogarDiarioUni/(case when (select sum(HogarDiarioUni) from _TempVentas)<>0 then (select sum(HogarDiarioUni) from _TempVentas) else null end))*100,"
                + "[%2]=(HogarDiarioVta/(case when (select sum(HogarDiarioVta) from _TempVentas)<>0 then (select sum(HogarDiarioVta) from _TempVentas) else null end))*100,"
                + "[%3]=(HogarDiarioUtil/(case when (select sum(HogarDiarioUtil) from _TempVentas)<>0 then (select sum(HogarDiarioUtil) from _TempVentas) else null  end))*100,"
                + "[%4]=(BisuteriaDiarioUni/(case when (select sum(BisuteriaDiarioUni) from _TempVentas)<>0 then (select sum(BisuteriaDiarioUni) from _TempVentas) else null end))*100,"
                + "[%5]=(BisuteriaDiarioVta/(case when (select sum(BisuteriaDiarioVta) from _TempVentas)<>0 then (select sum(BisuteriaDiarioVta) from _TempVentas) else null end))*100,"
                + "[%6]=(BisuteriaDiarioUtil/(case when (select sum(BisuteriaDiarioUtil) from _TempVentas)<>0 then (select sum(BisuteriaDiarioUtil) from _TempVentas) else null end))*100,"
                + "[%7]=(BellezaDiarioUni/(case when (select sum(BellezaDiarioUni) from _TempVentas)<>0 then (select sum(BellezaDiarioUni) from _TempVentas) else null end))*100,"
                + "[%8]=(BellezaDiarioVta/(case when (select sum(BellezaDiarioVta) from _TempVentas)<>0 then (select sum(BellezaDiarioVta) from _TempVentas) else null end))*100,"
                + "[%9]=(BellezaDiarioUtil/(case when (select sum(BellezaDiarioUtil) from _TempVentas)<>0 then (select sum(BellezaDiarioUtil) from _TempVentas) else null end))*100,"
                + "[%10]=(CalzadoDiarioUni/(case when (select sum(CalzadoDiarioUni) from _TempVentas)<>0 then (select sum(CalzadoDiarioUni) from _TempVentas) else null end))*100,"
                + "[%11]=(CalzadoDiarioVta/(case when (select sum(CalzadoDiarioVta) from _TempVentas)<>0 then (select sum(CalzadoDiarioVta) from _TempVentas) else null end))*100,"
                + "[%12]=(CalzadoDiarioUtil/(case when (select sum(CalzadoDiarioUtil) from _TempVentas)<>0 then (select sum(CalzadoDiarioUtil) from _TempVentas) else null end))*100,"
                + "[%13]=(RopaDiarioUni/(case when (select sum(RopaDiarioUni) from _TempVentas)<>0 then (select sum(RopaDiarioUni) from _TempVentas) else null end))*100,"
                + "[%14]=(RopaDiarioVta/(case when (select sum(RopaDiarioVta) from _TempVentas)<>0 then (select sum(RopaDiarioVta) from _TempVentas) else null end))*100,"
                + "[%15]=(RopaDiarioUtil/(case when (select sum(RopaDiarioUtil) from _TempVentas)<>0 then (select sum(RopaDiarioUtil) from _TempVentas) else null end))*100,"
                + "[%16]=(ServiciosDiarioUni/(case when (select sum(ServiciosDiarioUni) from _TempVentas)<>0 then (select sum(ServiciosDiarioUni) from _TempVentas) else null end))*100,"
                + "[%17]=(ServiciosDiarioVta/(case when (select sum(ServiciosDiarioVta) from _TempVentas)<>0 then (select sum(ServiciosDiarioVta) from _TempVentas) else null end))*100,"
                + "[%18]=(ServiciosDiarioUtil/(case when (select sum(ServiciosDiarioUtil) from _TempVentas)<>0 then (select sum(ServiciosDiarioUtil) from _TempVentas) else null end))*100,"
                + "[%19]=(BotanasDiarioUni/(case when (select sum(BotanasDiarioUni) from _TempVentas)<>0 then (select sum(BotanasDiarioUni) from _TempVentas) else null end))*100,"
                + "[%20]=(BotanasDiarioVta/(case when (select sum(BotanasDiarioVta) from _TempVentas)<>0 then (select sum(BotanasDiarioVta) from _TempVentas) else null end))*100,"
                + "[%21]=(BotanasDiarioUtil/(case when (select sum(BotanasDiarioUtil) from _TempVentas)<>0 then (select sum(BotanasDiarioUtil) from _TempVentas) else null end))*100,"
                + "[%22]=(TotalDiarioUni/(case when (select sum(TotalDiarioUni) from _TempVentas)<>0 then (select sum(TotalDiarioUni) from _TempVentas) else null end))*100,"
                + "[%23]=(TotalDiarioVta/(case when (select sum(TotalDiarioVta) from _TempVentas)<>0 then  (select sum(TotalDiarioVta) from _TempVentas) else null end))*100,"
                + "[%24]=(TotalDiarioUtil/(case when (select sum(TotalDiarioUtil) from _TempVentas)<>0 then (select sum(TotalDiarioUtil) from _TempVentas) else null end))*100";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "select _TempVentas.* from _TempVentas inner join establecimientos on _TempVentas.cod_estab=establecimientos.cod_estab "
                + "order by establecimientos.tipo_establecimiento desc,establecimientos.grupo_establecimiento asc,cast(establecimientos.cod_estab as int)";
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                da.Fill(dt);
                DG1.DataSource = dt;
                DG1.SelectAll();
                object objeto = DG1.GetClipboardContent();
                Microsoft.Office.Interop.Excel.Application excel;
                excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook libro;
                libro = excel.Workbooks.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                Worksheet hoja = new Worksheet();
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                hoja.Name = "VENTA DIARIA";
                Microsoft.Office.Interop.Excel.Range rango;
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja.Cells[1, 2] = "REPORTE DE VENTA DIARIA";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "BW1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B3", "B5");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "CODIGO";
                    rango = (Range)hoja.get_Range("C3", "C5");
                    rango.Select();
                    rango.Merge();
                    //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "SUCURSAL";

                    for (int i = 4; i <= 75; i += 9)
                    {
                        //rango = (Range)hoja.get_Range("4,4","12,4");
                        rango = (Range)hoja.get_Range(sCol(i) + "3", sCol(i + 8) + "3");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        switch (i)
                        {
                            case 4:
                                rango.Cells[1, 1] = " H O G A R";
                                break;
                            case 13:
                                rango.Cells[1, 1] = "B I S U T E R I A";
                                break;
                            case 22:
                                rango.Cells[1, 1] = "B E L L E Z A";
                                break;
                            case 31:
                                rango.Cells[1, 1] = "C A L Z A D O";
                                break;
                            case 40:
                                rango.Cells[1, 1] = "R O P A";
                                break;
                            case 49:
                                rango.Cells[1, 1] = "S E R V I C I O S";
                                break;
                            case 58:
                                rango.Cells[1, 1] = "B O T A N A S   Y   S N A C K S";
                                break;
                            case 67:
                                rango.Cells[1, 1] = "T O T A L   D I A R I O";
                                break;
                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            rango = (Range)hoja.get_Range(sCol(j) + "4", sCol(j + 2) + "4");
                            rango.Select();
                            rango.Merge();
                            rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            rango.Cells.Font.FontStyle = "Bold";
                            if (j == i)
                            {
                                rango.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                rango.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                rango.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k] = "DIARIAS";
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "5", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "5", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }

                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k] = "ACUMULADAS";
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "5", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "5", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "%";
                                    hoja.Range[sCol(k) + "5", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                hoja.Cells[5, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";
                                hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "5:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                            }
                        }
                    }

                    hoja.Range["B6", "C" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", "BW" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    hoja.Range["B6", "BW" + Convert.ToString(DG1.Rows.Count + 6)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    hoja.Cells[DG1.Rows.Count + 6, 3] = "T O T A L";
                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "BW" + Convert.ToString(DG1.Rows.Count + 10));
                    rango.EntireColumn.AutoFit();
                }
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //
                //                                          COMPARATIVO DIARIO
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Clipboard.Clear();
                objeto = null;
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                #region ConsultaComparativoDiario
                cmd.CommandText = "select Sucursales.cod_estab,Sucursales.nombre,isnull(HogarDiario.unidades,0) as HogarDiarioUni,isnull(HogarBase.unidades,0) as HogarBaseUni,"
                + "[Inc o Dec1]=case when HogarBase.unidades=0 then 0 when HogarBase.unidades is null then 0when HogarBase.unidades>0 then ((HogarDiario.unidades/HogarBase.unidades)-1)*100	end,"
                + "isnull(HogarDiario.Total,0) as HogarDiarioVta,isnull(HogarBase.Total,0) as HogarBaseVta,[Inc o Dec2]=case	when HogarBase.Total=0 then 0 when HogarBase.Total is null then 0 when HogarBase.Total>0 then ((HogarDiario.Total/HogarBase.Total)-1)*100 end,"
                + "isnull(HogarDiario.UtilBruta,0) as HogarDiarioUtil,isnull(HogarBase.UtilBruta,0) as HogarBaseUtil,[Inc o Dec3]=case when HogarBase.UtilBruta=0 then 0	when HogarBase.UtilBruta is null then 0	when HogarBase.UtilBruta>0 then ((HogarDiario.UtilBruta/HogarBase.UtilBruta)-1)*100	end,"
                + "isnull(BisuteriaDiario.unidades,0) as BisuteriaDiarioUni,	isnull(BisuteriaBase.unidades,0) as BisuteriaBaseUni,[Inc o Dec4]=case when BisuteriaBase.unidades=0 then 0	when BisuteriaBase.unidades is null then 0 when BisuteriaBase.unidades>0 then ((BisuteriaDiario.unidades/BisuteriaBase.unidades)-1)*100	end,"
                + "isnull(BisuteriaDiario.Total,0) as BisuteriaDiarioVta,isnull(BisuteriaBase.Total,0) as BisuteriaBaseVta,[Inc o Dec5]=case	when BisuteriaBase.Total=0 then 0 when BisuteriaBase.Total is null then 0 when BisuteriaBase.Total>0 then ((BisuteriaDiario.Total/BisuteriaBase.Total)-1)*100 end,"
                + "isnull(BisuteriaDiario.UtilBruta,0) as BisuteriaDiarioUtil,isnull(BisuteriaBase.UtilBruta,0) as BisuteriaBaseUtil,[Inc o Dec6]=case when BisuteriaBase.UtilBruta=0 then 0 when BisuteriaBase.UtilBruta is null then 0 when BisuteriaBase.UtilBruta>0 then ((BisuteriaDiario.UtilBruta/BisuteriaBase.UtilBruta)-1)*100	end,"
                + "isnull(BellezaDiario.unidades,0) as BellezaDiarioUni,isnull(BellezaBase.unidades,0) as BellezaBaseUni,[Inc o Dec7]=case when BellezaBase.unidades=0 then 0 when BellezaBase.unidades is null then 0 when BellezaBase.unidades>0 then ((BellezaDiario.unidades/BellezaBase.unidades)-1)*100 end,"
                + "isnull(BellezaDiario.Total,0) as BellezaDiarioVta,isnull(BellezaBase.Total,0) as BellezaBaseVta,[Inc o Dec8]=case when BellezaBase.Total=0 then 0 when BellezaBase.Total is null then 0 when BellezaBase.Total>0 then ((BellezaDiario.Total/BellezaBase.Total)-1)*100	end,"
                + "isnull(BellezaDiario.UtilBruta,0) as BellezaDiarioUtil,isnull(BellezaBase.UtilBruta,0) as BellezaBaseUtil,[Inc o Dec9]=case when BellezaBase.UtilBruta=0 then 0 when BellezaBase.UtilBruta is null then 0	when BellezaBase.UtilBruta>0 then ((BellezaDiario.UtilBruta/BellezaBase.UtilBruta)-1)*100 end,"
                + "isnull(CalzadoDiario.unidades,0) as CalzadoDiarioUni,isnull(CalzadoBase.unidades,0) as CalzadoBaseUni,[Inc o Dec10]=case when CalzadoBase.unidades=0 then 0 when CalzadoBase.unidades is null then 0 when CalzadoBase.unidades>0 then ((CalzadoDiario.unidades/CalzadoBase.unidades)-1)*100 end,"
                + "isnull(CalzadoDiario.Total,0) as CalzadoDiarioVta, isnull(CalzadoBase.Total,0) as CalzadoBaseVta,[Inc o Dec11]=case when CalzadoBase.Total=0 then 0 when CalzadoBase.Total is null then 0 when CalzadoBase.Total>0 then ((CalzadoDiario.Total/CalzadoBase.Total)-1)*100 end,"
                + "isnull(CalzadoDiario.UtilBruta,0) as CalzadoDiarioUtil,isnull(CalzadoBase.UtilBruta,0) as CalzadoBaseUtil, [Inc o Dec12]=case	when CalzadoBase.UtilBruta=0 then 0	when CalzadoBase.UtilBruta is null then 0 when CalzadoBase.UtilBruta>0 then ((CalzadoDiario.UtilBruta/CalzadoBase.UtilBruta)-1)*100	end,"
                + "isnull(RopaDiario.unidades,0) as RopaDiarioUni,isnull(RopaBase.unidades,0) as RopaBaseUni,[Inc o Dec13]=case when RopaBase.unidades=0 then 0 when RopaBase.unidades is null then 0 when RopaBase.unidades>0 then ((RopaDiario.unidades/RopaBase.unidades)-1)*100 end,"
                + "isnull(RopaDiario.Total,0) as RopaDiarioVta,isnull(RopaBase.Total,0) as RopaBaseVta,[Inc o Dec14]=case when RopaBase.Total=0 then 0 when RopaBase.Total is null then 0 when RopaBase.Total>0 then ((RopaDiario.Total/RopaBase.Total)-1)*100 end,"
                + "isnull(RopaDiario.UtilBruta,0) as RopaDiarioUtil,	isnull(RopaBase.UtilBruta,0) as RopaBaseUtil,[Inc o Dec15]=case when RopaBase.UtilBruta=0 then 0 when RopaBase.UtilBruta is null then 0	when RopaBase.UtilBruta>0 then ((RopaDiario.UtilBruta/RopaBase.UtilBruta)-1)*100 end,"
                + "isnull(ServiciosDiario.unidades,0) as ServiciosDiarioUni,	isnull(ServiciosBase.unidades,0) as ServiciosBaseUni,[Inc o Dec16]=case	when ServiciosBase.unidades=0 then 0 when ServiciosBase.unidades is null then 0	when ServiciosBase.unidades>0 then ((ServiciosDiario.unidades/ServiciosBase.unidades)-1)*100 end,"
                + "isnull(ServiciosDiario.Total,0) as ServiciosDiarioVta,isnull(ServiciosBase.Total,0) as ServiciosBaseVta,[Inc o Dec17]=case when ServiciosBase.Total=0 then 0 when ServiciosBase.Total is null then 0 when ServiciosBase.Total>0 then ((ServiciosDiario.Total/ServiciosBase.Total)-1)*100 end,"
                + "isnull(ServiciosDiario.UtilBruta,0) as ServiciosDiarioUtil,isnull(ServiciosBase.UtilBruta,0) as ServiciosBaseUtil,[Inc o Dec18]=case when ServiciosBase.UtilBruta=0 then 0 when ServiciosBase.UtilBruta is null then 0 when ServiciosBase.UtilBruta>0 then ((ServiciosDiario.UtilBruta/ServiciosBase.UtilBruta)-1)*100 end,"
                + "isnull(BotanasDiario.unidades,0) as BotanasDiarioUni, isnull(BotanasBase.unidades,0) as BotanasBaseUni,[Inc o Dec19]=case when BotanasBase.unidades=0 then 0 when BotanasBase.unidades is null then 0 when BotanasBase.unidades>0 then ((BotanasDiario.unidades/BotanasBase.unidades)-1)*100 end,"
                + "isnull(BotanasDiario.Total,0) as BotanasDiarioVta,isnull(BotanasBase.Total,0) as BotanasBaseVta,[Inc o Dec20]=case when BotanasBase.Total=0 then 0 when BotanasBase.Total is null then 0 when BotanasBase.Total>0 then ((BotanasDiario.Total/BotanasBase.Total)-1)*100 end,"
                + "isnull(BotanasDiario.UtilBruta,0) as BotanasDiarioUtil,isnull(BotanasBase.UtilBruta,0) as BotanasBaseUtil,[Inc o Dec21]=case	when BotanasBase.UtilBruta=0 then 0	when BotanasBase.UtilBruta is null then 0 when BotanasBase.UtilBruta>0 then ((BotanasDiario.UtilBruta/BotanasBase.UtilBruta)-1)*100 end "
                + "/*into ##TablaVentaCompDiario*/ from "
                + "((((((((((((((select cod_estab,nombre,tipo_establecimiento,grupo_establecimiento from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) as Sucursales "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BisuteriaDiario on Sucursales.cod_estab=BisuteriaDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BisuteriaBase on Sucursales.cod_estab=BisuteriaBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as RopaDiario on Sucursales.cod_estab=RopaDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as RopaBase on Sucursales.cod_estab=RopaBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as CalzadoDiario on Sucursales.cod_estab=CalzadoDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as CalzadoBase on Sucursales.cod_estab=CalzadoBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as HogarDiario on Sucursales.cod_estab=HogarDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as HogarBase on Sucursales.cod_estab=HogarBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ServiciosDiario on Sucursales.cod_estab=ServiciosDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(entysalVentas.Importe-(entysalVentas.costo)) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ServiciosBase on Sucursales.cod_estab=ServiciosBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BotanasDiario on Sucursales.cod_estab=BotanasDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BotanasBase on Sucursales.cod_estab=BotanasBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BellezaDiario on Sucursales.cod_estab=BellezaDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BellezaBase on Sucursales.cod_estab=BellezaBase.cod_estab "
                + "order by sucursales.tipo_establecimiento desc,sucursales.grupo_establecimiento asc, CAST(Sucursales.cod_estab as int) asc";
                #endregion
                //cmd.ExecuteNonQuery();
                //cmd.CommandText="select ##TablaVentaCompDiario.* from ##TablaVentaCompDiario inner join establecimientos on ##TablaVentaCompDiario.cod_estab=establecimientos.cod_estab "
                //+"order by establecimientos.tipo_establecimiento desc,establecimientos.grupo_establecimiento asc,cast(establecimientos.cod_estab as int)";
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                DG1.DataSource = dt;
                DG1.SelectAll();
                objeto = DG1.GetClipboardContent();
                if (objeto != null)
                {

                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Sheets.get_Item(2);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO DIARIO";
                    hoja.Cells[1, 2] = "COMPARATIVO DIARIO";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "BN1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B2", "BN2");
                    rango.Select();
                    rango.Merge();
                    rango.Cells[1.1, Type.Missing] = dia.ToString("d MMM yyyy");
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B3", "B5");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "CODIGO";
                    rango = (Range)hoja.get_Range("C3", "C5");
                    rango.Select();
                    rango.Merge();
                    //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "SUCURSAL";
                    for (int i = 4; i <= 66; i += 9)
                    {
                        //rango = (Range)hoja.get_Range("4,4","12,4");
                        rango = (Range)hoja.get_Range(sCol(i) + "3", sCol(i + 8) + "3");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        switch (i)
                        {
                            case 4:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO HOGAR";
                                break;
                            case 13:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO BISUTERIA";
                                break;
                            case 22:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO BELLEZA";
                                break;
                            case 31:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO CALZADO";
                                break;
                            case 40:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO ROPA";
                                break;
                            case 49:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO SERVICIOS";
                                break;
                            case 58:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO SNACKS Y BOTANAS";
                                break;
                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            rango = (Range)hoja.get_Range(sCol(j) + "4", sCol(j + 1) + "4");
                            rango.Select();
                            rango.Merge();
                            rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            rango.Cells.Font.FontStyle = "Bold";
                            hoja.Cells[4, j + 2] = "%";
                            hoja.Cells[4, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            hoja.Cells[4, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            hoja.Cells[4, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[4, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            if (j == i)
                            {
                                rango.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                rango.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                rango.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = dia.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = diabase.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "Inc o Dec";
                                    hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 7)].NumberFormat = "#,###,##0.00";
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 6) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 6) + ")-1)*100";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                hoja.Cells[5, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";

                            }
                        }
                    }
                    hoja.Range["B6", "C" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", "BN" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    hoja.Range["B6", "BN" + Convert.ToString(DG1.Rows.Count + 6)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    hoja.Cells[DG1.Rows.Count + 6, 3] = "T O T A L";
                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "BN" + Convert.ToString(DG1.Rows.Count + 10));
                    rango.EntireColumn.AutoFit();
                }
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //
                //                                          COMPARATIVO ACUMULADO
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Clipboard.Clear();
                objeto = null;
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                #region ComparativoAcumuladoMayoristas
                cmd.CommandText = "select Sucursales.cod_estab,Sucursales.nombre, "
                + "isnull(HogarAcu.unidades,0) as HogarAcuUni,isnull(HogarAcuBase.unidades,0) as HogarAcuBaseUni,[Inc o Dec1]=case when HogarAcuBase.unidades=0 then 0 when HogarAcuBase.unidades is null then 0 when HogarAcuBase.unidades>0 then ((HogarAcu.unidades/HogarAcuBase.unidades)-1)*100 end,"
                + "isnull(HogarAcu.Total,0) as BlisterDiarioVta,isnull(HogarAcuBase.Total,0) as BlisterBaseVta,[Inc o Dec2]=case when HogarAcuBase.Total=0 then 0 when HogarAcuBase.Total is null then 0	when HogarAcuBase.Total>0 then ((HogarAcu.Total/HogarAcuBase.Total)-1)*100 end,"
                + "isnull(HogarAcu.UtilBruta,0) as HogarAcuUtil,	isnull(HogarAcuBase.UtilBruta,0) as HogarAcuBaseUtil,[Inc o Dec3]=case when HogarAcuBase.UtilBruta=0 then 0	when HogarAcuBase.UtilBruta is null then 0 when HogarAcuBase.UtilBruta>0 then ((HogarAcu.UtilBruta/HogarAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(BisuteriaAcu.unidades,0) as BisuteriaAcuUni,isnull(BisuteriaAcuBase.unidades,0) as BisuteriaAcuBaseUni,[Inc o Dec4]=case when BisuteriaAcuBase.unidades=0 then 0 when BisuteriaAcuBase.unidades is null then 0 when BisuteriaAcuBase.unidades>0 then ((BisuteriaAcu.unidades/BisuteriaAcuBase.unidades)-1)*100 end,"
                + "isnull(BisuteriaAcu.Total,0) as BisuteriaAcuVta,isnull(BisuteriaAcuBase.Total,0) as BisuteriaAcuBaseVta,[Inc o Dec5]=case	when BisuteriaAcuBase.Total=0 then 0 when BisuteriaAcuBase.Total is null then 0 when BisuteriaAcuBase.Total>0 then ((BisuteriaAcu.Total/BisuteriaAcuBase.Total)-1)*100 end,"
                + "isnull(BisuteriaAcu.UtilBruta,0) as BisuteriaAcuUtil,	isnull(BisuteriaAcuBase.UtilBruta,0) as BisuteriaAcuBaseUtil,[Inc o Dec6]=case when BisuteriaAcuBase.UtilBruta=0 then 0 when BisuteriaAcuBase.UtilBruta is null then 0 when BisuteriaAcuBase.UtilBruta>0 then ((BisuteriaAcu.UtilBruta/BisuteriaAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(BellezaAcu.unidades,0) as BellezaAcuUni,isnull(BellezaAcuBase.unidades,0) as BellezaAcuBaseUni,[Inc o Dec7]=case when BellezaAcuBase.unidades=0 then 0 when BellezaAcuBase.unidades is null then 0 when BellezaAcuBase.unidades>0 then ((BellezaAcu.unidades/BellezaAcuBase.unidades)-1)*100 end,"
                + "isnull(BellezaAcu.Total,0) as BellezaAcuVta,isnull(BellezaAcuBase.Total,0) as BellezaAcuBaseVta,[Inc o Dec8]=case when BellezaAcuBase.Total=0 then 0 when BellezaAcuBase.Total is null then 0 when BellezaAcuBase.Total>0 then ((BellezaAcu.Total/BellezaAcuBase.Total)-1)*100 end,"
                + "isnull(BellezaAcu.UtilBruta,0) as BellezaAcuUtil,	isnull(BellezaAcuBase.UtilBruta,0) as BellezaAcuBaseUtil,[Inc o Dec9]=case when BellezaAcuBase.UtilBruta=0 then 0 when BellezaAcuBase.UtilBruta is null then 0 when BellezaAcuBase.UtilBruta>0 then ((BellezaAcu.UtilBruta/BellezaAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(CalzadoAcu.unidades,0) as CalzadoAcuUni,isnull(CalzadoAcuBase.unidades,0) as CalzadoAcuBaseUni,[Inc o Dec10]=case when CalzadoAcuBase.unidades=0 then 0 when CalzadoAcuBase.unidades is null then 0 when CalzadoAcuBase.unidades>0 then ((CalzadoAcu.unidades/CalzadoAcuBase.unidades)-1)*100 end,"
                + "isnull(CalzadoAcu.Total,0) as CalzadoAcuVta, isnull(CalzadoAcuBase.Total,0) as CalzadoAcuBaseVta,[Inc o Dec11]=case when CalzadoAcuBase.Total=0 then 0 when CalzadoAcuBase.Total is null then 0 when CalzadoAcuBase.Total>0 then ((CalzadoAcu.Total/CalzadoAcuBase.Total)-1)*100 end,"
                + "isnull(CalzadoAcu.UtilBruta,0) as CalzadoAcuUtil,	isnull(CalzadoAcuBase.UtilBruta,0) as CalzadoAcuBaseUtil,[Inc o Dec12]=case when CalzadoAcuBase.UtilBruta=0 then 0 when CalzadoAcuBase.UtilBruta is null then 0 when CalzadoAcuBase.UtilBruta>0 then ((CalzadoAcu.UtilBruta/CalzadoAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(RopaAcu.unidades,0) as NaviDiarioUni,isnull(RopaAcuBase.unidades,0) as NaviBaseUni,[Inc o Dec13]=case when RopaAcuBase.unidades=0 then 0 when RopaAcuBase.unidades is null then 0 when RopaAcuBase.unidades>0 then ((RopaAcu.unidades/RopaAcuBase.unidades)-1)*100 end,"
                + "isnull(RopaAcu.Total,0) as NaviDiarioVta,isnull(RopaAcuBase.Total,0) as NaviBaseVta,[Inc o Dec14]=case when RopaAcuBase.Total=0 then 0 when RopaAcuBase.Total is null then 0 when RopaAcuBase.Total>0 then ((RopaAcu.Total/RopaAcuBase.Total)-1)*100 end,"
                + "isnull(RopaAcu.UtilBruta,0) as NaviDiarioUtil,isnull(RopaAcuBase.UtilBruta,0) as NaviBaseUtil,[Inc o Dec15]=case when RopaAcuBase.UtilBruta=0 then 0 when RopaAcuBase.UtilBruta is null then 0 when RopaAcuBase.UtilBruta>0 then ((RopaAcu.UtilBruta/RopaAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(ServiciosAcu.unidades,0) as ServiciosAcuUni,isnull(ServiciosAcuBase.unidades,0) as ServiciosAcuBaseUni,[Inc o Dec16]=case when ServiciosAcuBase.unidades=0 then 0 when ServiciosAcuBase.unidades is null then 0 when ServiciosAcuBase.unidades>0 then ((ServiciosAcu.unidades/ServiciosAcuBase.unidades)-1)*100 end,"
                + "isnull(ServiciosAcu.Total,0) as ServiciosAcuVta,isnull(ServiciosAcuBase.Total,0) as ServiciosAcuBaseVta,[Inc o Dec17]=case when ServiciosAcuBase.Total=0 then 0 when ServiciosAcuBase.Total is null then 0 when ServiciosAcuBase.Total>0 then ((ServiciosAcu.Total/ServiciosAcuBase.Total)-1)*100 end,"
                + "isnull(ServiciosAcu.UtilBruta,0) as ServiciosAcuUtil,isnull(ServiciosAcuBase.UtilBruta,0) as ServiciosAcuBaseUtil,[Inc o Dec18]=case when ServiciosAcuBase.UtilBruta=0 then 0 when ServiciosAcuBase.UtilBruta is null then 0 when ServiciosAcuBase.UtilBruta>0 then ((ServiciosAcu.UtilBruta/ServiciosAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(BotanasAcu.unidades,0) as BotanasAcuUni,isnull(BotanasAcuBase.unidades,0) as BotanasAcuBaseUni,[Inc o Dec19]=case when BotanasAcuBase.unidades=0 then 0 when BotanasAcuBase.unidades is null then 0 when BotanasAcuBase.unidades>0 then ((BotanasAcu.unidades/BotanasAcuBase.unidades)-1)*100 end,"
                + "isnull(BotanasAcu.Total,0) as BotanasAcuVta,isnull(BotanasAcuBase.Total,0) as BotanasAcuBaseVta,[Inc o Dec20]= case when BotanasAcuBase.Total=0 then 0 when BotanasAcuBase.Total is null then 0 when BotanasAcuBase.Total>0 then ((BotanasAcu.Total/BotanasAcuBase.Total)-1)*100 end,"
                + "isnull(BotanasAcu.UtilBruta,0) as BotanasAcuUtil,	isnull(BotanasAcuBase.UtilBruta,0) as BotanasAcuBaseUtil,[Inc o Dec21]=case when BotanasAcuBase.UtilBruta=0 then 0	when BotanasAcuBase.UtilBruta is null then 0 when BotanasAcuBase.UtilBruta>0 then ((BotanasAcu.UtilBruta/BotanasAcuBase.UtilBruta)-1)*100 end "
                + "/*into ##TablaVentaCompAcumulado*/ from ((((((((((((((select cod_estab,nombre,tipo_establecimiento,grupo_establecimiento from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) as Sucursales "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BisuteriaAcu on Sucursales.cod_estab=BisuteriaAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BisuteriaAcuBase on Sucursales.cod_estab=BisuteriaAcuBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as RopaAcu on Sucursales.cod_estab=RopaAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as RopaAcuBase on Sucursales.cod_estab=RopaAcuBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as CalzadoAcu on Sucursales.cod_estab=CalzadoAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as CalzadoAcuBase on Sucursales.cod_estab=CalzadoAcuBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as HogarAcu on Sucursales.cod_estab=HogarAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as HogarAcuBase on Sucursales.cod_estab=HogarAcuBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ServiciosAcu on Sucursales.cod_estab=ServiciosAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(entysalVentas.Importe-(entysalVentas.costo)) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ServiciosAcuBase on Sucursales.cod_estab=ServiciosAcuBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BotanasAcu on Sucursales.cod_estab=BotanasAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BotanasAcuBase on Sucursales.cod_estab=BotanasAcuBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BellezaAcu on Sucursales.cod_estab=BellezaAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as BellezaAcuBase on Sucursales.cod_estab=BellezaAcuBase.cod_estab "
                + "order  by sucursales.tipo_establecimiento desc,sucursales.grupo_establecimiento asc, CAST(Sucursales.cod_estab as int) asc";
                #endregion
                //cmd.ExecuteNonQuery();        
                //cmd.CommandText = "select ##TablaVentaCompAcumulado.* from ##TablaVentaCompAcumulado inner join establecimientos on ##TablaVentaCompAcumulado.cod_estab=establecimientos.cod_estab "
                //+ "order by establecimientos.tipo_establecimiento desc,establecimientos.grupo_establecimiento asc,cast(establecimientos.cod_estab as int)";            
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                DG1.DataSource = dt;
                DG1.SelectAll();
                objeto = DG1.GetClipboardContent();
                if (objeto != null)
                {

                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Sheets.get_Item(3);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO ACUMULADO";
                    hoja.Cells[1, 2] = "COMPARATIVO ACUMULADO";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "BN1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B2", "BN2");
                    rango.Select();
                    rango.Merge();
                    rango.Cells[1.1, Type.Missing] = "Del " + primero.ToString("d MMM yyyy") + " al " + dia.ToString("d MMM yyyy");
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B3", "B5");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "CODIGO";
                    rango = (Range)hoja.get_Range("C3", "C5");
                    rango.Select();
                    rango.Merge();
                    //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "SUCURSAL";
                    for (int i = 4; i <= 66; i += 9)
                    {
                        //rango = (Range)hoja.get_Range("4,4","12,4");
                        rango = (Range)hoja.get_Range(sCol(i) + "3", sCol(i + 8) + "3");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        switch (i)
                        {
                            case 4:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO HOGAR";
                                break;
                            case 13:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO BISUTERIA";
                                break;
                            case 22:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO BELLEZA";
                                break;
                            case 31:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO CALZADO";
                                break;
                            case 40:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO ROPA";
                                break;
                            case 49:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO SERVICIOS";
                                break;
                            case 58:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO SNACKS Y BOTANAS";
                                break;

                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            rango = (Range)hoja.get_Range(sCol(j) + "4", sCol(j + 1) + "4");
                            rango.Select();
                            rango.Merge();
                            rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            rango.Cells.Font.FontStyle = "Bold";
                            hoja.Cells[4, j + 2] = "%";
                            hoja.Cells[4, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            hoja.Cells[4, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            hoja.Cells[4, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[4, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            if (j == i)
                            {
                                rango.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                rango.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                rango.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = dia.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = diabase.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "Inc o Dec";
                                    hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00";
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 6) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 6) + ")-1)*100";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                hoja.Cells[5, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";

                            }
                        }
                    }
                    hoja.Range["B6", "C" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", "BN" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    hoja.Range["B6", "BN" + Convert.ToString(DG1.Rows.Count + 6)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    hoja.Cells[DG1.Rows.Count + 6, 3] = "T O T A L";
                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "BN" + Convert.ToString(DG1.Rows.Count + 10));
                    rango.EntireColumn.AutoFit();
                }
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //
                //                                          COMPARATIVO TOTAL DIARIO Y ACUMULADO
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Clipboard.Clear();
                objeto = null;
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                #region ComparativoTotalSQL
                cmd.CommandText = "select Sucursales.cod_estab,Sucursales.nombre,"
                + "isnull(TotalDiario.unidades,0) as TotalDiarioUni,isnull(TotalBase.unidades,0) as TotalDiarioBaseUni,[Inc o Dec1]=case when TotalBase.unidades=0 then 0 when TotalBase.unidades is null then 0 when TotalBase.unidades>0 then ((TotalDiario.unidades/TotalBase.unidades)-1)*100 end,"
                + "isnull(TotalDiario.Total,0) as TotalDiarioVta,isnull(TotalBase.Total,0) as TotalDiarioBaseVta,[Inc o Dec2]=case when TotalBase.Total=0 then 0 when TotalBase.Total is null then 0 when TotalBase.Total>0 then ((TotalDiario.Total/TotalBase.Total)-1)*100 end,"
                + "isnull(TotalDiario.UtilBruta,0) as TotalDiarioUtil,isnull(TotalBase.UtilBruta,0) as TotalDiarioBaseUtil,[Inc o Dec3]=case when TotalBase.UtilBruta=0 then 0 when TotalBase.UtilBruta is null then 0 when TotalBase.UtilBruta>0 then ((TotalDiario.UtilBruta/TotalBase.UtilBruta)-1)*100 end,"
                + "isnull(TotalAcu.unidades,0) as TotalAcuUni,isnull(TotalAcuBase.unidades,0) as TotalAcuBaseUni,[Inc o Dec4]=case when TotalAcuBase.unidades=0 then 0 when TotalAcuBase.unidades is null then 0 when TotalAcuBase.unidades>0 then ((TotalAcu.unidades/TotalAcuBase.unidades)-1)*100 end,"
                + "isnull(TotalAcu.Total,0) as TotalAcuVta,isnull(TotalAcuBase.Total,0) as TotalAcuBaseVta,[Inc o Dec5]=case when TotalAcuBase.Total=0 then 0 when TotalAcuBase.Total is null then 0 when TotalAcuBase.Total>0 then ((TotalAcu.Total/TotalAcuBase.Total)-1)*100 end,"
                + "isnull(TotalAcu.UtilBruta,0) as TotalAcuUtil,isnull(TotalAcuBase.UtilBruta ,0) TotalAcuBaseUtil,[Inc o Dec6]=case when TotalAcuBase.UtilBruta=0 then 0 when TotalAcuBase.UtilBruta is null then 0 when TotalAcuBase.UtilBruta>0 then ((TotalAcu.UtilBruta/TotalAcuBase.UtilBruta)-1)*100 end "
                + "/*into ##TablaVentaCompTotal*/ from (((((select cod_estab,nombre,tipo_establecimiento,grupo_establecimiento from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) as Sucursales "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto in ('1','2','3','4','6','7','8') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between  '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TotalDiario on Sucursales.cod_estab=TotalDiario.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto in ('1','2','3','4','6','7','8') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TotalBase on Sucursales.cod_estab=TotalBase.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock)inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where p.linea_producto in ('1','2','3','4','6','7','8') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between  '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TotalAcu on Sucursales.cod_estab=TotalAcu.cod_estab) "
                + "left join "
                + "(select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod  "
                + "where p.linea_producto in ('1','2','3','4','6','7','8') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TotalAcuBase on Sucursales.cod_estab=TotalAcuBase .cod_estab) "
                + "order by sucursales.tipo_establecimiento desc,sucursales.grupo_establecimiento asc,CAST(Sucursales.cod_estab as int) asc";
                #endregion
                //cmd.ExecuteNonQuery();

                //cmd.CommandText = "select ##TablaVentaCompTotal.* from ##TablaVentaCompTotal inner join establecimientos on ##TablaVentaCompTotal.cod_estab=establecimientos.cod_estab "
                //+"order by establecimientos.tipo_establecimiento desc,establecimientos.grupo_establecimiento asc,cast(establecimientos.cod_estab as int)";
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                DG1.DataSource = dt;
                DG1.SelectAll();
                objeto = DG1.GetClipboardContent();
                if (objeto != null)
                {

                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Sheets.get_Item(4);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO TOTAL";
                    hoja.Cells[1, 2] = "COMPARATIVO TOTAL";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "U1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B2", "U2");
                    rango.Select();
                    rango.Merge();
                    rango.Cells[1.1, Type.Missing] = "Del " + primero.ToString("d MMM yyyy") + " al " + dia.ToString("d MMM yyyy");
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B3", "B5");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "CODIGO";
                    rango = (Range)hoja.get_Range("C3", "C5");
                    rango.Select();
                    rango.Merge();
                    //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "SUCURSAL";
                    for (int i = 4; i <= 21; i += 9)
                    {
                        //rango = (Range)hoja.get_Range("4,4","12,4");
                        rango = (Range)hoja.get_Range(sCol(i) + "3", sCol(i + 8) + "3");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        switch (i)
                        {
                            case 4:
                                rango.Cells[1, 1] = "COMPARATIVO TOTAL DIARIO";
                                break;
                            case 13:
                                rango.Cells[1, 1] = "COMPARATIVO TOTAL ACUMULADO";
                                break;

                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            rango = (Range)hoja.get_Range(sCol(j) + "4", sCol(j + 1) + "4");
                            rango.Select();
                            rango.Merge();
                            rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            rango.Cells.Font.FontStyle = "Bold";
                            hoja.Cells[4, j + 2] = "%";
                            hoja.Cells[4, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            hoja.Cells[4, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            hoja.Cells[4, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[4, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            if (j == i)
                            {
                                rango.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                rango.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                rango.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = dia.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = diabase.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "Inc o Dec";
                                    hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00";
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 6) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 6) + ")-1)*100";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                hoja.Cells[5, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";

                            }
                        }
                    }
                    hoja.Range["B6", "C" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", "U" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    hoja.Range["B6", "U" + Convert.ToString(DG1.Rows.Count + 6)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    hoja.Cells[DG1.Rows.Count + 6, 3] = "T O T A L";
                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "U" + Convert.ToString(DG1.Rows.Count + 10));
                    rango.EntireColumn.AutoFit();
                }
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //
                //                                                    COMPARATIVO POR CLASIFICACION
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Clipboard.Clear();
                objeto = null;
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                #region ComparativoClasificacion
                cmd.CommandText = "select clasificaciones.clasificacion_productos as producto,clasificaciones.nombre,"
                + "isnull(diario.unidades,0) as DiarioUni,isnull(basediario.unidades,0) as BaseDiarioUni,[Inc o Dec1]=case when basediario.unidades=0 then 0 when basediario.unidades is null then 0 when basediario.unidades>0 then ((diario.unidades/basediario.unidades)-1)*100 end,"
                + "isnull(diario.Total,0) as DiarioTotal,isnull(basediario.Total,0) as BaseDiarioTotal,[Inc o Dec2]=case when basediario.Total=0 then 0 when basediario.total is null then 0 when basediario.Total>0 then ((diario.Total/basediario.Total)-1)*100 end,"
                + "isnull(diario.UtilBruta,0) as DiarioUtil,isnull(basediario.UtilBruta,0) as BaseDiarioUtil,[Inc o Dec3]=case when basediario.UtilBruta=0 then 0 when basediario.UtilBruta is null then 0 when basediario.UtilBruta>0 then ((diario.UtilBruta/basediario.UtilBruta)-1)*100 end,"
                + "isnull(acumulado.unidades,0) AcuUni,isnull(baseacumulado.unidades,0) as BaseAcuUni,[Inc o Dec4]=case when baseacumulado.unidades=0 then 0 when baseacumulado.unidades is null then 0 when baseacumulado.unidades>0 then ((acumulado.unidades/baseacumulado.unidades)-1)*100 end,"
                + "isnull(acumulado.Total,0) as AcuTotal,isnull(baseacumulado.Total,0) as BaseAcuTotal,[Inc o Dec5]=case when baseacumulado.Total=0 then 0 when baseacumulado.total is null then 0 when baseacumulado.Total>0 then ((acumulado.Total/baseacumulado.Total)-1)*100 end,"
                + "isnull(acumulado.UtilBruta,0) as AcuUtilBruta,isnull(baseacumulado.UtilBruta,0) as BaseAcuBruta,[Inc o Dec6] =case when baseacumulado.UtilBruta=0 then 0 when baseacumulado.UtilBruta is null then 0 when baseacumulado.UtilBruta>0 then ((acumulado.UtilBruta/baseacumulado.UtilBruta)-1)*100 end "
                + "from (Select clasificacion_productos,nombre from clasificaciones_productos) as clasificaciones "
                + "left join (select p.clasificacion_productos,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and entysalVentas.cod_estab in (select cod_estab from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by p.clasificacion_productos) as diario on clasificaciones.clasificacion_productos=diario.clasificacion_productos "
                + "left join (select p.clasificacion_productos,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and entysalVentas.cod_estab in (select cod_estab from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by p.clasificacion_productos) as basediario on basediario.clasificacion_productos=clasificaciones.clasificacion_productos "
                + "left join (select p.clasificacion_productos,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and entysalVentas.cod_estab in (select cod_estab from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by p.clasificacion_productos) as acumulado on acumulado.clasificacion_productos=clasificaciones.clasificacion_productos "
                + "left join (select p.clasificacion_productos,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join productos as p on entysalVentas.cod_prod=p.cod_prod ) "
                + "where entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and entysalVentas.cod_estab in (select cod_estab from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by p.clasificacion_productos) as baseacumulado on baseacumulado.clasificacion_productos=clasificaciones.clasificacion_productos "
                + "where ISNULL(diario.unidades,0)<>0 or ISNULL(basediario.unidades,0)<>0 or ISNULL(diario.Total,0)<>0 or ISNULL(basediario.Total,0)<>0 or "
                + "ISNULL(diario.UtilBruta,0)<>0 or ISNULL(basediario.UtilBruta,0)<>0 or ISNULL(acumulado.unidades,0)<>0 or ISNULL(baseacumulado.unidades,0)<>0 or "
                + "ISNULL(acumulado.Total,0)<>0 or ISNULL(baseacumulado.Total,0)<>0 or ISNULL(acumulado.UtilBruta,0)<>0 or ISNULL(baseacumulado.UtilBruta,0)<>0 "
                + "order by isnull(acumulado.total,0)-isnull(baseacumulado.total,0) asc";
                #endregion
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                DG1.DataSource = dt;
                DG1.SelectAll();
                objeto = DG1.GetClipboardContent();
                if (objeto != null)
                {

                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Sheets.get_Item(5);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO POR PRODUCTO";
                    hoja.Cells[1, 2] = "COMPARATIVO POR PRODUCTO";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "U1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B2", "U2");
                    rango.Select();
                    rango.Merge();
                    rango.Cells[1.1, Type.Missing] = "Del " + primero.ToString("d MMM yyyy") + " al " + dia.ToString("d MMM yyyy");
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B3", "B5");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "PRODUCTO";
                    rango = (Range)hoja.get_Range("C3", "C5");
                    rango.Select();
                    rango.Merge();
                    //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "NOMBRE";
                    for (int i = 4; i <= 21; i += 9)
                    {
                        //rango = (Range)hoja.get_Range("4,4","12,4");
                        rango = (Range)hoja.get_Range(sCol(i) + "3", sCol(i + 8) + "3");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        switch (i)
                        {
                            case 4:
                                rango.Cells[1, 1] = "COMPARATIVO TOTAL DIARIO";
                                break;
                            case 13:
                                rango.Cells[1, 1] = "COMPARATIVO TOTAL ACUMULADO";
                                break;

                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            rango = (Range)hoja.get_Range(sCol(j) + "4", sCol(j + 1) + "4");
                            rango.Select();
                            rango.Merge();
                            rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            rango.Cells.Font.FontStyle = "Bold";
                            hoja.Cells[4, j + 2] = "%";
                            hoja.Cells[4, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            hoja.Cells[4, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            hoja.Cells[4, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[4, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            if (j == i)
                            {
                                rango.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                rango.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                rango.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = dia.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = diabase.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "Inc o Dec";
                                    hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00";
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 6) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 6) + ")-1)*100";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                hoja.Cells[5, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";

                            }
                        }
                    }
                    hoja.Range["B6", "C" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", "U" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    hoja.Range["B6", "U" + Convert.ToString(DG1.Rows.Count + 6)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    hoja.Cells[DG1.Rows.Count + 6, 3] = "T O T A L";
                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "U" + Convert.ToString(DG1.Rows.Count + 10));
                    rango.EntireColumn.AutoFit();
                }
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //
                //                                          COMPARATIVO DIARIO POR TIPO
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Clipboard.Clear();
                objeto = null;
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                #region ComparativoDiarioxTipo
                cmd.CommandText = "select Sucursales.cod_estab,Sucursales.nombre, "
                + "isnull(LineaDiario.unidades,0) as LineaDiarioUni,isnull(LineaDiarioBase.unidades,0) as LineaDiarioBaseUni,[Inc o Dec1]=case when LineaDiarioBase.unidades=0 then 0 when LineaDiarioBase.unidades is null then 0 when LineaDiarioBase.unidades>0 then ((LineaDiario.unidades/LineaDiarioBase.unidades)-1)*100 end,"
                + "isnull(LineaDiario.Total,0) as LineaDiarioVta,isnull(LineaDiarioBase.Total,0) as LineaDiarioBaseVta,[Inc o Dec2]=case	when LineaDiarioBase.Total=0 then 0 when LineaDiarioBase.Total is null then 0 when LineaDiarioBase.Total>0 then ((LineaDiario.Total/LineaDiarioBase.Total)-1)*100 end,"
                + "isnull(LineaDiario.UtilBruta,0) as LineaDiarioUtil,	isnull(LineaDiarioBase.UtilBruta,0) as LineaDiarioBaseUtil,[Inc o Dec3]=case when LineaDiarioBase.UtilBruta=0 then 0 when LineaDiarioBase.UtilBruta is null then 0 when LineaDiarioBase.UtilBruta>0 then ((LineaDiario.UtilBruta/LineaDiarioBase.UtilBruta)-1)*100 end,"
                + "isnull(ModaDiario.unidades,0) as ModaDiarioUni,isnull(ModaDiarioBase.unidades,0) as ModaDiarioBaseUni,[Inc o Dec4]=case when ModaDiarioBase.unidades=0 then 0 when ModaDiarioBase.unidades is null then 0 when ModaDiarioBase.unidades>0 then ((ModaDiario.unidades/ModaDiarioBase.unidades)-1)*100 end,"
                + "isnull(ModaDiario.Total,0) as ModaDiarioVta, isnull(ModaDiarioBase.Total,0) as ModaDiarioBaseVta,[Inc o Dec5]=case when ModaDiarioBase.Total=0 then 0 when ModaDiarioBase.Total is null then 0 when ModaDiarioBase.Total>0 then ((ModaDiario.Total/ModaDiarioBase.Total)-1)*100 end,"
                + "isnull(ModaDiario.UtilBruta,0) as ModaDiarioUtil,	isnull(ModaDiarioBase.UtilBruta,0) as ModaDiarioBaseUtil,[Inc o Dec6]=case when ModaDiarioBase.UtilBruta=0 then 0 when ModaDiarioBase.UtilBruta is null then 0 when ModaDiarioBase.UtilBruta>0 then ((ModaDiario.UtilBruta/ModaDiarioBase.UtilBruta)-1)*100 end,"
                + "isnull(TempDiario.unidades,0) as TempDiarioUni,isnull(TempDiarioBase.unidades,0) as TempBaseUni,[Inc o Dec7]=case when TempDiarioBase.unidades=0 then 0 when TempDiarioBase.unidades is null then 0 when TempDiarioBase.unidades>0 then ((TempDiario.unidades/TempDiarioBase.unidades)-1)*100 end,"
                + "isnull(TempDiario.Total,0) as TempDiarioVta,isnull(TempDiarioBase.Total,0) as TempBaseVta,[Inc o Dec8]=case when TempDiarioBase.Total=0 then 0 when TempDiarioBase.Total is null then 0 when TempDiarioBase.Total>0 then ((TempDiario.Total/TempDiarioBase.Total)-1)*100 end,"
                + "isnull(TempDiario.UtilBruta,0) as TempDiarioUtil,isnull(TempDiarioBase.UtilBruta,0) as TempBaseUtil,[Inc o Dec9]=case when TempDiarioBase.UtilBruta=0 then 0 when TempDiarioBase.UtilBruta is null then 0 when TempDiarioBase.UtilBruta>0 then ((TempDiario.UtilBruta/TempDiarioBase.UtilBruta)-1)*100 end "
                + "from (((((((select cod_estab,nombre,tipo_establecimiento,grupo_establecimiento from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) as Sucursales "
                + "left join "
                + "(select entysalventas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab) "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='LINEA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as LineaDiario on Sucursales.cod_estab=LineaDiario.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod  and entysalVentas.cod_estab=p.cod_estab "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='LINEA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as LineaDiarioBase on Sucursales.cod_estab=LineaDiarioBase.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab) "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='TEMP') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TempDiario on Sucursales.cod_estab=TempDiario.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='TEMP') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TempDiarioBase on Sucursales.cod_estab=TempDiarioBase.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab) "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='MODA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ModaDiario on Sucursales.cod_estab=ModaDiario.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='MODA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + diabase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ModaDiarioBase on Sucursales.cod_estab=ModaDiarioBase.cod_estab) "
                + "order  by sucursales.tipo_establecimiento desc,sucursales.grupo_establecimiento asc, CAST(Sucursales.cod_estab as int) asc";

                #endregion
                //cmd.ExecuteNonQuery();        
                //cmd.CommandText = "select ##TablaVentaCompAcumulado.* from ##TablaVentaCompAcumulado inner join establecimientos on ##TablaVentaCompAcumulado.cod_estab=establecimientos.cod_estab "
                //+ "order by establecimientos.tipo_establecimiento desc,establecimientos.grupo_establecimiento asc,cast(establecimientos.cod_estab as int)";            
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                DG1.DataSource = dt;
                DG1.SelectAll();
                objeto = DG1.GetClipboardContent();
                if (objeto != null)
                {

                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Sheets.get_Item(6);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO DIARIO X TIPO";
                    hoja.Cells[1, 2] = "COMPARATIVO DIARIO X TIPO";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "AD1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B2", "AD2");
                    rango.Select();
                    rango.Merge();
                    rango.Cells[1.1, Type.Missing] = "Del " + dia.ToString("d MMM yyyy");
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B3", "B5");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "CODIGO";
                    rango = (Range)hoja.get_Range("C3", "C5");
                    rango.Select();
                    rango.Merge();
                    //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "SUCURSAL";
                    for (int i = 4; i <= 30; i += 9)
                    {
                        //rango = (Range)hoja.get_Range("4,4","12,4");
                        rango = (Range)hoja.get_Range(sCol(i) + "3", sCol(i + 8) + "3");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        switch (i)
                        {
                            case 4:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO LINEA";
                                break;
                            case 13:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO MODA";
                                break;
                            case 22:
                                rango.Cells[1, 1] = "COMPARATIVO DIARIO TEMPORADA";
                                break;
                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            rango = (Range)hoja.get_Range(sCol(j) + "4", sCol(j + 1) + "4");
                            rango.Select();
                            rango.Merge();
                            rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            rango.Cells.Font.FontStyle = "Bold";
                            hoja.Cells[4, j + 2] = "%";
                            hoja.Cells[4, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            hoja.Cells[4, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            hoja.Cells[4, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[4, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            if (j == i)
                            {
                                rango.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                rango.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                rango.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = dia.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = diabase.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "Inc o Dec";
                                    hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00";
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 6) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 6) + ")-1)*100";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                hoja.Cells[5, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";

                            }
                        }
                    }
                    hoja.Range["B6", "C" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", "AD" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    hoja.Range["B6", "AD" + Convert.ToString(DG1.Rows.Count + 6)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    hoja.Cells[DG1.Rows.Count + 6, 3] = "T O T A L";
                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "AD" + Convert.ToString(DG1.Rows.Count + 10));
                    rango.EntireColumn.AutoFit();
                }
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //
                //                                          COMPARATIVO ACUMULADO POR TIPO
                //
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Clipboard.Clear();
                objeto = null;
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                #region ComparativoAcumuladoxTipo
                cmd.CommandText = "select Sucursales.cod_estab,Sucursales.nombre,"
                + "isnull(LineaAcu.unidades,0) as LineaAcuUni,isnull(LineaAcuBase.unidades,0) as LineaAcuBaseUni,[Inc o Dec1]=case when LineaAcuBase.unidades=0 then 0 when LineaAcuBase.unidades is null then 0 when LineaAcuBase.unidades>0 then ((LineaAcu.unidades/LineaAcuBase.unidades)-1)*100 end,"
                + "isnull(LineaAcu.Total,0) as LineaAcuVta,isnull(LineaAcuBase.Total,0) as LineaAcuBaseVta,[Inc o Dec2]=case	when LineaAcuBase.Total=0 then 0 when LineaAcuBase.Total is null then 0 when LineaAcuBase.Total>0 then ((LineaAcu.Total/LineaAcuBase.Total)-1)*100 end,"
                + "isnull(LineaAcu.UtilBruta,0) as LineaAcuUtil,	isnull(LineaAcuBase.UtilBruta,0) as LineaAcuBaseUtil,[Inc o Dec3]=case when LineaAcuBase.UtilBruta=0 then 0 when LineaAcuBase.UtilBruta is null then 0 when LineaAcuBase.UtilBruta>0 then ((LineaAcu.UtilBruta/LineaAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(ModaAcu.unidades,0) as ModaAcuUni,isnull(ModaAcuBase.unidades,0) as ModaAcuBaseUni,[Inc o Dec4]=case when ModaAcuBase.unidades=0 then 0 when ModaAcuBase.unidades is null then 0 when ModaAcuBase.unidades>0 then ((ModaAcu.unidades/ModaAcuBase.unidades)-1)*100 end,"
                + "isnull(ModaAcu.Total,0) as ModaAcuVta, isnull(ModaAcuBase.Total,0) as ModaAcuBaseVta,[Inc o Dec5]=case when ModaAcuBase.Total=0 then 0 when ModaAcuBase.Total is null then 0 when ModaAcuBase.Total>0 then ((ModaAcu.Total/ModaAcuBase.Total)-1)*100 end,"
                + "isnull(ModaAcu.UtilBruta,0) as ModaAcuUtil,	isnull(ModaAcuBase.UtilBruta,0) as ModaAcuBaseUtil,[Inc o Dec6]=case when ModaAcuBase.UtilBruta=0 then 0 when ModaAcuBase.UtilBruta is null then 0 when ModaAcuBase.UtilBruta>0 then ((ModaAcu.UtilBruta/ModaAcuBase.UtilBruta)-1)*100 end,"
                + "isnull(TempAcu.unidades,0) as TempDiarioUni,isnull(TempAcuBase.unidades,0) as TempBaseUni,[Inc o Dec7]=case when TempAcuBase.unidades=0 then 0 when TempAcuBase.unidades is null then 0 when TempAcuBase.unidades>0 then ((TempAcu.unidades/TempAcuBase.unidades)-1)*100 end,"
                + "isnull(TempAcu.Total,0) as TempDiarioVta,isnull(TempAcuBase.Total,0) as TempBaseVta,[Inc o Dec8]=case when TempAcuBase.Total=0 then 0 when TempAcuBase.Total is null then 0 when TempAcuBase.Total>0 then ((TempAcu.Total/TempAcuBase.Total)-1)*100 end,"
                + "isnull(TempAcu.UtilBruta,0) as TempDiarioUtil,isnull(TempAcuBase.UtilBruta,0) as TempBaseUtil,[Inc o Dec9]=case when TempAcuBase.UtilBruta=0 then 0 when TempAcuBase.UtilBruta is null then 0 when TempAcuBase.UtilBruta>0 then ((TempAcu.UtilBruta/TempAcuBase.UtilBruta)-1)*100 end "
                + "from (((((((select cod_estab,nombre,tipo_establecimiento,grupo_establecimiento from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65')) as Sucursales "
                + "left join "
                + "(select entysalventas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab) "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='LINEA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as LineaAcu on Sucursales.cod_estab=LineaAcu.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod  and entysalVentas.cod_estab=p.cod_estab "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='LINEA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as LineaAcuBase on Sucursales.cod_estab=LineaAcuBase.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab) "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='TEMP') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TempAcu on Sucursales.cod_estab=TempAcu.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='TEMP') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as TempAcuBase on Sucursales.cod_estab=TempAcuBase.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from (entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab) "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='MODA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primero.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ModaAcu on Sucursales.cod_estab=ModaAcu.cod_estab) "
                + "left join "
                + "(select entysalVentas.cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta "
                + "from entysalVentas with(nolock) inner join prodestab as p on entysalVentas.cod_prod=p.cod_prod and entysalVentas.cod_estab=p.cod_estab "
                + "where p.tipo_producto in (select tipo_producto from tipos_productos where abreviatura='MODA') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68')and fecha between '" + primerobase.ToString("yyyyMMdd") + "' and '" + diabase.ToString("yyyyMMdd HH:mm") + "' group by entysalVentas.cod_estab) as ModaAcuBase on Sucursales.cod_estab=ModaAcuBase.cod_estab) "
                + "order  by sucursales.tipo_establecimiento desc,sucursales.grupo_establecimiento asc, CAST(Sucursales.cod_estab as int) asc;";
                #endregion
                //cmd.ExecuteNonQuery();        
                //cmd.CommandText = "select ##TablaVentaCompAcumulado.* from ##TablaVentaCompAcumulado inner join establecimientos on ##TablaVentaCompAcumulado.cod_estab=establecimientos.cod_estab "
                //+ "order by establecimientos.tipo_establecimiento desc,establecimientos.grupo_establecimiento asc,cast(establecimientos.cod_estab as int)";            
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                DG1.DataSource = dt;
                DG1.SelectAll();
                objeto = DG1.GetClipboardContent();
                if (objeto != null)
                {

                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Sheets.get_Item(7);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO ACUMULADO X TIPO";
                    hoja.Cells[1, 2] = "COMPARATIVO ACUMULADO X TIPO";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "AD1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B2", "AD2");
                    rango.Select();
                    rango.Merge();
                    rango.Cells[1.1, Type.Missing] = "Del " + primero.ToString("d MMM yyyy") + " al " + dia.ToString("d MMM yyyy");
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango = (Range)hoja.get_Range("B3", "B5");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "CODIGO";
                    rango = (Range)hoja.get_Range("C3", "C5");
                    rango.Select();
                    rango.Merge();
                    //rango.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    rango.Cells[1, 1] = "SUCURSAL";
                    for (int i = 4; i <= 30; i += 9)
                    {
                        //rango = (Range)hoja.get_Range("4,4","12,4");
                        rango = (Range)hoja.get_Range(sCol(i) + "3", sCol(i + 8) + "3");
                        rango.Select();
                        rango.Merge();
                        rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        rango.Cells.Font.FontStyle = "Bold";
                        switch (i)
                        {
                            case 4:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO LINEA";
                                break;
                            case 13:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO MODA";
                                break;
                            case 22:
                                rango.Cells[1, 1] = "COMPARATIVO ACUMULADO TEMPORADA";
                                break;
                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            rango = (Range)hoja.get_Range(sCol(j) + "4", sCol(j + 1) + "4");
                            rango.Select();
                            rango.Merge();
                            rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            rango.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            rango.Cells.Font.FontStyle = "Bold";
                            hoja.Cells[4, j + 2] = "%";
                            hoja.Cells[4, j + 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            hoja.Cells[4, j + 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                            hoja.Cells[4, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            hoja.Cells[4, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                            if (j == i)
                            {
                                rango.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                rango.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                rango.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = dia.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                    hoja.Cells[5, k].NumberFormat = "@";
                                    hoja.Cells[5, k] = diabase.Year.ToString();
                                    if (j == i)
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0"; }
                                    else
                                    { hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00"; }
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=SUM(" + sCol(k) + "6:" + sCol(k) + Convert.ToString(DG1.Rows.Count + 5) + ")";
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "Inc o Dec";
                                    hoja.Range[sCol(k) + "6", sCol(k) + Convert.ToString(DG1.Rows.Count + 6)].NumberFormat = "#,###,##0.00";
                                    hoja.Cells[DG1.Rows.Count + 6, k].Formula = "=((" + sCol(k - 2) + Convert.ToString(DG1.Rows.Count + 6) + "/" + sCol(k - 1) + Convert.ToString(DG1.Rows.Count + 6) + ")-1)*100";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                hoja.Cells[5, k].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";

                            }
                        }
                    }
                    hoja.Range["B6", "C" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", "AD" + Convert.ToString(DG1.Rows.Count + 5)].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    hoja.Range["B6", "AD" + Convert.ToString(DG1.Rows.Count + 6)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    hoja.Cells[DG1.Rows.Count + 6, 3] = "T O T A L";
                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "AD" + Convert.ToString(DG1.Rows.Count + 10));
                    rango.EntireColumn.AutoFit();
                }

                hoja = (Worksheet)libro.Sheets.get_Item(1);
                hoja.Activate();
                if (cn.State.ToString() == "Open") { cn.Close(); }
                if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas Diarias al " + dia.ToString("dd MMMM yyyy") + ".xlsx"))
                {
                    File.Delete(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas Diarias al " + dia.ToString("dd MMMM yyyy") + ".xlsx");
                }

                libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas Diarias al " + dia.ToString("dd MMMM yyyy") + ".xlsx");
                libro.Close();
                excel.Quit();
                return System.Windows.Forms.Application.StartupPath + "\\Comparativo de Ventas Diarias al " + dia.ToString("dd MMMM yyyy") + ".xlsx";
                //string archivo=Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+"\\Comparativo de Ventas Diarias al " + dia.ToString("dd MMM yyyy") + ".xlsx";
                //if (File.Exists(archivo))
                //{
                //    File.Delete(archivo);
                //}
                //libro.SaveAs(archivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                //libro.Close();
                //excel.Quit();
                //return archivo;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "";
            }
        }

        private string ReporteApertura(DateTime dia)
        {
            try
            {
                SqlConnection cn = conexion.conectar("BMSNayar"); // 06/Nov/2018 -> Este reporte estaba apuntando a la BD de Tamazula
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandTimeout = 180;
                System.Data.DataTable dt = new System.Data.DataTable();
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                dia = dia.AddDays(-1);

                cmd.CommandText = "select case when salidas.cod_estab is null then entradas.cod_estab else salidas.cod_estab end as cod_estab,"
                + " case	when salidas.nombre is null then entradas.nombre else salidas.nombre end as establecimiento, case when salidas.dia is null then entradas.dia else salidas.dia end as dia,"
                + " entradas.entrada as primer_checada,salidas.salida as ultima_checada,(select MIN(fecha) from BDIntegrador..entysal where fecha between entradas.entrada and DATEADD(day,1,cast(floor(cast(salidas.salida as float)) as smalldatetime)) and entysal.cod_estab=entradas.cod_estab) as primer_Mov,"
                + " (select MAX(fecha) from BDIntegrador..entysal where fecha between entradas.entrada and DATEADD(day,1,cast(floor(cast(salidas.salida as float)) as smalldatetime)) and entysal.cod_estab=entradas.cod_estab) as ultimo_Mov"
                + " from (Select e.cod_estab,estabs.nombre,datepart(DAY,fecha) as dia,min(fecha) as entrada  from BMSNayar..entradas_salidas_empleados  as es inner join BMSNayar..empleados as e on es.empleado=e.empleado inner join establecimientos as estabs on e.cod_estab=estabs.cod_estab "
                + " where fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + " 23:59' and e.cod_estab not in ('1001','1002','1003','1004') group by e.cod_estab, datepart(DAY,fecha),estabs.nombre) as entradas full join "
                + " (Select e.cod_estab,estabs.nombre,datepart(DAY,fecha) as dia,max(fecha) as salida  from BMSNayar..entradas_salidas_empleados  as es inner join BMSNayar..empleados as e on es.empleado=e.empleado inner join establecimientos as estabs on e.cod_estab=estabs.cod_estab  "
                + " where fecha between '" + dia.ToString("yyyyMMdd") + "' and '" + dia.ToString("yyyyMMdd") + " 23:59' and e.cod_estab not in ('1001','1002','1003','1004') group by e.cod_estab, datepart(DAY,fecha),estabs.nombre) as salidas"
                + " on entradas.cod_estab=salidas.cod_estab and entradas.dia=salidas.dia order by cast(entradas.cod_estab as int),entradas.entrada";

                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                da.Fill(dt);
                DG1.DataSource = dt;
                DG1.SelectAll();
                object objeto = DG1.GetClipboardContent();
                Microsoft.Office.Interop.Excel.Application excel;
                excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook libro;
                libro = excel.Workbooks.Add();
                libro.Worksheets.Add();
                Worksheet hoja = new Worksheet();
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                hoja.Name = "APERTURA Y CIERRE";
                Microsoft.Office.Interop.Excel.Range rango;
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja.Cells[1, 2] = "APERTURA Y CIERRE";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "H1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Cells.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    hoja.Cells[2, 2] = "COD_ESTAB";
                    hoja.Cells[2, 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 2].Font.FontStyle = "Bold";

                    hoja.Cells[2, 3] = "ESTABLECIMIENTO";
                    hoja.Cells[2, 3].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 3].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 3].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 3].Font.FontStyle = "Bold";

                    hoja.Cells[2, 4] = "DIA";
                    hoja.Cells[2, 4].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 4].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 4].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 4].Font.FontStyle = "Bold";

                    hoja.Cells[2, 5] = "PRIMER CHECADA";
                    hoja.Cells[2, 5].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 5].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 5].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 5].Font.FontStyle = "Bold";

                    hoja.Cells[2, 6] = "ULTIMA CHECADA";
                    hoja.Cells[2, 6].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 6].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 6].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 6].Font.FontStyle = "Bold";

                    hoja.Cells[2, 7] = "PRIMER MOVIMIENTO";
                    hoja.Cells[2, 7].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 7].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 7].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 7].Font.FontStyle = "Bold";

                    hoja.Cells[2, 8] = "ULTIMO MOVIMIENTO";
                    hoja.Cells[2, 8].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 8].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 8].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 8].Font.FontStyle = "Bold";

                    hoja.Range["B3", "H" + Convert.ToString(DG1.Rows.Count + 1)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    rango = (Range)hoja.Cells[3, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "H" + Convert.ToString(DG1.Rows.Count + 1));
                    rango.EntireColumn.AutoFit();

                    if (cn.State.ToString() == "Open") { cn.Close(); }
                    if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Informe de Apertura y Cierre al " + dia.ToString("dd MMMM yyyy") + ".xlsx"))
                    {
                        File.Delete(System.Windows.Forms.Application.StartupPath + "\\Informe de Apertura y Cierre al " + dia.ToString("dd MMMM yyyy") + ".xlsx");
                    }

                    libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Informe de Apertura y Cierre al " + dia.ToString("dd MMMM yyyy") + ".xlsx");
                    libro.Close();
                    excel.Quit();

                    return System.Windows.Forms.Application.StartupPath + "\\Informe de Apertura y Cierre al " + dia.ToString("dd MMMM yyyy") + ".xlsx";
                }

                return "";
            }
            catch(Exception e)
            {
                lblEstado.Text = "Reporte no Generado: " + e.Message.ToString();
                return "";
            }
        }

        private string ReporteDescuento(DateTime dia)
        {
            try
            {
                SqlConnection cn = conexion.conectar("BMSNayar");
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandTimeout = 180;
                System.Data.DataTable dt = new System.Data.DataTable();
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                dia = dia.AddDays(-1);
                lblEstado.Text = "Enviando Reporte de Descuentos [EN PROCESO]...";

                cmd.CommandText = "Select  ventas.cod_estab,ventas.folio,transacciones.abreviatura as trans,ventas.cod_prod,ltrim(rtrim(productos.descripcion)) as descripcion,"
                + " ltrim(rtrim(lineas_productos.nombre)) as linea,ltrim(rtrim(familias.nombre)) as familia,clasificaciones_productos.nombre as clasif,	ltrim(rtrim(tipos_productos.nombre)) as [tipo producto],"
                + " round(dbo.precio_venta_rpt(ventas.cod_prod,'U',establecimientos.lista_precios,getdate(),'1',0,0,0)*1.16,2) as [precio normal],round(Ventas.precio_movto,2) as [precio movto],"
                + " ventas.cantidad,ventas.Importe,ventas.descuento,(ventas.descuento/ventas.Importe)*100 as [%Descto],ventas.VentaNeta,facremtick.cajero,cajeros.nombre,ventas.folios_cv,ventas.descuentos,"
                + " (select top 1 tipo_bonificacion from condiciones_venta where folio=replace(ventas.folios_cv,',','')) as TipoDesc"
                + " from (select entysal.cod_estab,entysal.folio,entysal.transaccion,entysal.cod_prod,SUM(case when entysal.importe<0 then case entysal.iva when 0 then (entysal.importe_descuento*-1)+entysal.importe else ((entysal.importe_descuento*((entysal.iva/entysal.importe)+1)))+(entysal.importe+entysal.iva) end"
                + " else case entysal.iva when 0 then entysal.importe_descuento+entysal.importe else (entysal.importe_descuento*((entysal.iva/entysal.importe)+1))+(entysal.importe+entysal.iva) end end)/SUM(entysal.cantidad) as precio_movto,"
                + " SUM(entysal.cantidad) as cantidad,SUM(case when entysal.importe<0 then case entysal.iva when 0 then (entysal.importe_descuento*-1)+entysal.importe else ((entysal.importe_descuento*((entysal.iva/entysal.importe)+1)))+(entysal.importe+entysal.iva) end "
                + " else case entysal.iva when 0 then entysal.importe_descuento+entysal.importe else (entysal.importe_descuento*((entysal.iva/entysal.importe)+1))+(entysal.importe+entysal.iva) end end) as Importe,"
                + " SUM(case when entysal.importe<0 then case entysal.iva when 0 then entysal.importe_descuento	else (entysal.importe_descuento*((entysal.iva/entysal.importe)+1)) end else"
                + "	case entysal.iva when 0 then entysal.importe_descuento else entysal.importe_descuento*((entysal.iva/entysal.importe)+1) end	end) as descuento,SUM(entysal.total) as VentaNeta,"
                + " SUM(entysal.cantidad*entysal.precio_lista) as ImporteSinIva,SUM(entysal.importe) as VentaNetaSinIva,SUM(entysal.costo) as costo,cv_aplicadas.folios_cv,cv_aplicadas.descuentos"
                + " from dbo.entysal with(nolock) inner join dbo.clientes with(nolock) on entysal.cod_cte=clientes.cod_cte inner join dbo.productos with(nolock) on entysal.cod_prod=productos.cod_prod inner join facremtick with(nolock) on facremtick.folio=entysal.folio and facremtick.transaccion=entysal.transaccion"
                + " left join cv_aplicadas on entysal.folio=cv_aplicadas.folio and entysal.transaccion=cv_aplicadas.transaccion and cv_aplicadas.cod_prod=entysal.cod_prod"
                + " where entysal.transaccion in ('36','37','38') and clientes.clasificacion_cliente<>'2' and productos.tipo_producto<>'7' and entysal.tipo_precio_venta<>'MS'"
                + " and facremtick.status='V' and cast(floor(cast(entysal.fecha as float)) as smalldatetime)='" + dia.ToString("yyyyMMdd") + "'"
                + " group by entysal.cod_estab,entysal.cod_prod,entysal.folio,entysal.transaccion,cv_aplicadas.folios_cv,cv_aplicadas.descuentos) as ventas"
                + " inner join productos with(nolock) on ventas.cod_prod=productos.cod_prod inner join facremtick with(nolock) on ventas.folio=facremtick.folio and ventas.transaccion=facremtick.transaccion"
                + " left join lineas_productos with(nolock) on productos.linea_producto=lineas_productos.linea_producto left join familias with(nolock) on productos.familia=familias.familia"
                + " left join clasificaciones_productos with(nolock) on clasificaciones_productos.clasificacion_productos=productos.clasificacion_productos left join tipos_productos with(nolock) on tipos_productos.tipo_producto=productos.tipo_producto"
                + " left join cajeros with(nolock) on facremtick.cajero=cajeros.cajero left join transacciones with(nolock) on facremtick.transaccion=transacciones.transaccion left join establecimientos with(nolock) on establecimientos.cod_estab=ventas.cod_estab"
                + " where productos.tipo_producto in ('1','15','16','2','3','4','5','') and ((productos.linea_producto in ('1','2','3','4') and ventas.descuento/ventas.Importe>.2) or (productos.linea_producto='8' and ventas.descuento/ventas.Importe>.1))"
                + " Order by cast(ventas.cod_estab as int),descuento desc";

                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                da.Fill(dt);
                DG1.DataSource = dt;
                DG1.SelectAll();
                object objeto = DG1.GetClipboardContent();
                Microsoft.Office.Interop.Excel.Application excel;
                excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook libro;
                libro = excel.Workbooks.Add();
                libro.Worksheets.Add();
                Worksheet hoja = new Worksheet();
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                hoja.Name = "NO TEMPORADA";
                Microsoft.Office.Interop.Excel.Range rango;
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja.Cells[1, 2] = "VENTAS CON DESCUENTO MAYOR AL PERMITIDO";
                    rango = (Range)hoja.get_Range("B1", "V1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Cells.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    hoja.Cells[2, 2] = "COD_ESTAB";
                    hoja.Cells[2, 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 2].Font.FontStyle = "Bold";

                    hoja.Cells[2, 3] = "FOLIO";
                    hoja.Cells[2, 3].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 3].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 3].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 3].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("C1", "C1");
                    rango.EntireColumn.NumberFormat = "@";

                    hoja.Cells[2, 4] = "TRANS";
                    hoja.Cells[2, 4].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 4].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 4].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 4].Font.FontStyle = "Bold";

                    hoja.Cells[2, 5] = "COD PROD";
                    hoja.Cells[2, 5].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 5].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 5].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 5].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("E1", "E1");
                    rango.EntireColumn.NumberFormat = "@";

                    hoja.Cells[2, 6] = "DESCRIPCION";
                    hoja.Cells[2, 6].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 6].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 6].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 6].Font.FontStyle = "Bold";

                    hoja.Cells[2, 7] = "LINEA";
                    hoja.Cells[2, 7].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 7].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 7].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 7].Font.FontStyle = "Bold";

                    hoja.Cells[2, 8] = "FAMILIA";
                    hoja.Cells[2, 8].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 8].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 8].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 8].Font.FontStyle = "Bold";

                    hoja.Cells[2, 9] = "CLASIFICACION";
                    hoja.Cells[2, 9].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 9].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 9].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 9].Font.FontStyle = "Bold";

                    hoja.Cells[2, 10] = "TIPO PRODUCTO";
                    hoja.Cells[2, 10].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 10].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 10].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 10].Font.FontStyle = "Bold";

                    hoja.Cells[2, 11] = "PRECIO NORMAL";
                    hoja.Cells[2, 11].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 11].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 11].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 11].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("K1", "K1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 12] = "PRECIO MOVTO";
                    hoja.Cells[2, 12].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 12].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 12].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 12].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("L1", "L1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 13] = "CANTIDAD";
                    hoja.Cells[2, 13].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 13].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 13].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 13].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("M1", "M1");
                    rango.EntireColumn.NumberFormat = "###,##0";

                    hoja.Cells[2, 14] = "IMPORTE";
                    hoja.Cells[2, 14].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 14].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 14].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 14].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("N1", "N1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 15] = "DESCUENTO";
                    hoja.Cells[2, 15].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 15].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 15].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 15].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("O1", "O1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 16] = "%DESCTO";
                    hoja.Cells[2, 16].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 16].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 16].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 16].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("P1", "P1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 17] = "VENTA NETA";
                    hoja.Cells[2, 17].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 17].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 17].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 17].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("Q1", "Q1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 18] = "CAJERO";
                    hoja.Cells[2, 18].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 18].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 18].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 18].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("R1", "R1");
                    rango.EntireColumn.NumberFormat = "@";

                    hoja.Cells[2, 19] = "NOMBRE";
                    hoja.Cells[2, 19].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 19].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 19].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 19].Font.FontStyle = "Bold";

                    hoja.Cells[2, 20] = "FOLIOS CV";
                    hoja.Cells[2, 20].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 20].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 20].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 20].Font.FontStyle = "Bold";

                    hoja.Cells[2, 21] = "DESCTO O PRECIO";
                    hoja.Cells[2, 21].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 21].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 21].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 21].Font.FontStyle = "Bold";

                    hoja.Cells[2, 22] = "TIPO DESCTO";
                    hoja.Cells[2, 22].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 22].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 22].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 22].Font.FontStyle = "Bold";

                    hoja.Range["B3", "V" + Convert.ToString(DG1.Rows.Count + 1)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    rango = (Range)hoja.Cells[3, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "V" + Convert.ToString(DG1.Rows.Count + 1));
                    rango.EntireColumn.AutoFit();
                }
                Clipboard.Clear();
                objeto = null;
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                cmd.CommandText = "Select  ventas.cod_estab,ventas.folio,transacciones.abreviatura as trans,ventas.cod_prod,ltrim(rtrim(productos.descripcion)) as descripcion,"
                + " ltrim(rtrim(lineas_productos.nombre)) as linea,ltrim(rtrim(familias.nombre)) as familia,clasificaciones_productos.nombre as clasif,	ltrim(rtrim(tipos_productos.nombre)) as [tipo producto],"
                + " round(dbo.precio_venta_rpt(ventas.cod_prod,'U',establecimientos.lista_precios,getdate(),'1',0,0,0)*1.16,2) as [precio normal],round(Ventas.precio_movto,2) as [precio movto],"
                + " ventas.cantidad,ventas.Importe,ventas.descuento,(ventas.descuento/ventas.Importe)*100 as [%Descto],ventas.VentaNeta,facremtick.cajero,cajeros.nombre,ventas.folios_cv,ventas.descuentos,"
                + " (select top 1 tipo_bonificacion from condiciones_venta where folio=replace(ventas.folios_cv,',','')) as TipoDesc"
                + " from (select entysal.cod_estab,entysal.folio,entysal.transaccion,entysal.cod_prod,SUM(case when entysal.importe<0 then case entysal.iva when 0 then (entysal.importe_descuento*-1)+entysal.importe else ((entysal.importe_descuento*((entysal.iva/entysal.importe)+1)))+(entysal.importe+entysal.iva) end"
                + " else case entysal.iva when 0 then entysal.importe_descuento+entysal.importe else (entysal.importe_descuento*((entysal.iva/entysal.importe)+1))+(entysal.importe+entysal.iva) end end)/SUM(entysal.cantidad) as precio_movto,"
                + " SUM(entysal.cantidad) as cantidad,SUM(case when entysal.importe<0 then case entysal.iva when 0 then (entysal.importe_descuento*-1)+entysal.importe else ((entysal.importe_descuento*((entysal.iva/entysal.importe)+1)))+(entysal.importe+entysal.iva) end "
                + " else case entysal.iva when 0 then entysal.importe_descuento+entysal.importe else (entysal.importe_descuento*((entysal.iva/entysal.importe)+1))+(entysal.importe+entysal.iva) end end) as Importe,"
                + " SUM(case when entysal.importe<0 then case entysal.iva when 0 then entysal.importe_descuento	else (entysal.importe_descuento*((entysal.iva/entysal.importe)+1)) end else"
                + "	case entysal.iva when 0 then entysal.importe_descuento else entysal.importe_descuento*((entysal.iva/entysal.importe)+1) end	end) as descuento,SUM(entysal.total) as VentaNeta,"
                + " SUM(entysal.cantidad*entysal.precio_lista) as ImporteSinIva,SUM(entysal.importe) as VentaNetaSinIva,SUM(entysal.costo) as costo,cv_aplicadas.folios_cv,cv_aplicadas.descuentos"
                + " from dbo.entysal with(nolock) inner join dbo.clientes with(nolock) on entysal.cod_cte=clientes.cod_cte inner join dbo.productos with(nolock) on entysal.cod_prod=productos.cod_prod inner join facremtick with(nolock) on facremtick.folio=entysal.folio and facremtick.transaccion=entysal.transaccion"
                + " left join cv_aplicadas on entysal.folio=cv_aplicadas.folio and entysal.transaccion=cv_aplicadas.transaccion and cv_aplicadas.cod_prod=entysal.cod_prod"
                + " where entysal.transaccion in ('36','37','38') and clientes.clasificacion_cliente<>'2' and productos.tipo_producto<>'7' and entysal.tipo_precio_venta<>'MS'"
                + " and facremtick.status='V' and cast(floor(cast(entysal.fecha as float)) as smalldatetime)='" + dia.ToString("yyyyMMdd") + "'"
                + " group by entysal.cod_estab,entysal.cod_prod,entysal.folio,entysal.transaccion,cv_aplicadas.folios_cv,cv_aplicadas.descuentos) as ventas"
                + " inner join productos with(nolock) on ventas.cod_prod=productos.cod_prod inner join facremtick with(nolock) on ventas.folio=facremtick.folio and ventas.transaccion=facremtick.transaccion"
                + " left join lineas_productos with(nolock) on productos.linea_producto=lineas_productos.linea_producto left join familias with(nolock) on productos.familia=familias.familia"
                + " left join clasificaciones_productos with(nolock) on clasificaciones_productos.clasificacion_productos=productos.clasificacion_productos left join tipos_productos with(nolock) on tipos_productos.tipo_producto=productos.tipo_producto"
                + " left join cajeros with(nolock) on facremtick.cajero=cajeros.cajero left join transacciones with(nolock) on facremtick.transaccion=transacciones.transaccion left join establecimientos with(nolock) on establecimientos.cod_estab=ventas.cod_estab"
                + " where productos.tipo_producto not in ('1','15','16','2','3','4','5','') and ventas.descuento/ventas.Importe>.5"
                + " Order by cast(ventas.cod_estab as int),descuento desc";

                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                DG1.DataSource = dt;
                DG1.SelectAll();
                objeto = DG1.GetClipboardContent();
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    libro.Worksheets.Add();
                    hoja = (Worksheet)libro.Sheets.get_Item(1);
                    hoja.Activate();
                    hoja.Name = "TEMPORADA";
                    hoja.Cells[1, 2] = "VENTAS CON DESCUENTO MAYOR AL PERMITIDO";
                    //ENCABEZADO VENTA DIARIA
                    rango = (Range)hoja.get_Range("B1", "V1");
                    rango.Select();
                    rango.Merge();
                    rango.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango.Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Cells.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Cells.Font.FontStyle = "Bold";
                    hoja.Cells[2, 2] = "COD_ESTAB";
                    hoja.Cells[2, 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 2].Font.FontStyle = "Bold";

                    hoja.Cells[2, 3] = "FOLIO";
                    hoja.Cells[2, 3].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 3].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 3].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 3].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("C1", "C1");
                    rango.EntireColumn.NumberFormat = "@";

                    hoja.Cells[2, 4] = "TRANS";
                    hoja.Cells[2, 4].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 4].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 4].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 4].Font.FontStyle = "Bold";

                    hoja.Cells[2, 5] = "COD PROD";
                    hoja.Cells[2, 5].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 5].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 5].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 5].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("E1", "E1");
                    rango.EntireColumn.NumberFormat = "@";

                    hoja.Cells[2, 6] = "DESCRIPCION";
                    hoja.Cells[2, 6].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 6].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 6].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 6].Font.FontStyle = "Bold";

                    hoja.Cells[2, 7] = "LINEA";
                    hoja.Cells[2, 7].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 7].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 7].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 7].Font.FontStyle = "Bold";

                    hoja.Cells[2, 8] = "FAMILIA";
                    hoja.Cells[2, 8].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 8].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 8].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 8].Font.FontStyle = "Bold";

                    hoja.Cells[2, 9] = "CLASIFICACION";
                    hoja.Cells[2, 9].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 9].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 9].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 9].Font.FontStyle = "Bold";

                    hoja.Cells[2, 10] = "TIPO PRODUCTO";
                    hoja.Cells[2, 10].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 10].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 10].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 10].Font.FontStyle = "Bold";

                    hoja.Cells[2, 11] = "PRECIO NORMAL";
                    hoja.Cells[2, 11].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 11].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 11].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 11].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("K1", "K1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 12] = "PRECIO MOVTO";
                    hoja.Cells[2, 12].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 12].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 12].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 12].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("L1", "L1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 13] = "CANTIDAD";
                    hoja.Cells[2, 13].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 13].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 13].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 13].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("M1", "M1");
                    rango.EntireColumn.NumberFormat = "###,##0";

                    hoja.Cells[2, 14] = "IMPORTE";
                    hoja.Cells[2, 14].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 14].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 14].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 14].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("N1", "N1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 15] = "DESCUENTO";
                    hoja.Cells[2, 15].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 15].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 15].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 15].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("O1", "O1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 16] = "%DESCTO";
                    hoja.Cells[2, 16].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 16].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 16].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 16].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("P1", "P1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 17] = "VENTA NETA";
                    hoja.Cells[2, 17].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 17].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 17].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 17].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("Q1", "Q1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";

                    hoja.Cells[2, 18] = "CAJERO";
                    hoja.Cells[2, 18].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 18].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 18].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 18].Font.FontStyle = "Bold";
                    rango = (Range)hoja.get_Range("R1", "R1");
                    rango.EntireColumn.NumberFormat = "@";

                    hoja.Cells[2, 19] = "NOMBRE";
                    hoja.Cells[2, 19].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 19].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 19].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 19].Font.FontStyle = "Bold";

                    hoja.Cells[2, 20] = "FOLIOS CV";
                    hoja.Cells[2, 20].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 20].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 20].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 20].Font.FontStyle = "Bold";

                    hoja.Cells[2, 21] = "DESCTO O PRECIO";
                    hoja.Cells[2, 21].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 21].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 21].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 21].Font.FontStyle = "Bold";

                    hoja.Cells[2, 22] = "TIPO DESCTO";
                    hoja.Cells[2, 22].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    hoja.Cells[2, 22].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[2, 22].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[2, 22].Font.FontStyle = "Bold";

                    hoja.Range["B3", "V" + Convert.ToString(DG1.Rows.Count + 1)].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    rango = (Range)hoja.Cells[3, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    rango = (Range)hoja.get_Range("A1", "V" + Convert.ToString(DG1.Rows.Count + 1));
                    rango.EntireColumn.AutoFit();

                }


                if (cn.State.ToString() == "Open") { cn.Close(); }
                if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Ventas con descuento del " + dia.ToString("dd MMMM yyyy") + ".xlsb"))
                {
                    File.Delete(System.Windows.Forms.Application.StartupPath + "\\Ventas con descuento del " + dia.ToString("dd MMMM yyyy") + ".xlsb");
                }

                libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Ventas con descuento del " + dia.ToString("dd MMMM yyyy") + ".xlsb", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12);
                libro.Close();
                excel.Quit();

                lblEstado.Text = "Reporte de Descuentos  [FINALIZADO]";
                return System.Windows.Forms.Application.StartupPath + "\\Ventas con descuento del " + dia.ToString("dd MMMM yyyy") + ".xlsb";
            }
            catch (Exception e)
            {
                lblEstado.Text = "Reporte no Generado: " + e.Message.ToString();
                return "";
            }
        }

        private void Incidencias()
        {
            SqlConnection cn = conexion.conectar("BMSNayar");
            SqlDataReader dr = (new SqlCommand("select establecimientos.cod_estab,establecimientos.nombre,mail.email,mail.email_cordinador  from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab=mail.cod_estab where establecimientos.status='V' and establecimientos.cod_estab not in ('1') order by cast(establecimientos.cod_estab as int) asc", cn)).ExecuteReader();
            if (dr.HasRows)
            {
                DateTime dia = DateTime.Now;
                dia = dia.AddDays(-1);
                dia = dia.AddHours((double)(dia.Hour * -1));
                dia = dia.AddMinutes((double)(dia.Minute * -1));
                while (dr.Read())
                {
                    string str = string.Concat(dr["email"].ToString().Trim(), ",", dr["email_cordinador"].ToString().Trim());
                    // Se agregan algunos destinatarios a propósito por cuestiones especiales
                    if (dr["cod_estab"].ToString().Trim() == "1003")
                    {
                        str = string.Concat(str, ",luis.guerrero@mercadodeimportaciones.com");
                    }

                    this.EnviaMailGmail(this.IncidenciasSucursal(dr["cod_estab"].ToString().Trim(), dr["nombre"].ToString(), dia), str);
                }
            }
        }

        private string IncidenciasSucursal(string cod_estab, string estab, DateTime FechaHora)
        {
            DateTime dia = FechaHora;
            DateTime primero = FechaHora;
            while (primero.DayOfWeek != DayOfWeek.Thursday)
            {
                primero = primero.AddDays(-1);
            }
            SqlConnection cn = conexion.conectar("BDReloj");
            SqlCommand sqlCommand = new SqlCommand()
            {
                Connection = cn,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Incidencias",
                CommandTimeout = 240
            };
            Microsoft.Office.Interop.Excel.Application excel;
            excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Application.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook libro;
            libro = excel.Workbooks.Add();
            Worksheet hoja = new Worksheet();
            sqlCommand.Parameters.Clear();
            sqlCommand.Parameters.AddWithValue("@cod_estab", cod_estab);
            sqlCommand.Parameters.AddWithValue("@fini", primero);
            sqlCommand.Parameters.AddWithValue("@ffin", dia);
            libro.Worksheets.Add();
            hoja = (Worksheet)libro.Worksheets.get_Item(1);
            hoja.Name = "INCIDENCIAS";
            SqlDataAdapter da = new SqlDataAdapter(sqlCommand);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt;
            this.DG1.SelectAll();
            object objeto = this.DG1.GetClipboardContent();
            if (objeto != null)
            {
                Clipboard.SetDataObject(objeto);
                foreach (DataGridViewColumn column in this.DG1.Columns)
                {
                    hoja.Cells[1, column.Index + 2] = column.Name.ToString().ToUpper();
                }
                Range rango = (Range)hoja.Cells[1, 3];
                rango.EntireColumn.NumberFormat = "@";
                rango = (Range)hoja.Cells[2, 1];
                rango.Select();
                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            Clipboard.Clear();
            objeto = null;
            if (File.Exists(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Incidencias del estab ", estab.Trim(), ".xlsb")))
            {
                File.Delete(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Incidencias del estab ", estab.Trim(), ".xlsb"));
            }
            libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Incidencias del estab " + estab.Trim() + ".xlsb", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12);
            libro.Close();
            excel.Quit();
            if (cn.State == ConnectionState.Open)
            {
                cn.Close();
            }
            return string.Concat(System.Windows.Forms.Application.StartupPath, "\\Incidencias del estab ", estab.Trim(), ".xlsb");
        }

        private string CierreCedis(DateTime FechaHora)
        {
            Range rango;
            string str;
            try
            {
                DateTime dia = FechaHora;
                DateTime primero = dia;
                DateTime primeroDelMes = Convert.ToDateTime("01/" + dia.Month.ToString("0#") + "/" + dia.Year.ToString() + " 00:00:00");
                primero = primero.AddHours((double)(primero.Hour * -1));
                primero = primero.AddMinutes((double)(primero.Minute * -1));

                SqlConnection cn = conexion.conectar("BMSNayar");
                SqlCommand comando = new SqlCommand();
                comando.Connection = cn;
                comando.CommandTimeout = 180;
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = comando;

                Microsoft.Office.Interop.Excel.Application excel;
                excel = new Microsoft.Office.Interop.Excel.Application();
                //excel.Application.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook libro;
                libro = excel.Workbooks.Add();
                Worksheet hoja = new Worksheet();

                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                /*Factura detalle,facturacliente**/
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                // faltaban 2 hojas
                libro.Worksheets.Add();
                libro.Worksheets.Add();

                for (int i = 1; i <= 20; i++)
                {
                    System.Data.DataTable dt = new System.Data.DataTable();
                    this.DG1.DataSource = null;
                    this.DG1.Rows.Clear();
                    this.DG1.Columns.Clear();
                    object objeto = null;
                    comando.Parameters.Clear();
                    switch (i)
                    {
                        case 1:
                            {
                                comando.CommandText = "select entysal.transaccion,transacciones.nombre,case when entysal.transaccion in ('35','65') then razones_transferencia.nombre else razones_aod_inventario.nombre end as razon,SUM(entysal.cantidad) as unidades,SUM(entysal.costo) as costo,COUNT(distinct entysal.cod_prod) as codigos,COUNT(distinct entysal.cod_prv)-1 as provedores,COUNT(distinct entysal.folio) as folios from entysal with(nolock) inner join transacciones on entysal.transaccion=transacciones.transaccion left join movimientos_internos with(nolock) on entysal.folio=movimientos_internos.folio and entysal.transaccion=movimientos_internos.transaccion and entysal.cod_estab=movimientos_internos.cod_estab left join razones_aod_inventario on razones_aod_inventario.razon_aod_inventario=movimientos_internos.razon_aod_inventario\tleft join razones_transferencia on movimientos_internos.razon_aod_inventario=razones_transferencia.razon_transferencia where entysal.cod_estab='65' and entysal.fecha between @fechaini and @fechafin group by entysal.transaccion,transacciones.nombre,razones_aod_inventario.nombre,razones_transferencia.nombre order by entysal.transaccion,razones_aod_inventario.nombre,razones_transferencia.nombre";
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Resumen de Movtos";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RESUMEN DE MOVIMIENTOS AL INVENTARIO";
                                rango = (Range)hoja.get_Range("B2", "I2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn column in this.DG1.Columns)
                                {
                                    hoja.Cells[3, column.Index + 2] = column.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, column.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (column.Index)
                                    {
                                        case 0:
                                        case 1:
                                        case 2:
                                            {
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                    }
                                    rango.EntireColumn.NumberFormat = "#,##0";
                                }
                                hoja.Range["B3", string.Concat("I", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("I", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 2:
                            {
                                string str1 = "select movimientos_internos.folio,movimientos_internos.fecha,movimientos_internos.cod_estab_alterno,establecimientos.nombre,movimientos_internos.costo,movimientos_internos.unidades+movimientos_internos.piezas as unidades,movimientos_internos.embarque from movimientos_internos with(nolock) inner join establecimientos on movimientos_internos.cod_estab_alterno=establecimientos.cod_estab where movimientos_internos.cod_estab='65' and movimientos_internos.transaccion='35' and movimientos_internos.fecha between @fechaini and @fechafin and movimientos_internos.status='V' order by CAST(movimientos_internos.cod_estab_alterno as int) asc, movimientos_internos.folio asc";
                                string str2 = str1;
                                comando.CommandText = str1;
                                comando.CommandText = str2;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Relacion de Transferencias";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RELACION DE TRANSFERENCIAS";
                                rango = (Range)hoja.get_Range("B2", "H2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn.Index + 2] = dataGridViewColumn.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 3:
                            {

                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);

                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str1 = "select movimientos_internos.folio,movimientos_internos.fecha,movimientos_internos.cod_estab_alterno,establecimientos.nombre,movimientos_internos.costo,movimientos_internos.unidades+movimientos_internos.piezas as unidades,movimientos_internos.embarque from movimientos_internos with(nolock) inner join establecimientos on movimientos_internos.cod_estab_alterno=establecimientos.cod_estab where movimientos_internos.cod_estab='65' and movimientos_internos.transaccion='35' and movimientos_internos.fecha between @fechaini and @fechafin and movimientos_internos.status='V' order by CAST(movimientos_internos.cod_estab_alterno as int) asc, movimientos_internos.folio asc";
                                string str2 = str1;
                                comando.CommandText = str1;
                                comando.CommandText = str2;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio);
                                comando.Parameters.AddWithValue("@fechafin", Fecha_actual);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Relacion de Transferencias Acum";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RELACION DE TRANSFERENCIAS ACUMULADAS DEL " + Fecha_inicio.ToString("dd/MM/yyyy") + " AL " + Fecha_actual.ToString("dd/MM/yyyy");
                                rango = (Range)hoja.get_Range("B2", "H2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn.Index + 2] = dataGridViewColumn.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 4:
                            {
                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);

                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str1 = "select movimientos_internos.folio,movimientos_internos.fecha,movimientos_internos.cod_estab_alterno,establecimientos.nombre,movimientos_internos.costo,movimientos_internos.unidades+movimientos_internos.piezas as unidades,movimientos_internos.embarque from movimientos_internos with(nolock) inner join establecimientos on movimientos_internos.cod_estab_alterno=establecimientos.cod_estab where movimientos_internos.cod_estab='65' and movimientos_internos.transaccion='35' and movimientos_internos.fecha between @fechaini and @fechafin and movimientos_internos.status='V' order by CAST(movimientos_internos.cod_estab_alterno as int) asc, movimientos_internos.folio asc";
                                string str2 = str1;
                                comando.CommandText = str1;
                                comando.CommandText = str2;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio);
                                comando.Parameters.AddWithValue("@fechafin", Fecha_actual);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Relacion de Transf Acum x Suc";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RELACION DE TRANSFERENCIAS ACUMULADAS X SUCURSAL DEL " + Fecha_inicio.ToString("dd/MM/yyyy") + " AL " + Fecha_actual.ToString("dd/MM/yyyy");
                                rango = (Range)hoja.get_Range("B2", "H2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn.Index + 2] = dataGridViewColumn.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 5:
                            /*actual*/
                            {

                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1).AddMonths(-1);

                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                //string str3 = "select movimientos_internos.folio,razones_transferencia.nombre as [razon transferencia],movimientos_internos.fecha,movimientos_internos.cod_estab_alterno,establecimientos.nombre,movimientos_internos.costo,movimientos_internos.unidades+movimientos_internos.piezas as unidades,movimientos_internos.embarque,movimientos_internos.notas from movimientos_internos with(nolock) inner join establecimientos on movimientos_internos.cod_estab_alterno=establecimientos.cod_estab left join razones_transferencia on razones_transferencia.razon_transferencia=movimientos_internos.razon_aod_inventario where movimientos_internos.cod_estab='65' and movimientos_internos.transaccion='35' and movimientos_internos.status='V' and movimientos_internos.embarque='' and movimientos_internos.razon_aod_inventario<>'5' order by  movimientos_internos.folio asc";
                                string str3 = "select * from(select movimientos_internos.folio , razones_transferencia.nombre as [razon transferencia],movimientos_internos.fecha,movimientos_internos.cod_estab_alterno,establecimientos.nombre,movimientos_internos.costo,movimientos_internos.unidades + movimientos_internos.piezas as unidades,movimientos_internos.embarque,movimientos_internos.notas from movimientos_internos with(nolock) inner join establecimientos on movimientos_internos.cod_estab_alterno = establecimientos.cod_estab left join razones_transferencia on razones_transferencia.razon_transferencia = movimientos_internos.razon_aod_inventario where movimientos_internos.cod_estab = '65' and movimientos_internos.transaccion = '35' and movimientos_internos.status = 'V' and movimientos_internos.embarque = '' and movimientos_internos.razon_aod_inventario <> '5' and movimientos_internos.fecha between @fechaini and @fechafin) as tran_estab  where tran_estab.folio not in (select folio_alterno from movimientos_internos where cod_estab <> '65' and transaccion = '65' and status = 'V' ) order by CAST(tran_estab.cod_estab_alterno as int),tran_estab.fecha,tran_estab.folio asc";
                                string str4 = str3;
                                comando.CommandText = str3;
                                comando.CommandText = str4;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@fechafin", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();

                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Relacion de Trans Sin embarque";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RELACION DE TRANSFERENCIAS SIN EMBARQUE (NO INCLUYE RAZON SOBRANTE DE TRANSFERENCIA)";
                                rango = (Range)hoja.get_Range("B2", "J2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn column1 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, column1.Index + 2] = column1.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, column1.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (column1.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 1:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("J", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("J", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 6:
                            {
                                string str3 = "SELECT e.folio, e.fecha, rt.nombre as [razon transferencia], SUM(e.cantidad) As unidades, SUM(e.costo) As costo, COUNT(DISTINCT e.cod_prod) As codigos FROM entysal e WITH (NoLock) LEFT JOIN movimientos_internos mi WITH(NoLock) ON e.folio = mi.folio AND e.transaccion = mi.transaccion AND e.cod_estab = mi.cod_estab LEFT JOIN razones_transferencia rt WITH(NoLock) ON mi.razon_aod_inventario = rt.razon_transferencia WHERE e.cod_estab='65' AND e.transaccion='35' AND e.fecha BETWEEN @fechaini AND @fechafin GROUP BY e.transaccion, rt.nombre, e.folio, e.fecha ORDER BY e.fecha";
                                string str4 = str3;
                                comando.CommandText = str3;
                                comando.CommandText = str4;
                                comando.Parameters.AddWithValue("@fechaini", primeroDelMes);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();

                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Relacion Det de Transf a Estabs";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RELACION DE DETALLADA DE TRANSFERENCIAS A ESTABLECIMIENTOS";
                                rango = (Range)hoja.get_Range("B2", "G2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn.Index + 2] = dataGridViewColumn.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("G", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 7:
                            {
                                string str5 = "select movimientos_internos.cod_estab_alterno,establecimientos.nombre as estab,SUM(movimientos_internos.unidades+movimientos_internos.piezas) as unidades,SUM(movimientos_internos.costo) as costo,SUM(isnull(membarques.cantidad1,0)) as [cajas de plastico],SUM(ISNULL(membarques.cantidad2,0)) as cartones from embarques left join movimientos_internos on embarques.folio=movimientos_internos.embarque left join membarques on membarques.folio=movimientos_internos.folio and membarques.transaccion=movimientos_internos.transaccion left join establecimientos on establecimientos.cod_estab=movimientos_internos.cod_estab_alterno where embarques.cod_estab='65' and embarques.fecha between @fechaini and @fechafin group by movimientos_internos.cod_estab_alterno,establecimientos.nombre order by CAST(movimientos_internos.cod_estab_alterno as int)";
                                string str6 = str5;
                                comando.CommandText = str5;
                                comando.CommandText = str6;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Concentrado de envios x estab";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "CONCENTRADO DE MERCANCIA ENVIADA POR ESTABLECIMIENTO";
                                rango = (Range)hoja.get_Range("B2", "G2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn1 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn1.Index + 2] = dataGridViewColumn1.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn1.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn1.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("G", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("G", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }

                        case 8:
                            {
                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);

                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str5 = "select embarques.folio as embarque,movimientos_internos.folio,embarques.fecha,movimientos_internos.cod_estab_alterno,establecimientos.nombre as estab,SUM(movimientos_internos.unidades+movimientos_internos.piezas) as unidades,SUM(movimientos_internos.costo) as costo,SUM(isnull(membarques.cantidad1,0)) as [cajas de plastico],SUM(ISNULL(membarques.cantidad2,0)) as cartones from embarques left join movimientos_internos on embarques.folio=movimientos_internos.embarque left join membarques on membarques.folio=movimientos_internos.folio and membarques.transaccion=movimientos_internos.transaccion left join establecimientos on establecimientos.cod_estab=movimientos_internos.cod_estab_alterno where embarques.cod_estab='65' and embarques.fecha between @fechaini and @fechafin group by embarques.folio,movimientos_internos.folio,embarques.fecha,movimientos_internos.cod_estab_alterno,establecimientos.nombre order by CAST(movimientos_internos.cod_estab_alterno as int),embarques.fecha,embarques.folio,movimientos_internos.folio";
                                string str6 = str5;
                                comando.CommandText = str5;
                                comando.CommandText = str6;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@fechafin", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Conc.Acum de env x estab";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "CONCENTRADO ACUMULADO DE MERCANCIA ENVIADA POR ESTABLECIMIENTO";
                                rango = (Range)hoja.get_Range("B2", "J2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn1 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn1.Index + 2] = dataGridViewColumn1.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn1.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn1.Index)
                                    {
                                        case 1:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }

                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;

                                            }

                                        case 7:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;

                                            }
                                        case 8:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;

                                            }



                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("J", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("J", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }

                        case 9:
                            {
                                string str7 = "select pedprv.folio as folio_orden_compra,pedprv.fecha,pedprv.cod_prv,proveedores.razon_social,pedprv.unidades,pedprv.piezas,pedprv.costo,eos_mercancia.folio as folio_entrada,eos_mercancia.fecha as fecha_entrada from pedprv with(nolock) inner join eos_mercancia with(nolock) on pedprv.folio=eos_mercancia.folio_referencia and pedprv.Transaccion=eos_mercancia.transaccion_referencia\tleft join recepcion_mercancia_proveedores with(nolock) on recepcion_mercancia_proveedores.orden_compra=pedprv.folio left join proveedores on proveedores.cod_prv=pedprv.cod_prv where pedprv.cod_estab='65' and eos_mercancia.fecha between @fechaini and @fechafin and eos_mercancia.transaccion_referencia='30' and recepcion_mercancia_proveedores.Folio is null";
                                string str8 = str7;
                                comando.CommandText = str7;
                                comando.CommandText = str8;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Ordenes de compra sin recepcion";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "ORDENES DE COMPRA CON ENTRADA SIN RECEPCION";
                                rango = (Range)hoja.get_Range("B2", "j2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn column2 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, column2.Index + 2] = column2.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, column2.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (column2.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 7:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("J", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("J", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 10:
                            {
                                string str9 = "select movimientos_internos.cod_estab,establecimientos.nombre,movimientos_internos.folio,movimientos_internos.fecha,movimientos_internos.costo,movimientos_internos.unidades+movimientos_internos.piezas as unidades,replace(replace(REPLACE(movimientos_internos.notas,CHAR(9),''),char(10),''),char(13),'') as notas from movimientos_internos with(nolock) inner join establecimientos on movimientos_internos.cod_estab=establecimientos.cod_estab where transaccion='35' and cod_estab_alterno='65' and movimientos_internos.status='V' and folio in (select transferencia from mercancia_transito_establecimientos where cod_estab_destino='65') order by movimientos_internos.fecha";
                                string str10 = str9;
                                comando.CommandText = str9;
                                comando.CommandText = str10;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Transferencias pendientes CEDIS";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "TRANSFERENCIAS CON MERCANCIA PENDIENTE POR RECIBIR EN CEDIS";
                                rango = (Range)hoja.get_Range("B2", "H2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn2 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn2.Index + 2] = dataGridViewColumn2.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn2.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn2.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 11:
                            {
                                string str11 = "select movimientos_internos.usuario,upper(usuarios.nombre) as nombre,count(distinct movimientos_internos.folio) as transferencias,COUNT(distinct case movimientos_internos.status when 'V' then movimientos_internos.folio end) as transferencias_vigentes,COUNT(distinct case movimientos_internos.status when 'C' then movimientos_internos.folio end) as transferencias_canceladas,SUM(entysal.cantidad) as unidades,SUM(entysal.costo) as costo,COUNT(entysal.cod_prod) as productos from movimientos_internos with(nolock) inner join entysal  with(nolock) on movimientos_internos.folio=entysal.folio and movimientos_internos.transaccion=entysal.transaccion left join usuarios on usuarios.usuario=movimientos_internos.usuario where movimientos_internos.transaccion='35' and movimientos_internos.cod_estab='65' and movimientos_internos.fecha between @fechaini and @fechafin group by movimientos_internos.usuario,usuarios.nombre,movimientos_internos.status order by SUM(entysal.cantidad) desc";
                                string str12 = str11;
                                comando.CommandText = str11;
                                comando.CommandText = str12;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Transferencias por usuario";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "TRANSFERENCIAS POR USUARIO";
                                rango = (Range)hoja.get_Range("B2", "I2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn column3 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, column3.Index + 2] = column3.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, column3.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (column3.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 7:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("I", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("I", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }

                        case 12:
                            {
                                string str13 = "select eos_mercancia.usuario,UPPER(usuarios.nombre) as nombre,COUNT(distinct eos_mercancia.folio) as entradas,COUNT(distinct case when eos_mercancia.status in ('R','V') then eos_mercancia.folio end) as entradas_vigentes,COUNT(distinct case when eos_mercancia.status='C' then eos_mercancia.folio end) as entradas_canceladas,SUM(m_eos_mercancia.cantidad) as unidades,COUNT(m_eos_mercancia.cod_prod) as productos from eos_mercancia with(nolock) inner join m_eos_mercancia with(nolock) on eos_mercancia.folio=m_eos_mercancia.folio and eos_mercancia.transaccion=m_eos_mercancia.transaccion left join usuarios on eos_mercancia.usuario=usuarios.usuario where eos_mercancia.cod_estab='65' and eos_mercancia.transaccion='56' and eos_mercancia.fecha between @fechaini and @fechafin group by eos_mercancia.usuario,usuarios.nombre";
                                string str14 = str13;
                                comando.CommandText = str13;
                                comando.CommandText = str14;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Entradas por usuario";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "ENTRADAS POR USUARIO";
                                rango = (Range)hoja.get_Range("B2", "H2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn3 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn3.Index + 2] = dataGridViewColumn3.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn3.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn3.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }

                        case 13:
                            {
                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);

                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str13 = "select eos_mercancia.usuario,UPPER(usuarios.nombre) as nombre,COUNT(distinct eos_mercancia.folio) as entradas,COUNT(distinct case when eos_mercancia.status in ('R','V') then eos_mercancia.folio end) as entradas_vigentes,COUNT(distinct case when eos_mercancia.status='C' then eos_mercancia.folio end) as entradas_canceladas,SUM(m_eos_mercancia.cantidad) as unidades,COUNT(m_eos_mercancia.cod_prod) as productos from eos_mercancia with(nolock) inner join m_eos_mercancia with(nolock) on eos_mercancia.folio=m_eos_mercancia.folio and eos_mercancia.transaccion=m_eos_mercancia.transaccion left join usuarios on eos_mercancia.usuario=usuarios.usuario where eos_mercancia.cod_estab='65' and eos_mercancia.transaccion='56' and eos_mercancia.fecha between @fechaini and @fechafin group by eos_mercancia.usuario,usuarios.nombre";
                                string str14 = str13;
                                comando.CommandText = str13;
                                comando.CommandText = str14;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@fechafin", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");

                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Entradas x usuario_acum";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "ENTRADAS POR USUARIO MES EN CURSO";
                                rango = (Range)hoja.get_Range("B2", "H2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn3 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn3.Index + 2] = dataGridViewColumn3.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn3.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn3.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("H", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;

                            }

                        case 14:
                            {
                                string str15 = "select rm.cod_prv,ltrim(rtrim(p.razon_social)) as Proveedor,COUNT(distinct(folio)) as [Cantidad de Recepciones],SUM(rm.unidades) as Unidades,sum(rm.costo) as Costo from recepcion_mercancia_proveedores as rm inner join proveedores as p on rm.cod_prv=p.cod_prv where rm.Fecha between @fechaini and @fechafin  and rm.status='V' group by\trm.cod_prv,p.razon_social";
                                string str16 = str15;
                                comando.CommandText = str15;
                                comando.CommandText = str16;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Recepciones x Prov";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RECEPCIONES POR PROVEEDOR";
                                rango = (Range)hoja.get_Range("B2", "F2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn column4 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, column4.Index + 2] = column4.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, column4.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (column4.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("F", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("F", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }

                        case 15:
                            {

                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;
                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);
                                DateTime Fecha_Final = new DateTime();
                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str15 = "select rm.cod_prv,ltrim(rtrim(p.razon_social)) as Proveedor,COUNT(distinct(folio)) as [Cantidad de Recepciones],SUM(rm.unidades) as Unidades,sum(rm.costo) as Costo from recepcion_mercancia_proveedores as rm inner join proveedores as p on rm.cod_prv=p.cod_prv where rm.Fecha between @fechaini and @fechafin  and rm.status='V' group by\trm.cod_prv,p.razon_social";
                                string str16 = str15;
                                comando.CommandText = str15;
                                comando.CommandText = str16;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@fechafin", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Recepciones x Prov Mensual";
                                if (objeto == null)
                                {
                                    break;
                                }

                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RECEPCIONES POR PROVEEDOR MENSUAL";
                                rango = (Range)hoja.get_Range("B2", "F2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn column4 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, column4.Index + 2] = column4.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, column4.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (column4.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("F", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("F", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();

                                break;
                            }

                        case 16:
                            {

                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;
                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);
                                DateTime Fecha_Final = new DateTime();
                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str15 = "select folio as [Folio],rm.fecha as [Fecha],rm.cod_prv,ltrim(rtrim(p.razon_social)) as Proveedor,rm.orden_compra,rm.entrada,rm.importe,rm.iva,rm.total,rm.costo as Costo, rm.unidades as Unidades from recepcion_mercancia_proveedores as rm inner join proveedores as p on rm.cod_prv = p.cod_prv where rm.Fecha between @fechaini and @fechafin and rm.status = 'V' group by rm.cod_prv,p.razon_social,rm.orden_compra,rm.entrada,rm.folio,rm.fecha,rm.importe,rm.iva,rm.total,rm.costo,rm.unidades";
                                string str16 = str15;
                                comando.CommandText = str15;
                                comando.CommandText = str16;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@fechafin", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = " Detalle Recep. x Prov Mensual";
                                if (objeto == null)
                                {
                                    break;
                                }

                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "DETALLE MENSUAL DE RECEPCIONES POR PROVEEDOR ";
                                rango = (Range)hoja.get_Range("B2", "L2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn column4 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, column4.Index + 2] = column4.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, column4.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (column4.Index)
                                    {
                                        case 1:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 7:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 8:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 9:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 10:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 11:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("L", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("L", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();

                                break;
                            }
                        case 17:
                            {
                                string str17 = "select lp.nombre,SUM(e.cantidad) as Unidades,sum(e.costo) as Costo,isnull(sum(rmg.gastos_compras),0) as Gasto from recepcion_mercancia_proveedores as rm inner join entysal as e on rm.Folio=e.folio and rm.Transaccion=e.transaccion left join mrecepcion_mercancia_gastos_compras as rmg on rmg.id_entysal=e.id left join productos as p on e.cod_prod=p.cod_prod left join lineas_productos as lp on p.linea_producto=lp.linea_producto where rm.transaccion='44' and rm.Fecha between @fechaini and @fechafin  and rm.status='V' group by lp.nombre";
                                string str18 = str17;
                                comando.CommandText = str17;
                                comando.CommandText = str18;
                                comando.Parameters.AddWithValue("@fechaini", primero);
                                comando.Parameters.AddWithValue("@fechafin", dia);
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Recepciones x Linea";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RECEPCIONES POR LINEA";
                                rango = (Range)hoja.get_Range("B2", "E2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn4 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn4.Index + 2] = dataGridViewColumn4.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn4.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn4.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 1:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("E", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("E", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }

                        case 18:
                            {

                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);


                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str17 = "select lp.nombre,SUM(e.cantidad) as Unidades,sum(e.costo) as Costo,isnull(sum(rmg.gastos_compras),0) as Gasto from recepcion_mercancia_proveedores as rm inner join entysal as e on rm.Folio=e.folio and rm.Transaccion=e.transaccion left join mrecepcion_mercancia_gastos_compras as rmg on rmg.id_entysal=e.id left join productos as p on e.cod_prod=p.cod_prod left join lineas_productos as lp on p.linea_producto=lp.linea_producto where rm.transaccion='44' and rm.Fecha between @fechaini and @fechafin  and rm.status='V' group by lp.nombre";
                                string str18 = str17;
                                comando.CommandText = str17;
                                comando.CommandText = str18;
                                comando.Parameters.AddWithValue("@fechaini", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@fechafin", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();
                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Recepciones x Linea Mensual";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "RECEPCIONES POR LINEA MENSUAL";
                                rango = (Range)hoja.get_Range("B2", "E2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                foreach (DataGridViewColumn dataGridViewColumn4 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn4.Index + 2] = dataGridViewColumn4.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn4.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn4.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 1:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0";
                                                continue;
                                            }
                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }
                                hoja.Range["B3", string.Concat("E", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("E", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();
                                break;
                            }
                        case 19:
                            {
                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);


                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str17 = "SELECT  entysal.folio as [FOLIO],entysal.fecha as [FECHA], entysal.cod_prod AS[CODIGO PRODUCTO], productos.descripcion_completa AS[DESCRIPCION], entysal.cantidad AS[CANTIDAD], entysal.precio_lista AS[LISTA PRECIO], entysal.importe_descuento AS[IMPORTE DESCUENTO], entysal.importe AS[IMPORTE], entysal.iva as [IVA],/* entysal.abreviatura_unidad,*/ entysal.costo AS[COSTO],/*, dbo.Comentario(entysal.id) AS comentario, entysal.cod_estab, ISNULL(establecimientos.nombre, '') AS establecimiento_nombre, ISNULL(entysal.ieps, 0) AS ieps,*/ entysal.total AS[TOTAL] FROM entysal with(nolock) LEFT OUTER JOIN establecimientos WITH(nolock) ON entysal.cod_estab = establecimientos.cod_estab LEFT OUTER JOIN productos WITH(nolock) ON entysal.cod_prod = productos.cod_prod WHERE entysal.folio in (SELECT folios.folio FROM( SELECT facremtick.folio, transacciones.transaccion,transacciones.abreviatura AS Abr,facremtick.fecha,facremtick.importe,facremtick.iva,facremtick.fecha_cancelacion, facremtick.ieps, facremtick.costo,facremtick.importe * facremtick.tipo_cambio - facremtick.costo as contribucion, facremtick.total,facremtick.tipo_cambio,facremtick.cod_estab, monedas.abreviatura,cast(1 as bit) as recoge_mercancia, facremtick.status,facremtick.folio_origen,facremtick.importe_descuento,facremtick.cond_pago,condiciones_pago.nombre as nombre_condpago, facremtick.cliente_alterno, facremtick.tipo_cliente_alterno,facremtick.vendedor, ISNULL(V.nombre, '') AS NombreVendedor, case when cliente_alterno = '' then c.razon_social else dbo.ClienteAlterno(facremtick.cliente_alterno, facremtick.tipo_cliente_alterno) end as razon_social, case when cliente_alterno = '' then facremtick.cod_cte else facremtick.cliente_alterno end as cod_cte, facremtick.notas,facremtick.usuario, usuarios.nombre as Nombre_usuario, ISNULL(status_ubicacion.abreviatura, '') AS StatusUbicacion, dbo.AbonosFRT(facremtick.folio,facremtick.transaccion, @FF)as abonos, dbo.SaldoDocCteFecha(facremtick.folio,facremtick.transaccion, @FF ) as saldo, /* @Enganche as enganche,*/c.codigo_nuestro_cliente,'' as chofer,c.nom_comercial, dbo.domicilio('C',c.cod_cte) AS domicilio, facremtick.transaccion_origen,c.plazo,/* @Saldo as SaldoCliente, */c.pobmunedo as poblacion_cliente, case transaccion_origen when '31' then isnull((select fecha from pedcte where folio = facremtick.folio_origen and facremtick.transaccion_origen = '31'),'') else '' end as fecha_folio_origen, isnull(V.tipo_vendedor, '') as tipo_vendedor, isnull(tipos_vendedores.nombre, '') as nombre_tipo_vendedor, ISNULL(facremtick.linea_fletera, '') AS linea_fletera, dbo.PoblacionDoc(facremtick.folio, facremtick.transaccion, 'N') AS destino, isnull((select sum(case when e.unidad = 'U' then e.cantidad else (e.cantidad / case when p.contenido_presentacion=0 then 1 else p.contenido_presentacion end) end) from entysal e inner join productos p on p.cod_prod = e.cod_prod where e.folio = facremtick.folio and transaccion = facremtick.transaccion),0) AS cajas, isnull(proveedores.razon_social,'') as Razon_social1, facremtick.embarque, isnull(razones_cancelacion_ventas.nombre,'') AS NomRazon, estab.nombre as NombreEstab,0 as ivaRetenido,0 as isrRetenido, facremtick.pedido_cliente, facremtick.condicion_financiera,ISNULL(condiciones_venta_financieras.nombre, '') AS NombreCondFin,ISNULL(UC.nombre, '') AS Nombre_UsuarioCan,  isnull((select notas2 from pedcte where folio = facremtick.folio_origen and facremtick.transaccion_origen = '31'),'') as notas2, c.nom_comercial as nombre_comercial,isnull(segmentos.nombre,'''') as nombre_segmento,'ejemplo' as familia,isnull(condiciones_distribucion.nombre,'') as condicion_distribucion, isnull(FD.fecha_recepcion_mercancia,'') as fecha_recepcion_mercancia,facremtick.iva_retenido, facremtick.isr_retenido, monedas.nombre as moneda,monedas.moneda as tipo_moneda FROM facremtick with(nolock) LEFT JOIN condiciones_pago with(nolock)  ON facremtick.cond_pago = condiciones_pago.condicion_pago LEFT OUTER JOIN vendedores V  with(nolock)  ON facremtick.vendedor = V.vendedor LEFT OUTER JOIN tipos_vendedores with(nolock)  ON tipos_vendedores.tipo_vendedor = V.tipo_vendedor LEFT OUTER JOIN usuarios with(nolock)  ON facremtick.usuario = usuarios.usuario LEFT OUTER JOIN usuarios UC with(nolock)  ON facremtick.usuario_cancelacion = UC.usuario INNER JOIN transacciones  with(nolock)  ON facremtick.transaccion = transacciones.transaccion LEFT OUTER JOIN razones_cancelacion_ventas with(nolock)  ON facremtick.razon_cancelacion_ventas = razones_cancelacion_ventas.razon_cancelacion_ventas LEFT OUTER JOIN status_ubicacion with(nolock)  ON facremtick.status_ubicacion = status_ubicacion.status_ubicacion INNER JOIN monedas  with(nolock)  ON facremtick.moneda = monedas.moneda LEFT OUTER JOIN  embarques  with(nolock) ON embarques.folio = facremtick.embarque LEFT OUTER JOIN domicilios_consignacion  with(nolock)  ON domicilios_consignacion.numdpc = facremtick.numdpc LEFT OUTER JOIN proveedores with(nolock)  ON facremtick.linea_fletera = proveedores.cod_prv INNER JOIN clientes c with(nolock)  on facremtick.cod_cte = c.cod_cte INNER JOIN establecimientos estab with(nolock)  ON facremtick.cod_estab = estab.cod_estab LEFT OUTER JOIN segmentos_mercado segmentos with(nolock)  ON segmentos.segmento=c.segmento LEFT OUTER JOIN condiciones_venta_financieras with(nolock)  ON facremtick.condicion_financiera = condiciones_venta_financieras.Folio LEFT OUTER JOIN condiciones_distribucion with(nolock)  on facremtick.condicion_distribucion=condiciones_distribucion.condicion_distribucion LEFT JOIN facremtick_datos FD with(nolock)  ON FD.folio = facremtick.folio and FD.transaccion = facremtick.transaccion    where facremtick.transaccion ='36' and facremtick.fecha  BETWEEN @FI AND @FF and facremtick.cod_estab='65') AS FOLIOS ) ORDER BY entysal.id,entysal.fecha asc";
                                string str18 = str17;
                                comando.CommandText = str17;
                                comando.CommandText = str18;
                                comando.Parameters.AddWithValue("@FI", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@FF", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();

                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Detalle Factura cliente";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "DETALLE DE FACTURAS CLIENTES";
                                rango = (Range)hoja.get_Range("B2", "L2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;

                                rango = (Range)hoja.Cells[3, 10];
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;



                                foreach (DataGridViewColumn dataGridViewColumn4 in this.DG1.Columns)
                                {
                                    hoja.Cells[3, dataGridViewColumn4.Index + 2] = dataGridViewColumn4.Name.ToString().ToUpper();
                                    rango = (Range)hoja.Cells[3, dataGridViewColumn4.Index + 2];
                                    rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                    rango.Borders.Weight = XlBorderWeight.xlMedium;
                                    rango.Font.Bold = true;
                                    rango.Font.Name = "Consolas";
                                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    switch (dataGridViewColumn4.Index)
                                    {
                                        case 0:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 1:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }

                                        case 2:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }
                                        case 3:
                                            {
                                                rango.EntireColumn.NumberFormat = "@";
                                                continue;
                                            }

                                        case 4:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##";
                                                continue;
                                            }
                                        case 5:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 6:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 7:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 8:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }
                                        case 9:
                                            {
                                                rango.EntireColumn.NumberFormat = "#,##0.00";
                                                continue;
                                            }


                                        default:
                                            {
                                                continue;
                                            }
                                    }
                                }


                                hoja.Range["B3", string.Concat("L", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("L", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();

                                break;
                            }
                        case 20:
                            {
                                DateTime Fecha_actual = new DateTime();
                                Fecha_actual = DateTime.Now;

                                DateTime Fecha_inicio = new DateTime(Fecha_actual.Year, Fecha_actual.Month, 1);


                                DateTime Fecha_Final = new DateTime();

                                Fecha_Final = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1).AddDays(-1);

                                string str17 = "SELECT facremtick.folio,/*transacciones.transaccion,*/transacciones.abreviatura AS Abr,facremtick.fecha,facremtick.importe,facremtick.iva,/*facremtick.fecha_cancelacion,*facremtick.ieps,*/ facremtick.costo,facremtick.importe * facremtick.tipo_cambio - facremtick.costo as contribucion, facremtick.total,isnull((select sum(case when e.unidad = 'U' then e.cantidad else (e.cantidad / case when p.contenido_presentacion=0 then 1 else p.contenido_presentacion end) end) from entysal e inner join productos p on p.cod_prod = e.cod_prod where e.folio = facremtick.folio and transaccion = facremtick.transaccion),0) AS Unidades, /* facremtick.tipo_cambio,facremtick.cod_estab, monedas.abreviatura,cast(1 as bit) as recoge_mercancia, facremtick.status,*/facremtick.folio_origen,case transaccion_origen when '31' then isnull((select fecha from pedcte where folio = facremtick.folio_origen and facremtick.transaccion_origen = '31'),'') else '' end as fecha_folio_origen ,/*facremtick.importe_descuento,facremtick.cond_pago,condiciones_pago.nombre as nombre_condpago, facremtick.cliente_alterno, facremtick.tipo_cliente_alterno,facremtick.vendedor, ISNULL(V.nombre, '') AS NombreVendedor, */case when cliente_alterno = '' then facremtick.cod_cte else facremtick.cliente_alterno end as cod_cte, case when cliente_alterno = '' then c.razon_social else dbo.ClienteAlterno(facremtick.cliente_alterno, facremtick.tipo_cliente_alterno) end as razon_social, facremtick.notas/*facremtick.usuario, usuarios.nombre as Nombre_usuario,ISNULL(status_ubicacion.abreviatura, '') AS StatusUbicacion,dbo.AbonosFRT(facremtick.folio,facremtick.transaccion, @FF)as abonos,dbo.SaldoDocCteFecha(facremtick.folio,facremtick.transaccion, @FF ) as saldo, @Enganche as enganche,c.codigo_nuestro_cliente,'' as chofer,c.nom_comercial, dbo.domicilio('C',c.cod_cte) AS domicilio, facremtick.transaccion_origen,c.plazo,@Saldo as SaldoCliente, c.pobmunedo as poblacion_cliente,isnull(V.tipo_vendedor, '') as tipo_vendedor,isnull(tipos_vendedores.nombre, '') as nombre_tipo_vendedor, ISNULL(facremtick.linea_fletera, '') AS linea_fletera,dbo.PoblacionDoc(facremtick.folio, facremtick.transaccion, 'N') AS destino, isnull(proveedores.razon_social,'') as Razon_social1, facremtick.embarque,isnull(razones_cancelacion_ventas.nombre,'') AS NomRazon,estab.nombre as NombreEstab,0 as ivaRetenido,0 as isrRetenido, facremtick.pedido_cliente, facremtick.condicion_financiera,ISNULL(condiciones_venta_financieras.nombre, '') AS NombreCondFin,ISNULL(UC.nombre, '') AS Nombre_UsuarioCan,isnull((select notas2 from pedcte where folio = facremtick.folio_origen and facremtick.transaccion_origen = '31'),'') as notas2,c.nom_comercial as nombre_comercial,isnull(segmentos.nombre,'') as nombre_segmento,'ejemplo' as familia,isnull(condiciones_distribucion.nombre,'') as condicion_distribucion,isnull(FD.fecha_recepcion_mercancia,'') as fecha_recepcion_mercancia,facremtick.iva_retenido, facremtick.isr_retenido, monedas.nombre as moneda,monedas.moneda as tipo_moneda*/ FROM facremtick with(nolock) LEFT JOIN condiciones_pago with(nolock)  ON facremtick.cond_pago = condiciones_pago.condicion_pago LEFT OUTER JOIN vendedores V  with(nolock)  ON facremtick.vendedor = V.vendedor LEFT OUTER JOIN tipos_vendedores with(nolock)  ON tipos_vendedores.tipo_vendedor = V.tipo_vendedor LEFT OUTER JOIN usuarios with(nolock)  ON facremtick.usuario = usuarios.usuario LEFT OUTER JOIN usuarios UC with(nolock)  ON facremtick.usuario_cancelacion = UC.usuario INNER JOIN transacciones  with(nolock)  ON facremtick.transaccion = transacciones.transaccion LEFT OUTER JOIN razones_cancelacion_ventas with(nolock)  ON facremtick.razon_cancelacion_ventas = razones_cancelacion_ventas.razon_cancelacion_ventas LEFT OUTER JOIN status_ubicacion with(nolock)  ON facremtick.status_ubicacion = status_ubicacion.status_ubicacion INNER JOIN monedas  with(nolock)  ON facremtick.moneda = monedas.moneda LEFT OUTER JOIN  embarques  with(nolock) ON embarques.folio = facremtick.embarque LEFT OUTER JOIN domicilios_consignacion  with(nolock)  ON domicilios_consignacion.numdpc = facremtick.numdpc LEFT OUTER JOIN proveedores with(nolock)  ON facremtick.linea_fletera = proveedores.cod_prv INNER JOIN clientes c with(nolock)  on facremtick.cod_cte = c.cod_cte INNER JOIN establecimientos estab with(nolock)  ON facremtick.cod_estab = estab.cod_estab LEFT OUTER JOIN segmentos_mercado segmentos with(nolock)  ON segmentos.segmento=c.segmento LEFT OUTER JOIN condiciones_venta_financieras with(nolock)  ON facremtick.condicion_financiera = condiciones_venta_financieras.Folio LEFT OUTER JOIN condiciones_distribucion with(nolock)  on facremtick.condicion_distribucion=condiciones_distribucion.condicion_distribucion LEFT JOIN facremtick_datos FD with(nolock)  ON FD.folio = facremtick.folio and FD.transaccion = facremtick.transaccion    where facremtick.transaccion ='36' and facremtick.fecha  BETWEEN @FI AND @FF and facremtick.cod_estab='65' order by facremtick.folio asc";
                                string str18 = str17;
                                comando.CommandText = str17;
                                comando.CommandText = str18;
                                comando.Parameters.AddWithValue("@FI", Fecha_inicio.ToString("dd/MM/yyyy") + " 00:00:00");
                                comando.Parameters.AddWithValue("@FF", Fecha_Final.ToString("dd/MM/yyyy") + " 23:59:59");
                                da.Fill(dt);
                                this.DG1.DataSource = dt;
                                this.DG1.SelectAll();

                                objeto = this.DG1.GetClipboardContent();
                                hoja = (Worksheet)libro.Worksheets.get_Item(i);
                                hoja.Activate();
                                hoja.Name = "Facturas Clientes";
                                if (objeto == null)
                                {
                                    break;
                                }
                                Clipboard.SetDataObject(objeto);
                                hoja.Cells[2, 2] = "Facturas a Clientes";
                                rango = (Range)hoja.get_Range("B2", "O2");
                                rango.Merge();
                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;

                                hoja.Cells[3, 2] = "FOLIO";
                                hoja.Cells[3, 3] = "ABR";
                                hoja.Cells[3, 4] = "FECHA";
                                hoja.Cells[3, 5] = "IMPORTE";
                                hoja.Cells[3, 5].EntireColumn.NumberFormat = "#,##0.00";

                                hoja.Cells[3, 6] = "IVA";
                                hoja.Cells[3, 6].EntireColumn.NumberFormat = "#,##0.00";
                                hoja.Cells[3, 7] = "COSTO";
                                hoja.Cells[3, 7].EntireColumn.NumberFormat = "#,##0.00";
                                hoja.Cells[3, 8] = "CONTRIBUCION";
                                hoja.Cells[3, 8].EntireColumn.NumberFormat = "#,##0.00";
                                hoja.Cells[3, 9] = "TOTAL";
                                hoja.Cells[3, 9].EntireColumn.NumberFormat = "#,##0.00";
                                hoja.Cells[3, 10] = "UNIDADES";
                                hoja.Cells[3, 10].EntireColumn.NumberFormat = "#,##0";

                                hoja.Cells[3, 11] = "FOLIO ORIGEN";
                                hoja.Cells[3, 11].EntireColumn.NumberFormat = "@";

                                hoja.Cells[3, 12] = "FECHA FOLIO ORIGEN";
                                hoja.Cells[3, 12].EntireColumn.NumberFormat = "@";
                                hoja.Cells[3, 13] = "COD CLIENTE";
                                hoja.Cells[3, 13].EntireColumn.NumberFormat = "@";
                                hoja.Cells[3, 14] = "RAZON SOCIAL";
                                hoja.Cells[3, 14].EntireColumn.NumberFormat = "@";
                                hoja.Cells[3, 15] = "NOTAS";
                                hoja.Cells[3, 15].EntireColumn.NumberFormat = "@";



                                rango = (Range)hoja.get_Range("B3", "O3");

                                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                                rango.Borders.Weight = XlBorderWeight.xlMedium;
                                rango.Font.Bold = true;
                                rango.Font.Name = "Consolas";
                                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                rango.VerticalAlignment = XlVAlign.xlVAlignCenter;

                                hoja.Range["B3", string.Concat("O", Convert.ToString(this.DG1.Rows.Count + 2))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                                rango = (Range)hoja.Cells[4, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                                rango = hoja.Range["A1", string.Concat("O", Convert.ToString(this.DG1.Rows.Count + 1))];
                                rango.EntireColumn.AutoFit();

                                break;
                            }

                    }
                }
                hoja = (Worksheet)libro.Sheets[1];
                hoja.Activate();
                if (cn.State.ToString() == "Open")
                {
                    cn.Close();
                }
                if (File.Exists(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Informe de Cierre de Operaciones al ", dia.ToString("dd MMMM yyyy"), ".xlsb")))
                {
                    File.Delete(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Informe de Cierre de Operaciones al ", dia.ToString("dd MMMM yyyy"), ".xlsb"));
                }
                libro.SaveAs(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Informe de Cierre de Operaciones al ", dia.ToString("dd MMMM yyyy"), ".xlsb"), Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12);
                libro.Close();
                excel.Quit();
                str = string.Concat(System.Windows.Forms.Application.StartupPath, "\\Informe de Cierre de Operaciones al ", dia.ToString("dd MMMM yyyy"), ".xlsb");
            }
            catch (Exception ex)
            {
                lblError.Text = "Hubo un error: " + ex.Message.ToString();
                str = "";
            }
            return str;
        }

        private void Presupuesto()
        {
            SqlConnection cn = conexion.conectar("BMSNayar");
            SqlDataReader dr = (new SqlCommand("select establecimientos.cod_estab,establecimientos.nombre,mail.email,mail.email_cordinador  from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab=mail.cod_estab where establecimientos.status='V' and establecimientos.cod_estab not in ('65','1001','1002','1003','1004','1005','1006') order by cast(establecimientos.cod_estab as int) asc", cn)).ExecuteReader();
            if (dr.HasRows)
            {
                DateTime dia = DateTime.Now;
                while (dr.Read())
                {
                    string str = string.Concat(dr["email"].ToString().Trim(), ",", dr["email_cordinador"].ToString().Trim());
                    if (dr["cod_estab"].ToString().Trim() == "1")
                    {
                        str = string.Concat(str, ",eduardo@mercadodeimportaciones.com,mercadotecniaauxiliar@mercadodeimportaciones.com,maferperezle01@gmail.com");
                    }
                    // Este if se puede quitar, pero se puso para que le lleguen los correos al jefe de sistemas y pueda estar validando que el envío automático de los reportes se está haciendo.
                    //else if ((dr["cod_estab"].ToString().Trim() == "9") || (dr["cod_estab"].ToString().Trim() == "3"))
                    //{
                    str = string.Concat(str, ",analista.comercial@mercadodeimportaciones.com,annacelia.soto@mercadodeimportaciones.com,luis.cota@mercadodeimportaciones.com,mario.serrano@mercadodeimportaciones.com,auxiliar.ventas@mercadodeimportaciones.com");
                    //}

                    //this.PresupuestoSucursal(dr["cod_estab"].ToString().Trim(), dr["nombre"].ToString(), dia);

                    try
                    {
                        this.EnviaMailGmail(this.PresupuestoSucursal(dr["cod_estab"].ToString().Trim(), dr["nombre"].ToString(), dia), str);
                    }
                    catch (Exception e)
                    {
                        // Si no se hace nada

                    }
                }
            }
        }

        private string PresupuestoSucursal(string cod_estab, string estab, DateTime FechaHora)
        {
            try
            {
                DateTime dia = FechaHora;
                SqlConnection cn = conexion.conectar("BMSNayar");
                SqlCommand comando = new SqlCommand();
                comando.Connection = cn;
                comando.CommandType = CommandType.StoredProcedure;
                comando.CommandText = "MI_ProyectadoPresupuesto";
                comando.CommandTimeout = 240;

                Microsoft.Office.Interop.Excel.Application excel;
                excel = new Microsoft.Office.Interop.Excel.Application();
                //excel.Application.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook libro;
                libro = excel.Workbooks.Add();
                Worksheet hoja = new Worksheet();

                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@fecha", dia);
                if (cod_estab.Trim() != "1")
                {
                    comando.Parameters.AddWithValue("@estab", cod_estab);
                }
                else
                {
                    comando.Parameters.AddWithValue("@estab", "T");
                }
                comando.Parameters.AddWithValue("@tipo_reporte", 1);
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                hoja.Activate();
                hoja.Name = "Indicador Concentrado";
                hoja.Cells[2, 2] = "INDICADOR CONCENTRADO DE PRESUPUESTO";
                Microsoft.Office.Interop.Excel.Range rango = (Range)hoja.get_Range("B2", "K2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                rango.Borders.Weight = XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "BOLD";
                rango.Font.Name = "Consolas";
                rango.Interior.Color = Color.Gainsboro;
                SqlDataAdapter da = new SqlDataAdapter(comando);
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt;
                this.DG1.SelectAll();
                object objeto = this.DG1.GetClipboardContent();
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    foreach (DataGridViewColumn column in this.DG1.Columns)
                    {
                        hoja.Cells[3, column.Index + 2] = column.Name.ToString().ToUpper();
                        rango = (Range)hoja.Cells[3, column.Index + 2];
                        rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                        rango.Borders.Weight = XlBorderWeight.xlMedium;
                        rango.Font.Bold = true;
                        rango.Font.Name = "Consolas";
                        rango.WrapText = true;
                        rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        if (column.Index <= 1)
                        {
                            continue;
                        }
                        rango.EntireColumn.NumberFormat = "#,##0.00";
                    }
                    rango = (Range)hoja.Cells[4, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    hoja.Cells[5 + this.DG1.Rows.Count, 3] = "TOTAL";
                    ((Range)hoja.get_Range(string.Concat("C", Convert.ToString(5 + this.DG1.Rows.Count)), string.Concat("J", Convert.ToString(5 + this.DG1.Rows.Count)))).Font.Bold = true;
                    ((Range)hoja.get_Range("B4", string.Concat("K", Convert.ToString(5 + this.DG1.Rows.Count)))).Font.Name = "Consolas";
                    ((dynamic)hoja.Cells[5 + this.DG1.Rows.Count, 4]).Formula = string.Concat("=SUM(D4:D", Convert.ToString(4 + this.DG1.Rows.Count), ")");
                    ((dynamic)hoja.Cells[5 + this.DG1.Rows.Count, 5]).Formula = string.Concat("=SUM(E4:E", Convert.ToString(4 + this.DG1.Rows.Count), ")");
                    dynamic cells = hoja.Cells[5 + this.DG1.Rows.Count, 6];
                    string[] str = new string[] { "=(E", Convert.ToString(5 + this.DG1.Rows.Count), "/D", Convert.ToString(5 + this.DG1.Rows.Count), ")*100" };
                    cells.Formula = string.Concat(str);
                    hoja.Cells[5 + this.DG1.Rows.Count, 7].Formula = string.Concat("=SUM(G4:G", Convert.ToString(4 + this.DG1.Rows.Count), ")");
                    dynamic obj = hoja.Cells[5 + this.DG1.Rows.Count, 8];
                    string[] strArrays = new string[] { "=(G", Convert.ToString(5 + this.DG1.Rows.Count), "/D", Convert.ToString(5 + this.DG1.Rows.Count), ")*100" };
                    obj.Formula = string.Concat(strArrays);
                    hoja.Cells[5 + this.DG1.Rows.Count, 9].Formula = string.Concat("=G", Convert.ToString(5 + this.DG1.Rows.Count), "-D", Convert.ToString(5 + this.DG1.Rows.Count));
                    dynamic cells1 = hoja.Cells[5 + this.DG1.Rows.Count, 10];
                    string[] str1 = new string[] { "=((G", Convert.ToString(5 + this.DG1.Rows.Count), "/D", Convert.ToString(5 + this.DG1.Rows.Count), ")-1)*100" };
                    cells1.Formula = string.Concat(str1);
                    ((Range)hoja.get_Range("B2", string.Concat("K", Convert.ToString(5 + this.DG1.Rows.Count)))).BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                    ((Range)hoja.get_Range("B2", string.Concat("K", Convert.ToString(5 + this.DG1.Rows.Count)))).EntireColumn.AutoFit();
                    IconSetCondition iconSets = null;
                    iconSets = (IconSetCondition)((Range)hoja.get_Range("H4", string.Concat("H", Convert.ToString(4 + this.DG1.Rows.Count)))).FormatConditions.AddIconSetCondition();
                    iconSets.IconSet = libro.IconSets.get_Item(XlIconSet.xl4TrafficLights);
                    iconSets.IconCriteria[2].Type = XlConditionValueTypes.xlConditionValueNumber;
                    iconSets.IconCriteria[2].Value = 90;
                    iconSets.IconCriteria[2].Operator = 7;
                    iconSets.IconCriteria[3].Type = XlConditionValueTypes.xlConditionValueNumber;
                    iconSets.IconCriteria[3].Value = 100;
                    iconSets.IconCriteria[3].Operator = 7;
                    iconSets.IconCriteria[4].Type = XlConditionValueTypes.xlConditionValueNumber;
                    iconSets.IconCriteria[4].Value = 115;
                    iconSets.IconCriteria[4].Operator = 7;
                }
                Clipboard.Clear();
                objeto = null;
                hoja = (Worksheet)libro.Worksheets.get_Item(2);
                hoja.Activate();
                hoja.Name = "Indicador Detallado";
                hoja.Cells[2, 2] = "INDICADOR DETALLADO DE PRESUPUESTO";
                rango = (Range)hoja.get_Range("B2", "N2");
                rango.Select();
                rango.Merge();
                rango.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                rango.Borders.Weight = XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "BOLD";
                rango.Font.Name = "Consolas";
                rango.Interior.Color = Color.Gainsboro;
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@fecha", dia);
                if (cod_estab.Trim() != "1")
                {
                    comando.Parameters.AddWithValue("@estab", cod_estab);
                }
                else
                {
                    comando.Parameters.AddWithValue("@estab", "T");
                }
                comando.Parameters.AddWithValue("@tipo_reporte", 0);
                dt = new System.Data.DataTable();
                da.Fill(dt);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt;
                this.DG1.SelectAll();
                objeto = this.DG1.GetClipboardContent();
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    foreach (DataGridViewColumn dataGridViewColumn in this.DG1.Columns)
                    {
                        hoja.Cells[3, dataGridViewColumn.Index + 2] = dataGridViewColumn.Name.ToString().ToUpper();
                        rango = (Range)hoja.Cells[3, dataGridViewColumn.Index + 2];
                        rango.Borders.LineStyle = XlLineStyle.xlContinuous;
                        rango.Borders.Weight = XlBorderWeight.xlMedium;
                        rango.Font.Bold = true;
                        rango.Font.Name = "Consolas";
                        rango.WrapText = true;
                        rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        rango.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        if (dataGridViewColumn.Index <= 4)
                        {
                            continue;
                        }
                        rango.EntireColumn.NumberFormat = "#,##0.00";
                    }
                    rango = (Range)hoja.Cells[4, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    ((Range)hoja.get_Range("B4", string.Concat("N", Convert.ToString(5 + this.DG1.Rows.Count)))).Font.Name = "Consolas";
                    ((Range)hoja.get_Range("B2", string.Concat("N", Convert.ToString(5 + this.DG1.Rows.Count)))).BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                    if (cod_estab.Trim() != "1")
                    {
                        rango = (Range)hoja.get_Range("B3", string.Concat("N", Convert.ToString(2 + this.DG1.Rows.Count)));
                        rango.Select();
                        string[] strArrays1 = new string[] { "6", "7", "9", "11" };
                        rango.Subtotal(3, XlConsolidationFunction.xlSum, strArrays1, Type.Missing, Type.Missing, XlSummaryRow.xlSummaryBelow);
                        int count = this.DG1.Rows.Count + 3;
                        //while (true)
                        //{
                        //    if (hoja.Cells[string.Concat("D", count.ToString())] == null)
                        //    {
                        //        break;
                        //    }
                        //    count++;
                        //}
                        //count--;
                        rango = hoja.Range["B3", string.Concat("N", Convert.ToString(count))];
                        rango.Select();
                        rango.Subtotal(4, XlConsolidationFunction.xlSum, strArrays1, false, Type.Missing, XlSummaryRow.xlSummaryBelow);
                        hoja.Outline.ShowLevels(2, Type.Missing);
                    }
                    else
                    {
                        rango = (Range)hoja.get_Range("B3", string.Concat("N", Convert.ToString(2 + this.DG1.Rows.Count)));
                        rango.Select();
                        string[] strArrays2 = new string[] { "6", "7", "9", "11" };
                        rango.Subtotal(2, XlConsolidationFunction.xlSum, strArrays2, Type.Missing, Type.Missing, XlSummaryRow.xlSummaryBelow);
                        int num = this.DG1.Rows.Count + 3;
                        //while (true)
                        //{
                        //    //if (hoja.Cells[string.Concat("C", num.ToString())] == null)
                        //    //if ((dynamic)hoja.Cells[string.Concat("C", num.ToString()), Type.Missing][Type.Missing] == (dynamic)null)
                        //    if ((dynamic)hoja.Cells[string.Concat("C", num.ToString()), string.Concat("C", num.ToString())] == (dynamic)null)
                        //    {
                        //        break;
                        //    }
                        //    num++;
                        //}
                        //num--;
                        //rango = hoja.Cells["B3", string.Concat("N", Convert.ToString(num))];
                        rango = hoja.Range["B3", string.Concat("N", Convert.ToString(num))];
                        rango.Select();
                        rango.Subtotal(3, XlConsolidationFunction.xlSum, strArrays2, false, Type.Missing, XlSummaryRow.xlSummaryBelow);
                        //while (true)
                        //{
                        //    dynamic value = (dynamic)hoja.Cells[string.Concat("C", num.ToString())] != null;
                        //    if ((value ? value == null : (value | (dynamic)hoja.Cells[string.Concat("D", num.ToString())] != null) == 0))
                        //    {
                        //        break;
                        //    }
                        //    num++;
                        //}
                        //num--;
                        //rango = hoja.Cells["B3", string.Concat("N", Convert.ToString(num))];
                        rango = hoja.Range["B3", string.Concat("N", Convert.ToString(num))];
                        rango.Select();
                        rango.Subtotal(4, XlConsolidationFunction.xlSum, strArrays2, false, Type.Missing, XlSummaryRow.xlSummaryBelow);
                        hoja.Outline.ShowLevels(2, Type.Missing);
                    }
                    hoja.Range["B2", string.Concat("N", Convert.ToString(5 + this.DG1.Rows.Count))].EntireColumn.AutoFit();
                    libro.Worksheets[1].Activate();
                }
                Clipboard.Clear();
                objeto = null;
                if (File.Exists(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Indicador de presupuesto del estab ", estab.Trim(), ".xlsb")))
                {
                    File.Delete(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Indicador de presupuesto del estab ", estab.Trim(), ".xlsb"));
                }
                libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Indicador de presupuesto del estab " + estab.Trim() + ".xlsb", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12);
                libro.Close();
                excel.Quit();
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
            }
            return string.Concat(System.Windows.Forms.Application.StartupPath, "\\Indicador de presupuesto del estab ", estab.Trim(), ".xlsb");
        }

        private void Top80()
        {

            SqlConnection cn = conexion.conectar("BMSNayar");
            SqlCommand comando = new SqlCommand("select establecimientos.cod_estab,establecimientos.nombre,mail.email,mail.email_cordinador "
            + " from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab=mail.cod_estab"
            + " where establecimientos.status='V' and establecimientos.cod_estab not in ('1','1001','1002','1003','1004','1005','1006')"
            + " order by cast(establecimientos.cod_estab as int) asc", cn);
            SqlDataReader dr = comando.ExecuteReader();
            if (dr.HasRows)
            {
                DateTime dia = DateTime.Now;
                dia = dia.AddDays(-1);
                dia = dia.AddHours(dia.Hour * -1);
                dia = dia.AddMinutes(dia.Minute * -1);
                dia = dia.AddHours(23);
                dia = dia.AddMinutes(59);
                while (dr.Read())
                {
                    string destinatarios = dr["email"].ToString().Trim() + "," + dr["email_cordinador"].ToString().Trim();

                    if (dr["cod_estab"].ToString().Trim() == "1")
                    {
                        destinatarios = string.Concat(destinatarios, ",luis.guerrero@mercadodeimportaciones.com");
                    }
                    EnviaMailGmail(Top80Sucursal(dr["cod_estab"].ToString().Trim(), dr["nombre"].ToString(), dia), destinatarios);
                }
            }

        }

        private string Top80Sucursal(string cod_estab, string estab, DateTime FechaHora)
        {
            DateTime dia = FechaHora;
            DateTime primero = dia.AddDays((dia.Day - 1) * -1);
            primero = primero.AddHours(primero.Hour * -1);
            primero = primero.AddMinutes(primero.Minute * -1);
            SqlConnection cn = conexion.conectar("BMSNayar");
            SqlCommand comando = new SqlCommand();
            comando.Connection = cn;
            comando.CommandType = CommandType.StoredProcedure;
            comando.CommandText = "MI_Top80";
            comando.CommandTimeout = 240;
            Microsoft.Office.Interop.Excel.Application excel;
            excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Application.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook libro;
            libro = excel.Workbooks.Add();
            Worksheet hoja = new Worksheet();
            Microsoft.Office.Interop.Excel.Range rango;
            int nreportes = 1;
            if (cod_estab == "1") { nreportes = 8; } else { nreportes = 6; }
            for (int i = 1; i <= nreportes; i++)
            {
                comando.Parameters.Clear();
                libro.Worksheets.Add();
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                switch (i)
                {
                    case 1:
                        if (cod_estab == "1") { comando.Parameters.AddWithValue("@reporte", "FTG"); } else { comando.Parameters.AddWithValue("@reporte", "FT"); }
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia);
                        hoja.Name = "FALTANTE TRANSFERENCIA";
                        break;
                    case 2:
                        if (cod_estab == "1") { comando.Parameters.AddWithValue("@reporte", "AIFDG"); } else { comando.Parameters.AddWithValue("@reporte", "AIFD"); }
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia);
                        hoja.Name = "AJUSTE DISMINUCION";
                        break;
                    case 3:
                        if (cod_estab == "1") { comando.Parameters.AddWithValue("@reporte", "AIFAG"); } else { comando.Parameters.AddWithValue("@reporte", "AIFA"); }
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia);
                        hoja.Name = "AJUSTE AUMENTO";
                        break;
                    case 4:
                        if (cod_estab == "1") { comando.Parameters.AddWithValue("@reporte", "MERG"); } else { comando.Parameters.AddWithValue("@reporte", "MER"); }
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia);
                        hoja.Name = "MERMAS";
                        break;
                    case 5:
                        if (cod_estab == "1") { comando.Parameters.AddWithValue("@reporte", "CGP"); } else { comando.Parameters.AddWithValue("@reporte", "CE"); }
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia); ;
                        hoja.Name = "CONTRIBUCION X ARTICULO";
                        break;
                    case 6:
                        if (cod_estab == "1") { comando.Parameters.AddWithValue("@reporte", "VGP"); } else { comando.Parameters.AddWithValue("@reporte", "VPE"); }
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia);
                        hoja.Name = "VENTAS X ARTICULO";
                        break;
                    case 7:
                        comando.Parameters.AddWithValue("@reporte", "CGE");
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia);
                        hoja.Name = "CONTRIBUCION X ESTAB";
                        break;
                    case 8:
                        comando.Parameters.AddWithValue("@reporte", "VGE");
                        comando.Parameters.AddWithValue("@estab", cod_estab.Trim());
                        comando.Parameters.AddWithValue("@fini", primero);
                        comando.Parameters.AddWithValue("@ffin", dia);
                        hoja.Name = "VENTAS X ESTAB";
                        break;
                }
                SqlDataAdapter da = new SqlDataAdapter(comando);
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                DG1.DataSource = dt;
                DG1.SelectAll();
                object objeto = DG1.GetClipboardContent();

                try
                {

                    if (objeto != null)
                    {
                        Clipboard.SetDataObject(objeto);
                        foreach (DataGridViewColumn columna in DG1.Columns)
                        {
                            hoja.Cells[1, columna.Index + 2] = columna.Name.ToString().ToUpper();
                        }
                        rango = (Range)hoja.Cells[2, 1];
                        rango.Select();
                        hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    }
                    Clipboard.Clear();
                    objeto = null;
                }
                catch (Exception e)
                {
                    // Si no se hace nada cuando pasa un error sino que solo se muestra en la etiqueta informativa, el proceso de envío del reporte continúa...
                    //MessageBox.Show(e.Message);
                    lblEstado.Text = e.Message;
                }

            }


            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Top80 del estab " + estab.Trim() + ".xlsb"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Top80 del estab " + estab.Trim() + ".xlsb");
            }

            libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Top80 del estab " + estab.Trim() + ".xlsb", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12);
            libro.Close();
            excel.Quit();
            if (cn.State == ConnectionState.Open) { cn.Close(); }
            return System.Windows.Forms.Application.StartupPath + "\\Top80 del estab " + estab.Trim() + ".xlsb";
        }

        private void Vigencias()
        {
            SqlConnection cn = conexion.conectar("BMSNayar");
            SqlCommand comando = new SqlCommand("select establecimientos.cod_estab,establecimientos.nombre,mail.email,mail.email_cordinador "
            + " from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab=mail.cod_estab"
            + " where establecimientos.status='V' and establecimientos.cod_estab not in ('1','1001','1002','1003','1004','1005','1006')"
            + " order by cast(establecimientos.cod_estab as int) asc", cn);
            SqlDataReader dr = comando.ExecuteReader();
            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    EnviaMailGmail(VigenciaSucursal(dr["cod_estab"].ToString(), dr["nombre"].ToString().Trim()), dr["email"].ToString() + ',' + dr["email_cordinador"].ToString());
                    //string archivo = VigenciaSucursal(dr["cod_estab"].ToString(),dr["nombre"].ToString().Trim());
                }
            }
            if (dr.IsClosed == false) { dr.Close(); }
            if (cn.State == ConnectionState.Open) { cn.Close(); }
        }

        private string VigenciaSucursal(string cod_estab, string estab)
        {
            SqlConnection cn = conexion.conectar("BMSNayar");
            SqlCommand comando = new SqlCommand();
            comando.Connection = cn;
            comando.CommandType = CommandType.StoredProcedure;
            comando.CommandText = "MI_Vigencias";

            Microsoft.Office.Interop.Excel.Application excel;
            excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Application.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook libro;
            libro = excel.Workbooks.Add();
            Worksheet hoja = new Worksheet();
            Microsoft.Office.Interop.Excel.Range rango;

            for (int i = 0; i < 4; i++)
            {
                comando.Parameters.Clear();
                libro.Worksheets.Add();
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                switch (i)
                {
                    case 0:
                        comando.Parameters.AddWithValue("@cod_estab", cod_estab);
                        comando.Parameters.AddWithValue("@vigencia", "0");
                        comando.Parameters.AddWithValue("@aplica", 1);
                        hoja.Name = "SIN VIGENCIA";
                        break;
                    case 1:
                        comando.Parameters.AddWithValue("@cod_estab", cod_estab);
                        comando.Parameters.AddWithValue("@vigencia", "1");
                        comando.Parameters.AddWithValue("@aplica", 1);
                        hoja.Name = "VIGENCIA 1";
                        break;
                    case 2:
                        comando.Parameters.AddWithValue("@cod_estab", cod_estab);
                        comando.Parameters.AddWithValue("@vigencia", "2");
                        comando.Parameters.AddWithValue("@aplica", 1);
                        hoja.Name = "VIGENCIA 2";
                        break;
                    case 3:
                        comando.Parameters.AddWithValue("@cod_estab", cod_estab);
                        comando.Parameters.AddWithValue("@vigencia", "3");
                        comando.Parameters.AddWithValue("@aplica", 1);
                        hoja.Name = "VIGENCIA 3";
                        break;
                }
                SqlDataAdapter da = new SqlDataAdapter(comando);
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();
                DG1.DataSource = dt;
                DG1.SelectAll();
                object objeto = DG1.GetClipboardContent();

                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    foreach (DataGridViewColumn columna in DG1.Columns)
                    {
                        hoja.Cells[1, columna.Index + 2] = columna.Name.ToString().ToUpper();
                    }
                    rango = (Range)hoja.get_Range("A1", "H1");
                    rango.EntireColumn.NumberFormat = "@";
                    rango = (Range)hoja.get_Range("L1", "O1");
                    rango.EntireColumn.NumberFormat = "###,##0.00";
                    rango = (Range)hoja.Cells[2, 1];
                    rango.Select();
                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                }
                Clipboard.Clear();
                objeto = null;

            }
            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Vigencias del estab " + estab.Trim() + ".xlsb"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Vigencias del estab " + estab.Trim() + ".xlsb");
            }

            libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Vigencias del estab " + estab.Trim() + ".xlsb", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12);
            libro.Close();
            excel.Quit();
            if (cn.State == ConnectionState.Open) { cn.Close(); }
            return System.Windows.Forms.Application.StartupPath + "\\Vigencias del estab " + estab.Trim() + ".xlsb";

        }

        private string ComparativoPresupuesto(DateTime FechaHora)
        {
            Range range;
            string str;
            int num;
            try
            {
                SqlConnection cn = conexion.conectar("BMSNayar");
                SqlCommand comando = new SqlCommand();
                comando.Connection = cn;
                comando.CommandTimeout = 240;
                System.Data.DataTable dt = new System.Data.DataTable();
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = comando;

                DateTime dia = FechaHora;
                DateTime primero = dia.AddDays((double)((dia.Day - 1) * -1));
                dia.AddYears(-1);
                primero.AddYears(-1);

                Microsoft.Office.Interop.Excel.Application excel;
                excel = new Microsoft.Office.Interop.Excel.Application();
                //excel.Application.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook libro;
                libro = excel.Workbooks.Add();
                Worksheet hoja = new Worksheet();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();
                libro.Worksheets.Add();

                comando.CommandText = "delete from _TempVentas";
                comando.ExecuteNonQuery();
                string[] strArrays = new string[] { "insert into _TempVentas select Sucursales.cod_estab,Sucursales.nombre,isnull(HogarDiario.unidades,0) as HogarDiarioUni,isnull(HogarAcu.unidades,0) as HogarAcuUni,cast(0.0 as float) as '%1',isnull(HogarDiario.Total,0) as HogarDiarioVta,isnull(HogarAcu.Total,0) as HogarAcuVta,cast(0.0 as float) as '%2',isnull(HogarDiario.UtilBruta,0) as HogarDiarioUtil,isnull(HogarAcu.UtilBruta,0) as HogarAcuUtil,cast(0.0 as float) as '%3',isnull(BisuteriaDiario.unidades,0) as BisuteriaDiarioUni,isnull(BisuteriaAcu.unidades,0) as BisuteriaAcuUni,cast(0.0 as float) as '%4',isnull(BisuteriaDiario.Total,0) as BisuteriaDiarioVta,isnull(BisuteriaAcu.Total,0) as BisuteriaAcuVta,cast(0.0 as float) as '%5',isnull(BisuteriaDiario.UtilBruta,0) as BisuteriaDiarioUtil,isnull(BisuteriaAcu.UtilBruta,0) as BisuteriaAcuUtil,cast(0.0 as float) as '%6',isnull(BellezaDiario.unidades,0) as BellezaDiarioUni,isnull(BellezaAcu.unidades,0) as BellezaAcuUni,cast(0.0 as float) as '%7',isnull(BellezaDiario.Total,0) as BellezaDiarioVta,isnull(BellezaAcu.Total,0) as BellezaAcuVta,cast(0.0 as float) as '%8',isnull(BellezaDiario.UtilBruta,0) as BellezaDiarioUtil,isnull(BellezaAcu.UtilBruta,0) as BellezaAcuUtil,cast(0.0 as float) as '%9',isnull(CalzadoDiario.unidades,0) as CalzadoDiarioUni,isnull(CalzadoAcu.unidades,0) as CalzadoAcuUni,cast(0.0 as float) as '%10',isnull(CalzadoDiario.Total,0) as CalzadoDiarioVta,isnull(CalzadoAcu.Total,0) as CalzadoAcuVta,cast(0.0 as float) as '%11',isnull(CalzadoDiario.UtilBruta,0) as CalzadoDiarioUtil,isnull(CalzadoAcu.UtilBruta,0) as CalzadoAcuUtil,cast(0.0 as float) as '%12',isnull(RopaDiario.unidades,0) as RopaDiarioUni,isnull(RopaAcu.unidades,0) as RopaAcuUni,cast(0.0 as float) as '%13',isnull(RopaDiario.Total,0) as RopaDiarioVta,isnull(RopaAcu.Total,0) as RopaAcuVta,cast(0.0 as float) as '%14',isnull(RopaDiario.UtilBruta,0) as RopaDiarioUtil,isnull(RopaAcu.UtilBruta,0) as RopaAcuUtil,cast(0.0 as float) as '%15',isnull(ServiciosDiario.unidades,0) as ServiciosDiarioUni,isnull(ServiciosAcu.unidades,0) as ServiciosAcuUni,cast(0.0 as float) as '%16',isnull(ServiciosDiario.Total,0) as ServiciosDiarioVta,isnull(ServiciosAcu.Total,0) as ServiciosAcuVta,cast(0.0 as float) as '%17',isnull(ServiciosDiario.UtilBruta,0) as ServiciosDiarioUtil,isnull(ServiciosAcu.UtilBruta,0) as ServiciosAcuUtil,cast(0.0 as float) as '%18',/*isnull(BotanasDiario.unidades,0) as BotanasDiarioUni,isnull(BotanasAcu.unidades,0) as BotanasAcuUni,cast(0.0 as float) as '%19',isnull(BotanasDiario.Total,0) as BotanasDiarioVta,isnull(BotanasAcu.Total,0) as BotanasAcuVta,cast(0.0 as float) as '%20',isnull(BotanasDiario.UtilBruta,0) as BotanasDiarioUtil,isnull(BotanasAcu.UtilBruta,0) as BotanasAcuUtil,cast(0.0 as float) as '%21',*/isnull(AbarrotesDiario.unidades,0) as AbarrotesDiarioUni,isnull(AbarrotesAcu.unidades,0) as AbarrotesAcuUni,cast(0.0 as float) as '%22',isnull(AbarrotesDiario.Total,0) as AbarrotesDiarioVta,isnull(AbarrotesAcu.Total,0) as AbarrotesAcuVta,cast(0.0 as float) as '%23',isnull(AbarrotesDiario.UtilBruta,0) as AbarrotesDiarioUtil,isnull(AbarrotesAcu.UtilBruta,0) as AbarrotesAcuUtil,cast(0.0 as float) as '%24',isnull(TemporadaDiario.unidades,0) as TemporadaDiarioUni,isnull(TemporadaAcu.unidades,0) as TemporadaAcuUni,cast(0.0 as float) as '%25',isnull(TemporadaDiario.Total,0) as TemporadaDiarioVta,isnull(TemporadaAcu.Total,0) as TemporadaAcuVta,cast(0.0 as float) as '%26',isnull(TemporadaDiario.UtilBruta,0) as TemporadaDiarioUtil,isnull(TemporadaAcu.UtilBruta,0) as TemporadaAcuUtil,cast(0.0 as float) as '%27',isnull(TotalDiario.unidades,0) as TotalDiarioUni,isnull(TotalAcu.unidades,0) as TotalAcuUni,Cast(0.0 as float) as '%28',isnull(TotalDiario.Total,0) as TotalDiarioVta,isnull(TotalAcu.Total,0) as TotalAcuVta,cast(0.0 as float) as '%29',isnull(TotalDiario.UtilBruta,0) as TotalDiarioUtil,isnull(TotalAcu.UtilBruta,0) as TotalAcuUtil,cast(0.0 as float) as '%30'  from (((((((((((((((select cod_estab,nombre from establecimientos where status='V' and cod_estab not in ('1','1001','1002','1003','1004','1005','1006','67','65', '55')) as Sucursales left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock) inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as BisuteriaDiario on Sucursales.cod_estab=BisuteriaDiario.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='1' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as BisuteriaAcu on Sucursales.cod_estab=BisuteriaAcu.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total,SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as RopaDiario on Sucursales.cod_estab=RopaDiario.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='2' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as RopaAcu on Sucursales.cod_estab=RopaAcu.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as CalzadoDiario on Sucursales.cod_estab=CalzadoDiario.cod_estab) left join  (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='3' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as CalzadoAcu on Sucursales.cod_estab=CalzadoAcu.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as HogarDiario on Sucursales.cod_estab=HogarDiario.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='4' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as HogarAcu on Sucursales.cod_estab=HogarAcu.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as ServiciosDiario on Sucursales.cod_estab=ServiciosDiario.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='6' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as ServiciosAcu on Sucursales.cod_estab=ServiciosAcu.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as BotanasDiario on Sucursales.cod_estab=BotanasDiario.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='7' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as BotanasAcu on Sucursales.cod_estab=BotanasAcu.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as BellezaDiario on Sucursales.cod_estab=BellezaDiario.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='8' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as BellezaAcu on Sucursales.cod_estab=BellezaAcu.cod_estab left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='9'  and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as AbarrotesDiario on Sucursales.cod_estab=AbarrotesDiario.cod_estab left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='9'  and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as AbarrotesAcu on Sucursales.cod_estab=AbarrotesAcu.cod_estab left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='10' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as TemporadaDiario on Sucursales.cod_estab=TemporadaDiario.cod_estab left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto='10' and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as TemporadaAcu on Sucursales.cod_estab=TemporadaAcu.cod_estab left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto in ('1','2','3','4','6','7','8','9','10') and fecha between  '", dia.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as TotalDiario on Sucursales.cod_estab=TotalDiario.cod_estab) left join (select cod_estab,sum(case when transaccion in ('36','37','38') then cantidad  else cantidad*-1 end) as unidades,SUM(case when transaccion in  ('36','37','38') then entysalVentas.total else entysalVentas.total*-1 end) as Total, SUM(case when transaccion in ('36','37','38') then entysalVentas.utilidad_bruta else entysalVentas.utilidad_bruta*-1 end) as UtilBruta from (entysalVentas with(nolock)inner join productos as p with(nolock) on entysalVentas.cod_prod=p.cod_prod ) where p.linea_producto in ('1','2','3','4','6','7','8','9','10') and entysalVentas.cod_cte not in (select cod_cte from clientes where clasificacion_cliente='2') and transaccion in ('36','37','38','308','34','68') and fecha between '", primero.ToString("yyyyMMdd"), "' and '", dia.ToString("yyyyMMdd HH:mm"), "' group by entysalVentas.cod_estab) as TotalAcu on Sucursales.cod_estab=TotalAcu.cod_estab order by CAST(Sucursales.cod_estab as int) asc" };
                comando.CommandText = string.Concat(strArrays);
                comando.ExecuteNonQuery();
                comando.CommandText = "update _TempVentas set [%1]=(HogarDiarioUni/(case when (select sum(HogarDiarioUni) from _TempVentas)<>0 then (select sum(HogarDiarioUni) from _TempVentas) else null end))*100,[%2]=(HogarDiarioVta/(case when (select sum(HogarDiarioVta) from _TempVentas)<>0 then (select sum(HogarDiarioVta) from _TempVentas) else null end))*100,[%3]=(HogarDiarioUtil/(case when (select sum(HogarDiarioUtil) from _TempVentas)<>0 then (select sum(HogarDiarioUtil) from _TempVentas) else null  end))*100,[%4]=(BisuteriaDiarioUni/(case when (select sum(BisuteriaDiarioUni) from _TempVentas)<>0 then (select sum(BisuteriaDiarioUni) from _TempVentas) else null end))*100,[%5]=(BisuteriaDiarioVta/(case when (select sum(BisuteriaDiarioVta) from _TempVentas)<>0 then (select sum(BisuteriaDiarioVta) from _TempVentas) else null end))*100,[%6]=(BisuteriaDiarioUtil/(case when (select sum(BisuteriaDiarioUtil) from _TempVentas)<>0 then (select sum(BisuteriaDiarioUtil) from _TempVentas) else null end))*100,[%7]=(BellezaDiarioUni/(case when (select sum(BellezaDiarioUni) from _TempVentas)<>0 then (select sum(BellezaDiarioUni) from _TempVentas) else null end))*100,[%8]=(BellezaDiarioVta/(case when (select sum(BellezaDiarioVta) from _TempVentas)<>0 then (select sum(BellezaDiarioVta) from _TempVentas) else null end))*100,[%9]=(BellezaDiarioUtil/(case when (select sum(BellezaDiarioUtil) from _TempVentas)<>0 then (select sum(BellezaDiarioUtil) from _TempVentas) else null end))*100,[%10]=(CalzadoDiarioUni/(case when (select sum(CalzadoDiarioUni) from _TempVentas)<>0 then (select sum(CalzadoDiarioUni) from _TempVentas) else null end))*100,[%11]=(CalzadoDiarioVta/(case when (select sum(CalzadoDiarioVta) from _TempVentas)<>0 then (select sum(CalzadoDiarioVta) from _TempVentas) else null end))*100,[%12]=(CalzadoDiarioUtil/(case when (select sum(CalzadoDiarioUtil) from _TempVentas)<>0 then (select sum(CalzadoDiarioUtil) from _TempVentas) else null end))*100,[%13]=(RopaDiarioUni/(case when (select sum(RopaDiarioUni) from _TempVentas)<>0 then (select sum(RopaDiarioUni) from _TempVentas) else null end))*100,[%14]=(RopaDiarioVta/(case when (select sum(RopaDiarioVta) from _TempVentas)<>0 then (select sum(RopaDiarioVta) from _TempVentas) else null end))*100,[%15]=(RopaDiarioUtil/(case when (select sum(RopaDiarioUtil) from _TempVentas)<>0 then (select sum(RopaDiarioUtil) from _TempVentas) else null end))*100,[%16]=(ServiciosDiarioUni/(case when (select sum(ServiciosDiarioUni) from _TempVentas)<>0 then (select sum(ServiciosDiarioUni) from _TempVentas) else null end))*100,[%17]=(ServiciosDiarioVta/(case when (select sum(ServiciosDiarioVta) from _TempVentas)<>0 then (select sum(ServiciosDiarioVta) from _TempVentas) else null end))*100,[%18]=(ServiciosDiarioUtil/(case when (select sum(ServiciosDiarioUtil) from _TempVentas)<>0 then (select sum(ServiciosDiarioUtil) from _TempVentas) else null end))*100,/*[%19]=(BotanasDiarioUni/(case when (select sum(BotanasDiarioUni) from _TempVentas)<>0 then (select sum(BotanasDiarioUni) from _TempVentas) else null end))*100,[%20]=(BotanasDiarioVta/(case when (select sum(BotanasDiarioVta) from _TempVentas)<>0 then (select sum(BotanasDiarioVta) from _TempVentas) else null end))*100,[%21]=(BotanasDiarioUtil/(case when (select sum(BotanasDiarioUtil) from _TempVentas)<>0 then (select sum(BotanasDiarioUtil) from _TempVentas) else null end))*100,*/[%22]=(AbarrotesDiarioUni/(case when (select sum(AbarrotesDiarioUni) from _TempVentas)<>0 then (select sum(AbarrotesDiarioUni) from _TempVentas) else null end))*100,[%23]=(AbarrotesDiarioVta/(case when (select sum(AbarrotesDiarioVta) from _TempVentas)<>0 then (select sum(AbarrotesDiarioVta) from _TempVentas) else null end))*100,[%24]=(AbarrotesDiarioUtil/(case when (select sum(AbarrotesDiarioUtil) from _TempVentas)<>0 then (select sum(AbarrotesDiarioUtil) from _TempVentas) else null end))*100,[%25]=(TemporadaDiarioUni/(case when (select sum(TemporadaDiarioUni) from _TempVentas)<>0 then (select sum(TemporadaDiarioUni) from _TempVentas) else null end))*100,[%26]=(TemporadaDiarioVta/(case when (select sum(TemporadaDiarioVta) from _TempVentas)<>0 then (select sum(TemporadaDiarioVta) from _TempVentas) else null end))*100,[%27]=(TemporadaDiarioUtil/(case when (select sum(TemporadaDiarioUtil) from _TempVentas)<>0 then (select sum(TemporadaDiarioUtil) from _TempVentas) else null end))*100,[%28]=(TotalDiarioUni/(case when (select sum(TotalDiarioUni) from _TempVentas)<>0 then (select sum(TotalDiarioUni) from _TempVentas) else null end))*100,[%29]=(TotalDiarioVta/(case when (select sum(TotalDiarioVta) from _TempVentas)<>0 then  (select sum(TotalDiarioVta) from _TempVentas) else null end))*100,[%30]=(TotalDiarioUtil/(case when (select sum(TotalDiarioUtil) from _TempVentas)<>0 then (select sum(TotalDiarioUtil) from _TempVentas) else null end))*100";
                comando.ExecuteNonQuery();
                // 22/Nov/2019 - Eduardo Pérez Jr. solicitó que los establecimientos se ordenaran desde mazatlán hasta los cabos colocando primero los establecimientos del centro de cada ciudad y después los que estuvieran en las plazas o periferia
                //comando.CommandText = "select _TempVentas.* from _TempVentas inner join establecimientos on _TempVentas.cod_estab=establecimientos.cod_estab order by establecimientos.tipo_establecimiento desc,establecimientos.grupo_establecimiento asc,cast(establecimientos.cod_estab as int)";
                comando.CommandText = "select _TempVentas.* from _TempVentas inner join establecimientos on _TempVentas.cod_estab=establecimientos.cod_estab left join _MI_OrdenEstab_Comparativos oec on _TempVentas.cod_estab = oec.cod_estab order by oec.orden";
                this.DG1.DataSource = null;

                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                da.Fill(dt);
                this.DG1.DataSource = dt;
                this.DG1.SelectAll();
                object objeto = this.DG1.GetClipboardContent();
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                hoja.Name = "VENTA DIARIA";
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja.Cells[1, 2] = "REPORTE DE VENTA DIARIA";
                    range = hoja.Range["B1", "BW1"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B3", "B5"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "CODIGO";
                    range = hoja.Range["C3", "C5"];
                    range.Select();
                    range.Merge();
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "SUCURSAL";
                    for (int i = 4; i <= 84; i += 9)
                    {
                        range = (Range)hoja.get_Range(string.Concat(this.sCol(i), "3"), string.Concat(this.sCol(i + 8), "3"));
                        range.Select();
                        range.Merge();
                        range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        range.Borders.LineStyle = XlLineStyle.xlContinuous;
                        range.Borders.Weight = XlBorderWeight.xlMedium;
                        range.Cells.Font.FontStyle = "Bold";
                        num = i;
                        if (num <= 31)
                        {
                            if (num <= 13)
                            {
                                if (num == 4)
                                {
                                    range.Cells[1, 1] = " H O G A R";
                                }
                                else if (num == 13)
                                {
                                    range.Cells[1, 1] = "B I S U T E R I A";
                                }
                            }
                            else if (num == 22)
                            {
                                range.Cells[1, 1] = "B E L L E Z A";
                            }
                            else if (num == 31)
                            {
                                range.Cells[1, 1] = "C A L Z A D O";
                            }
                        }
                        else if (num <= 49)
                        {
                            if (num == 40)
                            {
                                range.Cells[1, 1] = "R O P A";
                            }
                            else if (num == 49)
                            {
                                range.Cells[1, 1] = "S E R V I C I O S";
                            }
                        }
                        /*else if (num == 58)
                        {
                            range.Cells[1, 1] = "B O T A N A S   Y   S N A C K S";
                        }*/
                        else if (num == 58)
                        {
                            range.Cells[1, 1] = "A B A R R O T E S";
                        }
                        else if (num == 67)
                        {
                            range.Cells[1, 1] = "T E M P O R A D A";
                        }
                        else if (num == 76)
                        {
                            range.Cells[1, 1] = "T O T A L   D I A R I O";
                        }
                        for (int j = i; j <= i + 8; j += 3)
                        {
                            range = (Range)hoja.get_Range(string.Concat(this.sCol(j), "4"), string.Concat(this.sCol(j + 2), "4"));
                            range.Select();
                            range.Merge();
                            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                            range.Borders.LineStyle = XlLineStyle.xlContinuous;
                            range.Borders.Weight = XlBorderWeight.xlMedium;
                            range.Cells.Font.FontStyle = "Bold";
                            if (j == i)
                            {
                                range.Cells[1, 1] = "UNIDADES";
                            }
                            else if (j == i + 3)
                            {
                                range.Cells[1, 1] = "VENTA NETA";
                            }
                            else if (j == i + 6)
                            {
                                range.Cells[1, 1] = "CONTRIBUCION";
                            }
                            for (int k = j; k <= j + 2; k++)
                            {
                                if (k == j)
                                {
                                    hoja.Cells[5, k] = "DIARIAS";
                                    if (j != i)
                                    {
                                        hoja.Range[string.Concat(this.sCol(k), "5"), string.Concat(this.sCol(k), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                    }
                                    else
                                    {
                                        hoja.Range[string.Concat(this.sCol(k), "5"), string.Concat(this.sCol(k), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0";
                                    }
                                }
                                else if (k == j + 1)
                                {
                                    hoja.Cells[5, k] = "ACUMULADAS";
                                    if (j != i)
                                    {
                                        hoja.Range[string.Concat(this.sCol(k), "5"), string.Concat(this.sCol(k), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                    }
                                    else
                                    {
                                        hoja.Range[string.Concat(this.sCol(k), "5"), string.Concat(this.sCol(k), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0";
                                    }
                                }
                                else if (k == j + 2)
                                {
                                    hoja.Cells[5, k] = "%";
                                    hoja.Range[string.Concat(this.sCol(k), "5"), string.Concat(this.sCol(k), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                }
                                hoja.Cells[5, k].Borders.LineStyle = 1;
                                hoja.Cells[5, k].Borders.Weight = -4138;
                                hoja.Cells[5, k].Font.FontStyle = "Bold";
                                dynamic cells = hoja.Cells[this.DG1.Rows.Count + 6, k];
                                string[] strArrays1 = new string[] { "=SUM(", this.sCol(k), "5:", this.sCol(k), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                cells.Formula = string.Concat(strArrays1);
                            }
                        }
                    }
                    hoja.Range["B6", string.Concat("C", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", string.Concat("CF", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignRight;
                    hoja.Range["B6", string.Concat("CF", Convert.ToString(this.DG1.Rows.Count + 6))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                    hoja.Cells[this.DG1.Rows.Count + 6, 3] = "T O T A L";
                    range = (Range)hoja.Cells[6, 1];
                    range.Select();
                    hoja.PasteSpecial(range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    range = hoja.Range["A1", string.Concat("BW", Convert.ToString(this.DG1.Rows.Count + 10))];
                    range.EntireColumn.AutoFit();
                }
                Clipboard.Clear();
                objeto = null;
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                comando.CommandType = CommandType.StoredProcedure;
                comando.CommandText = "MI_ComparativoPresupuesto";
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@fechaInicial", Convert.ToDateTime(dia.ToString("yyyy-MM-dd")));
                comando.Parameters.AddWithValue("@fechaFinal", dia);
                comando.Parameters.AddWithValue("@tipo", 1);
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter();
                da.SelectCommand = comando;
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                this.DG1.DataSource = dt;
                this.DG1.SelectAll();
                objeto = this.DG1.GetClipboardContent();
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Sheets[2];
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO DIARIO";
                    hoja.Cells[1, 2] = "COMPARATIVO DIARIO";
                    range = hoja.Range["B1", "AL1"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B2", "AL2"];
                    range.Select();
                    range.Merge();
                    range.Cells[1.1, Type.Missing] = dia.ToString("d MMM yyyy");
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B3", "B5"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "CODIGO";
                    range = hoja.Range["C3", "C5"];
                    range.Select();
                    range.Merge();
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "SUCURSAL";
                    for (int l = 4; l <= 39; l += 5)
                    {
                        range = hoja.Range[string.Concat(this.sCol(l), "3"), string.Concat(this.sCol(l + 4), "3")];
                        range.Select();
                        range.Merge();
                        range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        range.Borders.LineStyle = XlLineStyle.xlContinuous;
                        range.Borders.Weight = XlBorderWeight.xlMedium;
                        range.Cells.Font.FontStyle = "Bold";
                        int num1 = l;
                        if (num1 <= 14)
                        {
                            if (num1 == 4)
                            {
                                range.Cells[1, 1] = "COMPARATIVO DIARIO HOGAR";
                            }
                            else if (num1 == 9)
                            {
                                range.Cells[1, 1] = "COMPARATIVO DIARIO BISUTERIA";
                            }
                            else if (num1 == 14)
                            {
                                range.Cells[1, 1] = "COMPARATIVO DIARIO BELLEZA";
                            }
                        }
                        else if (num1 <= 24)
                        {
                            if (num1 == 19)
                            {
                                range.Cells[1, 1] = "COMPARATIVO DIARIO CALZADO";
                            }
                            else if (num1 == 24)
                            {
                                range.Cells[1, 1] = "COMPARATIVO DIARIO ROPA";
                            }
                        }
                        else if (num1 == 29)
                        {
                            range.Cells[1, 1] = "COMPARATIVO DIARIO SERVICIOS";
                        }
                        /*else if (num1 == 34)
                        {
                            range.Cells[1, 1] = "COMPARATIVO DIARIO SNACKS Y BOTANAS";
                        }*/
                        else if (num1 == 34)
                        {
                            range.Cells[1, 1] = "COMPARATIVO DIARIO ABARROTES";
                        }
                        else if (num1 == 39)
                        {
                            range.Cells[1, 1] = "COMPARATIVO DIARIO TEMPORADA";
                        }
                        for (int m = l; m <= l + 4; m++)
                        {
                            if (m <= l)
                            {
                                range = hoja.Range[string.Concat(this.sCol(m), "4"), string.Concat(this.sCol(m + 1), "4")];
                                range.Select();
                                range.Merge();
                                range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                                range.Borders.Weight = XlBorderWeight.xlMedium;
                                range.Cells.Font.FontStyle = "Bold";
                                range.Cells[1, 1] = "VENTA NETA";
                            }
                            else
                            {
                                hoja.Cells[4, m] = "%";
                                hoja.Cells[4, m].HorizontalAlignment = -4108;
                                hoja.Cells[4, m].VerticalAlignment = -4108;
                                hoja.Cells[4, m].Borders.LineStyle = 1;
                                hoja.Cells[4, m].Borders.Weight = -4138;
                            }
                            if (m == l)
                            {
                                hoja.Cells[5, m] = "OBJETIVO";
                                hoja.Cells[5, m].HorizontalAlignment = -4108;
                                hoja.Cells[5, m].VerticalAlignment = -4108;
                                hoja.Cells[5, m].Borders.LineStyle = 1;
                                hoja.Cells[5, m].Borders.Weight = -4138;
                                hoja.Range[string.Concat(this.sCol(m), "6"), string.Concat(this.sCol(m), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic obj = hoja.Cells[this.DG1.Rows.Count + 6, m];
                                string[] strArrays2 = new string[] { "=SUM(", this.sCol(m), "6:", this.sCol(m), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                obj.Formula = string.Concat(strArrays2);
                            }
                            else if (m == l + 1)
                            {
                                hoja.Cells[5, m].HorizontalAlignment = -4108;
                                hoja.Cells[5, m].NumberFormat = "@";
                                hoja.Cells[5, m] = "REALIZADO";
                                hoja.Range[string.Concat(this.sCol(m), "6"), string.Concat(this.sCol(m), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic cells1 = hoja.Cells[this.DG1.Rows.Count + 6, m];
                                string[] strArrays3 = new string[] { "=SUM(", this.sCol(m), "6:", this.sCol(m), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                cells1.Formula = string.Concat(strArrays3);
                            }
                            else if (m == l + 2)
                            {
                                hoja.Cells[5, m].HorizontalAlignment = -4108;
                                hoja.Cells[5, m].NumberFormat = "@";
                                hoja.Cells[5, m] = "CUMPL";
                                hoja.Range[string.Concat(this.sCol(m), "6"), string.Concat(this.sCol(m), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic obj1 = hoja.Cells[this.DG1.Rows.Count + 6, m];
                                string[] strArrays4 = new string[] { "=(", this.sCol(m - 1), Convert.ToString(this.DG1.Rows.Count + 6), "/", this.sCol(m - 2), Convert.ToString(this.DG1.Rows.Count + 6), ")*100" };
                                obj1.Formula = string.Concat(strArrays4);
                            }
                            else if (m == l + 3)
                            {
                                hoja.Cells[5, m].HorizontalAlignment = -4108;
                                hoja.Cells[5, m].NumberFormat = "@";
                                hoja.Cells[5, m] = "COSTO";
                                hoja.Range[string.Concat(this.sCol(m), "6"), string.Concat(this.sCol(m), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            else if (m == l + 4)
                            {
                                hoja.Cells[5, m].HorizontalAlignment = -4108;
                                hoja.Cells[5, m].NumberFormat = "@";
                                hoja.Cells[5, m] = "MARGEN";
                                hoja.Range[string.Concat(this.sCol(m), "6"), string.Concat(this.sCol(m), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            hoja.Cells[5, m].Borders.LineStyle = 1;
                            hoja.Cells[5, m].Borders.Weight = -4138;
                            hoja.Cells[5, m].Font.FontStyle = "Bold";
                        }
                    }
                    hoja.Range["B6", string.Concat("C", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", string.Concat("AQ", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignRight;
                    hoja.Range["B6", string.Concat("AQ", Convert.ToString(this.DG1.Rows.Count + 6))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                    hoja.Cells[this.DG1.Rows.Count + 6, 3] = "T O T A L";
                    range = (Range)hoja.Cells[6, 1];
                    range.Select();
                    hoja.PasteSpecial(range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    range = hoja.Range["A1", string.Concat("AL", Convert.ToString(this.DG1.Rows.Count + 10))];
                    range.EntireColumn.AutoFit();
                }
                Clipboard.Clear();
                objeto = null;
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@fechaInicial", Convert.ToDateTime(primero.ToString("yyyy-MM-dd")));
                comando.Parameters.AddWithValue("@fechaFinal", dia);
                comando.Parameters.AddWithValue("@tipo", 1);
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter()
                {
                    SelectCommand = comando
                };
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                this.DG1.DataSource = dt;
                this.DG1.SelectAll();
                objeto = this.DG1.GetClipboardContent();
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Worksheets.get_Item(3);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO ACUMULADO";
                    hoja.Cells[1, 2] = "COMPARATIVO ACUMULADO";
                    range = hoja.Range["B1", "AL1"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B2", "AL2"];
                    range.Select();
                    range.Merge();
                    range.Cells[1.1, Type.Missing] = dia.ToString("d MMM yyyy");
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B3", "B5"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "CODIGO";
                    range = hoja.Range["C3", "C5"];
                    range.Select();
                    range.Merge();
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "SUCURSAL";
                    for (int n = 4; n <= 39; n += 5)
                    {
                        range = hoja.Range[string.Concat(this.sCol(n), "3"), string.Concat(this.sCol(n + 4), "3")];
                        range.Select();
                        range.Merge();
                        range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        range.Borders.LineStyle = XlLineStyle.xlContinuous;
                        range.Borders.Weight = XlBorderWeight.xlMedium;
                        range.Cells.Font.FontStyle = "Bold";
                        num = n;
                        if (num <= 14)
                        {
                            if (num == 4)
                            {
                                range.Cells[1, 1] = "COMPARATIVO ACUMULADO HOGAR";
                            }
                            else if (num == 9)
                            {
                                range.Cells[1, 1] = "COMPARATIVO ACUMULADO BISUTERIA";
                            }
                            else if (num == 14)
                            {
                                range.Cells[1, 1] = "COMPARATIVO ACUMULADO BELLEZA";
                            }
                        }
                        else if (num <= 24)
                        {
                            if (num == 19)
                            {
                                range.Cells[1, 1] = "COMPARATIVO ACUMULADO CALZADO";
                            }
                            else if (num == 24)
                            {
                                range.Cells[1, 1] = "COMPARATIVO ACUMULADO ROPA";
                            }
                        }
                        else if (num == 29)
                        {
                            range.Cells[1, 1] = "COMPARATIVO ACUMULADO SERVICIOS";
                        }
                        /*else if (num == 34)
                        {
                            range.Cells[1, 1] = "COMPARATIVO ACUMULADO SNACKS Y BOTANAS";
                        }*/
                        else if (num == 34)
                        {
                            range.Cells[1, 1] = "COMPARATIVO ACUMULADO ABARROTES";
                        }
                        else if (num == 39)
                        {
                            range.Cells[1, 1] = "COMPARATIVO ACUMULADO TEMPORADA";
                        }
                        for (int o = n; o <= n + 4; o++)
                        {
                            if (o <= n)
                            {
                                range = hoja.Range[string.Concat(this.sCol(o), "4"), string.Concat(this.sCol(o + 1), "4")];
                                range.Select();
                                range.Merge();
                                range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                                range.Borders.Weight = XlBorderWeight.xlMedium;
                                range.Cells.Font.FontStyle = "Bold";
                                range.Cells[1, 1] = "VENTA NETA";
                            }
                            else
                            {
                                hoja.Cells[4, o] = "%";
                                hoja.Cells[4, o].HorizontalAlignment = -4108;
                                hoja.Cells[4, o].VerticalAlignment = -4108;
                                hoja.Cells[4, o].Borders.LineStyle = 1;
                                hoja.Cells[4, o].Borders.Weight = -4138;
                            }
                            if (o == n)
                            {
                                hoja.Cells[5, o] = "OBJETIVO";
                                hoja.Cells[5, o].HorizontalAlignment = -4108;
                                hoja.Cells[5, o].VerticalAlignment = -4108;
                                hoja.Cells[5, o].Borders.LineStyle = 1;
                                hoja.Cells[5, o].Borders.Weight = -4138;
                                hoja.Range[string.Concat(this.sCol(o), "6"), string.Concat(this.sCol(o), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic cells2 = hoja.Cells[this.DG1.Rows.Count + 6, o];
                                strArrays = new string[] { "=SUM(", this.sCol(o), "6:", this.sCol(o), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                cells2.Formula = string.Concat(strArrays);
                            }
                            else if (o == n + 1)
                            {
                                hoja.Cells[5, o].HorizontalAlignment = -4108;
                                hoja.Cells[5, o].NumberFormat = "@";
                                hoja.Cells[5, o] = "REALIZADO";
                                hoja.Range[string.Concat(this.sCol(o), "6"), string.Concat(this.sCol(o), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic obj2 = hoja.Cells[this.DG1.Rows.Count + 6, o];
                                strArrays = new string[] { "=SUM(", this.sCol(o), "6:", this.sCol(o), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                obj2.Formula = string.Concat(strArrays);
                            }
                            else if (o == n + 2)
                            {
                                hoja.Cells[5, o].HorizontalAlignment = -4108;
                                hoja.Cells[5, o].NumberFormat = "@";
                                hoja.Cells[5, o] = "CUMPL";
                                hoja.Range[string.Concat(this.sCol(o), "6"), string.Concat(this.sCol(o), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic cells3 = hoja.Cells[this.DG1.Rows.Count + 6, o];
                                strArrays = new string[] { "=(", this.sCol(o - 1), Convert.ToString(this.DG1.Rows.Count + 6), "/", this.sCol(o - 2), Convert.ToString(this.DG1.Rows.Count + 6), ")*100" };
                                cells3.Formula = string.Concat(strArrays);
                            }
                            else if (o == n + 3)
                            {
                                hoja.Cells[5, o].HorizontalAlignment = -4108;
                                hoja.Cells[5, o].NumberFormat = "@";
                                hoja.Cells[5, o] = "COSTO";
                                hoja.Range[string.Concat(this.sCol(o), "6"), string.Concat(this.sCol(o), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            else if (o == n + 4)
                            {
                                hoja.Cells[5, o].HorizontalAlignment = -4108;
                                hoja.Cells[5, o].NumberFormat = "@";
                                hoja.Cells[5, o] = "MARGEN";
                                hoja.Range[string.Concat(this.sCol(o), "6"), string.Concat(this.sCol(o), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            hoja.Cells[5, o].Borders.LineStyle = 1;
                            hoja.Cells[5, o].Borders.Weight = -4138;
                            hoja.Cells[5, o].Font.FontStyle = "Bold";
                        }
                    }
                    hoja.Range["B6", string.Concat("C", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", string.Concat("AQ", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignRight;
                    hoja.Range["B6", string.Concat("AQ", Convert.ToString(this.DG1.Rows.Count + 6))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                    hoja.Cells[this.DG1.Rows.Count + 6, 3] = "T O T A L";
                    range = (Range)hoja.Cells[6, 1];
                    range.Select();
                    hoja.PasteSpecial(range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    range = hoja.Range["A1", string.Concat("AL", Convert.ToString(this.DG1.Rows.Count + 10))];
                    range.EntireColumn.AutoFit();
                }
                Clipboard.Clear();
                objeto = null;
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@fechaInicial", Convert.ToDateTime(primero.ToString("yyyy-MM-dd")));
                comando.Parameters.AddWithValue("@fechaFinal", dia);
                comando.Parameters.AddWithValue("@tipo", 2);
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter()
                {
                    SelectCommand = comando
                };
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                this.DG1.DataSource = dt;
                this.DG1.SelectAll();
                objeto = this.DG1.GetClipboardContent();
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Worksheets.get_Item(4);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO TOTAL";
                    hoja.Cells[1, 2] = "COMPARATIVO TOTAL";
                    range = hoja.Range["B1", "M1"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B2", "M2"];
                    range.Select();
                    range.Merge();
                    range.Cells[1.1, Type.Missing] = dia.ToString("d MMM yyyy");
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B3", "B5"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "CODIGO";
                    range = hoja.Range["C3", "C5"];
                    range.Select();
                    range.Merge();
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "SUCURSAL";
                    for (int p = 4; p <= 13; p += 5)
                    {
                        range = hoja.Range[string.Concat(this.sCol(p), "3"), string.Concat(this.sCol(p + 4), "3")];
                        range.Select();
                        range.Merge();
                        range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        range.Borders.LineStyle = XlLineStyle.xlContinuous;
                        range.Borders.Weight = XlBorderWeight.xlMedium;
                        range.Cells.Font.FontStyle = "Bold";
                        num = p;
                        if (num == 4)
                        {
                            range.Cells[1, 1] = "COMPARATIVO TOTAL DIARIO";
                        }
                        else if (num == 9)
                        {
                            range.Cells[1, 1] = "COMPARATIVO TOTAL ACUMULADO";
                        }
                        for (int q = p; q <= p + 4; q++)
                        {
                            if (q <= p)
                            {
                                range = hoja.Range[string.Concat(this.sCol(q), "4"), string.Concat(this.sCol(q + 1), "4")];
                                range.Select();
                                range.Merge();
                                range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                                range.Borders.Weight = XlBorderWeight.xlMedium;
                                range.Cells.Font.FontStyle = "Bold";
                                range.Cells[1, 1] = "VENTA NETA";
                            }
                            else
                            {
                                hoja.Cells[4, q] = "%";
                                hoja.Cells[4, q].HorizontalAlignment = -4108;
                                hoja.Cells[4, q].VerticalAlignment = -4108;
                                hoja.Cells[4, q].Borders.LineStyle = 1;
                                hoja.Cells[4, q].Borders.Weight = -4138;
                            }
                            if (q == p)
                            {
                                hoja.Cells[5, q] = "OBJETIVO";
                                hoja.Cells[5, q].HorizontalAlignment = -4108;
                                hoja.Cells[5, q].VerticalAlignment = -4108;
                                hoja.Cells[5, q].Borders.LineStyle = 1;
                                hoja.Cells[5, q].Borders.Weight = -4138;
                                hoja.Range[string.Concat(this.sCol(q), "6"), string.Concat(this.sCol(q), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic obj3 = hoja.Cells[this.DG1.Rows.Count + 6, q];
                                strArrays = new string[] { "=SUM(", this.sCol(q), "6:", this.sCol(q), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                obj3.Formula = string.Concat(strArrays);
                            }
                            else if (q == p + 1)
                            {
                                hoja.Cells[5, q].HorizontalAlignment = -4108;
                                hoja.Cells[5, q].VerticalAlignment = -4108;
                                hoja.Cells[5, q].NumberFormat = "@";
                                hoja.Cells[5, q] = "REALIZADO";
                                hoja.Range[string.Concat(this.sCol(q), "6"), string.Concat(this.sCol(q), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic cells4 = hoja.Cells[this.DG1.Rows.Count + 6, q];
                                strArrays = new string[] { "=SUM(", this.sCol(q), "6:", this.sCol(q), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                cells4.Formula = string.Concat(strArrays);
                            }
                            else if (q == p + 2)
                            {
                                hoja.Cells[5, q].HorizontalAlignment = -4108;
                                hoja.Cells[5, q].VerticalAlignment = -4108;
                                hoja.Cells[5, q].NumberFormat = "@";
                                hoja.Cells[5, q] = "CUMPL";
                                hoja.Range[string.Concat(this.sCol(q), "6"), string.Concat(this.sCol(q), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic obj4 = hoja.Cells[this.DG1.Rows.Count + 6, q];
                                strArrays = new string[] { "=(", this.sCol(q - 1), Convert.ToString(this.DG1.Rows.Count + 6), "/", this.sCol(q - 2), Convert.ToString(this.DG1.Rows.Count + 6), ")*100" };
                                obj4.Formula = string.Concat(strArrays);
                            }
                            else if (q == p + 3)
                            {
                                hoja.Cells[5, q].HorizontalAlignment = -4108;
                                hoja.Cells[5, q].VerticalAlignment = -4108;
                                hoja.Cells[5, q].NumberFormat = "@";
                                hoja.Cells[5, q] = "COSTO";
                                hoja.Range[string.Concat(this.sCol(q), "6"), string.Concat(this.sCol(q), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            else if (q == p + 4)
                            {
                                hoja.Cells[5, q].HorizontalAlignment = -4108;
                                hoja.Cells[5, q].VerticalAlignment = -4108;
                                hoja.Cells[5, q].NumberFormat = "@";
                                hoja.Cells[5, q] = "MARGEN";
                                hoja.Range[string.Concat(this.sCol(q), "6"), string.Concat(this.sCol(q), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            hoja.Cells[5, q].Borders.LineStyle = 1;
                            hoja.Cells[5, q].Borders.Weight = -4138;
                            hoja.Cells[5, q].Font.FontStyle = "Bold";
                        }
                    }

                    // Se agrega la sección de los encabezados de los tickets x hora.
                    int colIndex;
                    int gridTotalCols = this.DG1.ColumnCount;
                    range = hoja.Range["N4", this.sCol(gridTotalCols + 1) + "4"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "CLIENTES";
                    // Se formatean las columnas con los tickets registrados x hora y establecimiento.
                    for (int p = 14; p <= (gridTotalCols + 1); p++)
                    {
                        // Se establecen los encabezados de las columnas y se giran 90° a la izquierda.
                        colIndex = p - 2;
                        hoja.Cells[5, p] = this.DG1.Columns[colIndex].HeaderText;
                        hoja.Cells[5, p].HorizontalAlignment = -4108;
                        hoja.Cells[5, p].VerticalAlignment = -4108;
                        hoja.Cells[5, p].Orientation = 90;
                        hoja.Cells[5, p].Borders.LineStyle = 1;
                        hoja.Cells[5, p].Borders.Weight = -4138;
                        hoja.Cells[5, p].Font.FontStyle = "Bold";
                        hoja.Cells[5, p].Font.Size = 10;
                        // Se coloca la fómula de Sumatoria de los tickets x hora.
                        dynamic cells4 = hoja.Cells[this.DG1.Rows.Count + 6, p];
                        strArrays = new string[] { "=SUM(", this.sCol(p), "6:", this.sCol(p), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                        cells4.Formula = string.Concat(strArrays);
                    }
                    hoja.Range[this.sCol(14) + "6", string.Concat(this.sCol(gridTotalCols + 1), Convert.ToString(this.DG1.Rows.Count + 6))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);

                    hoja.Range["B6", string.Concat("C", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", string.Concat("M", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignRight;
                    hoja.Range["B6", string.Concat("M", Convert.ToString(this.DG1.Rows.Count + 6))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                    hoja.Cells[this.DG1.Rows.Count + 6, 3] = "T O T A L";
                    range = (Range)hoja.Cells[6, 1];
                    range.Select();
                    hoja.PasteSpecial(range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    //range = hoja.Range["A1", string.Concat("M", Convert.ToString(this.DG1.Rows.Count + 10))];
                    range = hoja.Range["A1", string.Concat(this.sCol(gridTotalCols + 1), Convert.ToString(this.DG1.Rows.Count + 10))];
                    range.EntireColumn.AutoFit();
                }
                Clipboard.Clear();
                objeto = null;
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@fechaInicial", Convert.ToDateTime(primero.ToString("yyyy-MM-dd")));
                comando.Parameters.AddWithValue("@fechaFinal", dia);
                comando.Parameters.AddWithValue("@tipo", 3);
                dt = new System.Data.DataTable();
                da = new SqlDataAdapter()
                {
                    SelectCommand = comando
                };
                da.Fill(dt);
                System.Windows.Forms.Application.DoEvents();
                this.DG1.DataSource = dt;
                this.DG1.SelectAll();
                objeto = this.DG1.GetClipboardContent();
                if (objeto != null)
                {
                    Clipboard.SetDataObject(objeto);
                    hoja = (Worksheet)libro.Worksheets.get_Item(5);
                    hoja.Activate();
                    hoja.Name = "COMPARATIVO POR CLASIFICACION";
                    hoja.Cells[1, 2] = "COMPARATIVO CLASIFICACION";
                    range = hoja.Range["B1", "M1"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B2", "M2"];
                    range.Select();
                    range.Merge();
                    range.Cells[1.1, Type.Missing] = dia.ToString("d MMM yyyy");
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = hoja.Range["B3", "B5"];
                    range.Select();
                    range.Merge();
                    range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "CLASIFICACION";
                    range = hoja.Range["C3", "C5"];
                    range.Select();
                    range.Merge();
                    range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Cells.Font.FontStyle = "Bold";
                    range.Cells[1, 1] = "NOMBRE";
                    for (int r = 4; r <= 13; r += 5)
                    {
                        range = hoja.Range[string.Concat(this.sCol(r), "3"), string.Concat(this.sCol(r + 4), "3")];
                        range.Select();
                        range.Merge();
                        range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        range.Borders.LineStyle = XlLineStyle.xlContinuous;
                        range.Borders.Weight = XlBorderWeight.xlMedium;
                        range.Cells.Font.FontStyle = "Bold";
                        num = r;
                        if (num == 4)
                        {
                            range.Cells[1, 1] = "COMPARATIVO TOTAL DIARIO";
                        }
                        else if (num == 9)
                        {
                            range.Cells[1, 1] = "COMPARATIVO TOTAL ACUMULADO";
                        }
                        for (int s = r; s <= r + 4; s++)
                        {
                            if (s <= r)
                            {
                                range = hoja.Range[string.Concat(this.sCol(s), "4"), string.Concat(this.sCol(s + 1), "4")];
                                range.Select();
                                range.Merge();
                                range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                range.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                                range.Borders.Weight = XlBorderWeight.xlMedium;
                                range.Cells.Font.FontStyle = "Bold";
                                range.Cells[1, 1] = "VENTA NETA";
                            }
                            else
                            {
                                hoja.Cells[4, s] = "%";
                                hoja.Cells[4, s].HorizontalAlignment = -4108;
                                hoja.Cells[4, s].VerticalAlignment = -4108;
                                hoja.Cells[4, s].Borders.LineStyle = 1;
                                hoja.Cells[4, s].Borders.Weight = -4138;
                            }
                            if (s == r)
                            {
                                hoja.Cells[5, s] = "OBJETIVO";
                                hoja.Cells[5, s].HorizontalAlignment = -4108;
                                hoja.Cells[5, s].VerticalAlignment = -4108;
                                hoja.Cells[5, s].Borders.LineStyle = 1;
                                hoja.Cells[5, s].Borders.Weight = -4138;
                                hoja.Range[string.Concat(this.sCol(s), "6"), string.Concat(this.sCol(s), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic cells5 = hoja.Cells[this.DG1.Rows.Count + 6, s];
                                strArrays = new string[] { "=SUM(", this.sCol(s), "6:", this.sCol(s), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                cells5.Formula = string.Concat(strArrays);
                            }
                            else if (s == r + 1)
                            {
                                hoja.Cells[5, s].HorizontalAlignment = -4108;
                                hoja.Cells[5, s].NumberFormat = "@";
                                hoja.Cells[5, s] = "REALIZADO";
                                hoja.Range[string.Concat(this.sCol(s), "6"), string.Concat(this.sCol(s), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic obj5 = hoja.Cells[this.DG1.Rows.Count + 6, s];
                                strArrays = new string[] { "=SUM(", this.sCol(s), "6:", this.sCol(s), Convert.ToString(this.DG1.Rows.Count + 5), ")" };
                                obj5.Formula = string.Concat(strArrays);
                            }
                            else if (s == r + 2)
                            {
                                hoja.Cells[5, s].HorizontalAlignment = -4108;
                                hoja.Cells[5, s].NumberFormat = "@";
                                hoja.Cells[5, s] = "CUMPL";
                                hoja.Range[string.Concat(this.sCol(s), "6"), string.Concat(this.sCol(s), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                                dynamic cells6 = hoja.Cells[this.DG1.Rows.Count + 6, s];
                                strArrays = new string[] { "=(", this.sCol(s - 1), Convert.ToString(this.DG1.Rows.Count + 6), "/", this.sCol(s - 2), Convert.ToString(this.DG1.Rows.Count + 6), ")*100" };
                                cells6.Formula = string.Concat(strArrays);
                            }
                            else if (s == r + 3)
                            {
                                hoja.Cells[5, s].HorizontalAlignment = -4108;
                                hoja.Cells[5, s].NumberFormat = "@";
                                hoja.Cells[5, s] = "COSTO";
                                hoja.Range[string.Concat(this.sCol(s), "6"), string.Concat(this.sCol(s), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            else if (s == r + 4)
                            {
                                hoja.Cells[5, s].HorizontalAlignment = -4108;
                                hoja.Cells[5, s].NumberFormat = "@";
                                hoja.Cells[5, s] = "MARGEN";
                                hoja.Range[string.Concat(this.sCol(s), "6"), string.Concat(this.sCol(s), Convert.ToString(this.DG1.Rows.Count + 6))].NumberFormat = "#,###,##0.00";
                            }
                            hoja.Cells[5, s].Borders.LineStyle = 1;
                            hoja.Cells[5, s].Borders.Weight = -4138;
                            hoja.Cells[5, s].Font.FontStyle = "Bold";
                        }
                    }
                    hoja.Range["B6", string.Concat("C", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    hoja.Range["D6", string.Concat("M", Convert.ToString(this.DG1.Rows.Count + 5))].HorizontalAlignment = XlHAlign.xlHAlignRight;
                    hoja.Range["B6", string.Concat("M", Convert.ToString(this.DG1.Rows.Count + 6))].BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                    hoja.Cells[this.DG1.Rows.Count + 6, 3] = "T O T A L";
                    range = (Range)hoja.Cells[6, 1];
                    range.Select();
                    hoja.PasteSpecial(range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    range = hoja.Range["A1", string.Concat("M", Convert.ToString(this.DG1.Rows.Count + 10))];
                    range.EntireColumn.AutoFit();
                }
                hoja = (Worksheet)libro.Worksheets.get_Item(1);
                hoja.Activate();
                if (cn.State.ToString() == "Open")
                {
                    cn.Close();
                }
                if (File.Exists(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Indicador Comparativo VS Presupuesto al ", dia.ToString("dd MMMM yyyy"), ".xlsx")))
                {
                    File.Delete(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Indicador Comparativo VS Presupuesto al ", dia.ToString("dd MMMM yyyy"), ".xlsx"));
                }
                libro.SaveAs(string.Concat(System.Windows.Forms.Application.StartupPath, "\\Indicador Comparativo VS Presupuesto al ", dia.ToString("dd MMMM yyyy"), ".xlsx"), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                libro.Close();
                excel.Quit();
                str = string.Concat(System.Windows.Forms.Application.StartupPath, "\\Indicador Comparativo VS Presupuesto al ", dia.ToString("dd MMMM yyyy"), ".xlsx");
            }
            catch (Exception exception)
            {
                str = "";
            }
            return str;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;

            /**
             * Reporte de Apertura
             * Reporte de Ventas con Descuentos
             * Reporte de Incidencias
             **/
            if (now.ToString("HH:mm").Trim() == "01:01")
            {
                this.lblEstado.Text = "Enviando Reporte de Apertura...";
                this.EnviaMailGmail(this.ReporteApertura(DateTime.Now), "luis.guerrero@mercadodeimportaciones.com,alberto.martinez@mercadodeimportaciones.com,prevencion@mercadodeimportaciones.com,francisco.ontiveros@mercadodeimportaciones.com,Gerencia.Sistemas@mercadodeimportaciones.com,maferperezle01@gmail.com");
                
                this.lblEstado.Text = "Enviando Reporte de Descuentos...";
                this.EnviaMailGmail(this.ReporteDescuento(DateTime.Now), "monica.perez@mercadodeimportaciones.com,jassiel.perez@mercadodeimportaciones.com,alberto.martinez@mercadodeimportaciones.com,francisco.ontiveros@mercadodeimportaciones.com,luis.guerrero@mercadodeimportaciones.com,mercadotecniaauxiliar@mercadodeimportaciones.com,luis.cota@mercadodeimportaciones.com,mario.serrano@mercadodeimportaciones.com,annacelia.soto@mercadodeimportaciones.com,asofia.mercadodeimportaciones@gmail.com,Gerencia.Sistemas@mercadodeimportaciones.com,analista.comercial@mercadodeimportaciones.com,hibrajid.lara@mercadodeimportaciones.com");
                
                this.lblEstado.Text = "Enviando Reporte de Incidencias...";
                this.Incidencias();
                
                this.lblEstado.Text = "Proceso de envío de Reporte de Apertura, Reporte de Descuentos y Reporte de Incidencias - FINALIZADO EXITOSAMENTE.";
            }

            /**
             * Reporte Top 80
             **/
            if (now.ToString("HH:mm").Trim() == "01:30")
            {
                this.lblEstado.Text = "Enviando Reporte Top 80...";
                this.Top80();
                this.lblEstado.Text = "Proceso de envío de Reporte Top 80 - FINALIZADO EXITOSAMENTE.";
            }

            //if (now.ToString("HH:mm").Trim() == "05:00")
            //{
            //    DateTime dateTime = DateTime.Now;
            //    DayOfWeek dayOfWeek = dateTime.DayOfWeek;
            //}

            /**
             *  Reporte Comparativo vs Presupuesto
             **/
            if (/*(now.ToString("HH:mm").Trim() == "12:00") | (now.ToString("HH:mm").Trim() == "15:00") | */(now.ToString("HH:mm").Trim() == "18:00") | (now.ToString("HH:mm").Trim() == "22:35"))
            {
                this.lblEstado.Text = "Enviando Reporte Comparativo vs Presupuesto...";
                this.EnviaMailGmail(this.ComparativoPresupuesto(DateTime.Now), "director@mercadodeimportaciones.com,guillermina.cervantes@mercadodeimportaciones.com,culiacan@mercadodeimportaciones.com,guamuchilcentro@mercadodeimportaciones.com,guamuchilplaza@mercadodeimportaciones.com,guasavecentro@mercadodeimportaciones.com,guaymas@mercadodeimportaciones.com,hermosillocentro@mercadodeimportaciones.com,hermosillosendero@mercadodeimportaciones.com,mochis@mercadodeimportaciones.com,navojoa@mercadodeimportaciones.com,navolato@mercadodeimportaciones.com,obregon@mercadodeimportaciones.com,guasaveplaza@mercadodeimportaciones.com,luis.guerrero@mercadodeimportaciones.com,karmina.alcala@mercadodeimportaciones.com,carrasco@mercadodeimportaciones.com,escobedo@mercadodeimportaciones.com,navolato.imp@mercadodeimportaciones.com,luis.cota@mercadodeimportaciones.com,mario.serrano@mercadodeimportaciones.com,eduardope01@hotmail.com,annacelia.soto@mercadodeimportaciones.com,alberto.martinez@mercadodeimportaciones.com,recepcion@mercadodeimportaciones.com,rubi@mercadodeimportaciones.com,hermosillomonterrey@mercadodeimportaciones.com,ladiesobregon@mercadodeimportaciones.com,guaymas.serdan@mercadodeimportaciones.com,monica.perez@mercadodeimportaciones.com,terranova@mercadodeimportaciones.com,asofia.mercadodeimportaciones@gmail.com,cabosendero@mercadodeimportaciones.com,jassiel.perez@mercadodeimportaciones.com,soporte.sinaloa@mercadodeimportaciones.com,mazatlanelmar@mercadodeimportaciones.com,jesus.suarez@mercadodeimportaciones.com,senderos.culiacan@mercadodeimportaciones.com,abastos@mercadodeimportaciones.com,mazatlanaquiles@mercadodeimportaciones.com,mercadotecnia@mercadodeimportaciones.com,analista.comercial@mercadodeimportaciones.com,barrancos@mercadodeimportaciones.com,mochissendero@mercadodeimportaciones.com,obregonsendero@mercadodeimportaciones.com,san.isidro@mercadodeimportaciones.com,hermosillomatamoros@mercadodeimportaciones.com,jesus.salazar@mercadodeimportaciones.com,huatabampo@mercadodeimportaciones.com,silvia.bojorquez@mercadodeimportaciones.com,cabosendero2@mercadodeimportaciones.com,auxiliar.sistemas@mercadodeimportaciones.com,maferperezle01@gmail.com,auxiliar.ventas@mercadodeimportaciones.com,lapaz@mercadodeimportaciones.com,obregonplaza@mercadodeimportaciones.com");
                //this.EnviaMail(this.ComparativoPresupuesto(DateTime.Now), "director@mercadodeimportaciones.com,guillermina.cervantes@mercadodeimportaciones.com,culiacan@mercadodeimportaciones.com,guamuchilcentro@mercadodeimportaciones.com,guamuchilplaza@mercadodeimportaciones.com,guasavecentro@mercadodeimportaciones.com,guaymas@mercadodeimportaciones.com,hermosillocentro@mercadodeimportaciones.com,hermosillosendero@mercadodeimportaciones.com,mochis@mercadodeimportaciones.com,navojoa@mercadodeimportaciones.com,navolato@mercadodeimportaciones.com,obregon@mercadodeimportaciones.com,guasaveplaza@mercadodeimportaciones.com,luis.guerrero@mercadodeimportaciones.com,karmina.alcala@mercadodeimportaciones.com,carrasco@mercadodeimportaciones.com,escobedo@mercadodeimportaciones.com,navolato.imp@mercadodeimportaciones.com,luis.cota@mercadodeimportaciones.com,mario.serrano@mercadodeimportaciones.com,eduardope01@hotmail.com,annacelia.soto@mercadodeimportaciones.com,alberto.martinez@mercadodeimportaciones.com,recepcion@mercadodeimportaciones.com,rubi@mercadodeimportaciones.com,hermosillomonterrey@mercadodeimportaciones.com,ladiesobregon@mercadodeimportaciones.com,guaymas.serdan@mercadodeimportaciones.com,monica.perez@mercadodeimportaciones.com,terranova@mercadodeimportaciones.com,asofia.mercadodeimportaciones@gmail.com,cabosendero@mercadodeimportaciones.com,jassiel.perez@mercadodeimportaciones.com,soporte.sinaloa@mercadodeimportaciones.com,mazatlanelmar@mercadodeimportaciones.com,jesus.suarez@mercadodeimportaciones.com,senderos.culiacan@mercadodeimportaciones.com,abastos@mercadodeimportaciones.com,mazatlanaquiles@mercadodeimportaciones.com,mercadotecnia@mercadodeimportaciones.com,analista.comercial@mercadodeimportaciones.com,barrancos@mercadodeimportaciones.com,mochissendero@mercadodeimportaciones.com,obregonsendero@mercadodeimportaciones.com,san.isidro@mercadodeimportaciones.com,hermosillomatamoros@mercadodeimportaciones.com,jesus.salazar@mercadodeimportaciones.com,huatabampo@mercadodeimportaciones.com,silvia.bojorquez@mercadodeimportaciones.com,cabosendero2@mercadodeimportaciones.com,auxiliar.sistemas@mercadodeimportaciones.com,maferperezle01@gmail.com,auxiliar.ventas@mercadodeimportaciones.com");

                this.lblEstado.Text = "Proceso de envío de Reporte Comparativo vs Presupuesto - FINALIZADO EXITOSAMENTE.";
            }

            /**
             * Reporte Cierre CEDIS
             **/
            if (now.ToString("HH:mm").Trim() == "17:30" || now.ToString("HH:mm").Trim() == "22:00")
            {
                this.lblEstado.Text = "Enviando Reporte Cierre CEDIS...";
                this.EnviaMailGmail(this.CierreCedis(DateTime.Now), "guillermina.cervantes@mercadodeimportaciones.com,cedis@mercadodeimportaciones.com,cedis.recibo@mercadodeimportaciones.com,cedis.inventarios@mercadodeimportaciones.com,monica.perez@mercadodeimportaciones.com,jassiel.perez@mercadodeimportaciones.com,gilberto.govea@mercadodeimportaciones.com,hibrajid.lara@mercadodeimportaciones.com");
                this.lblEstado.Text = "Proceso de envío de Reporte Cierre CEDIS - FINALIZADO EXITOSAMENTE.";
            }

            /**
             * 
             * Reporte Indicador de Presupuesto
             * 
             **/
            if (now.ToString("HH:mm").Trim() == "22:45")
            {
                this.lblEstado.Text = "Enviando Reporte Indicador de Presupuesto...";
                this.Presupuesto();
                this.lblEstado.Text = "Proceso de envío de Reporte de Presupuesto - FINALIZADO EXITOSAMENTE.";
            }

            /**
             * 
             * Reporte Existencias En Dias Ventas
             * 
             **/
            if (now.ToString("HH:mm").Trim() == "06:00")
            {
                DateTime dt = DateTime.Now;

                if (dt.DayOfWeek == DayOfWeek.Monday || dt.DayOfWeek == DayOfWeek.Friday)
                {
                    this.lblEstado.Text = "Enviando Reporte Existencias En Dias Ventas...";
                    this.EnviaMailGmail(this.Existencias_dias_ventas(DateTime.Now), "guillermina.cervantes@mercadodeimportaciones.com,annacelia.soto@mercadodeimportaciones.com,luis.cota@mercadodeimportaciones.com,mario.serrano@mercadodeimportaciones.com,Gerencia.Sistemas@mercadodeimportaciones.com,monica.perez@mercadodeimportaciones.com,jassiel.perez@mercadodeimportaciones.com,gilberto.govea@mercadodeimportaciones.com,alberto.martinez@mercadodeimportaciones.com,luis.guerrero@mercadodeimportaciones.com,analista.comercial@mercadodeimportaciones.com,jesus.suarez@mercadodeimportaciones.com,hibrajid.lara@mercadodeimportaciones.com,maferperezle01@gmail.com");
                    this.lblEstado.Text = "Proceso de envio de Reporte Existencias dias Ventas - FINALIZADO EXITOSAMENTE";
                }

            }

            /**
             * 
             * Reporte Venta De Articulos 30 Dias Venta
             * 
             **/
            //  de 1 hr a 36 min  se coloco 40
            // No mandar luis guerrero y cordinadores
            if (now.ToString("HH:mm").Trim() == "05:00")
            {
                DateTime dt = DateTime.Now;
                //DateTime dt = new DateTime(2021, 6, 14);
                if (dt.DayOfWeek == DayOfWeek.Monday)
                {
                    SqlConnection cn1 = conexion.conectar("BMSNayar");
                    SqlCommand comando1 = new SqlCommand("select mail.cod_estab,substring(replace(upper(establecimientos.nombre),'.',''),1,30) as nombre,mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ( '1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
                    SqlDataReader dr1 = comando1.ExecuteReader();
                    if (dr1.HasRows)
                    {
                        while (dr1.Read())
                        {
                            this.lblEstado.Text = "Enviando Reporte Venta De Articulos 30 Dias Venta...";
                            this.EnviaMailGmail(this.Ventas_Articulos_30(dt, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), dr1["email"].ToString() + ",analista.comercial@mercadodeimportaciones.com");
                        }
                    }
                    if (dr1.IsClosed == false) { dr1.Close(); }
                    if (cn1.State == ConnectionState.Open) { cn1.Close(); }

                    this.lblEstado.Text = "Proceso de envio de Reporte Venta De Articulos 30 Dias Venta - FINALIZADO EXITOSAMENTE";

                }
            }

            /**
             * 
             * Reporte Comparativo semana-semana
             * 
             **/
            // No mandar a luis guerrero y cordinadores  
            //Comparativo venta semana-semana  1/2 hora a 10 min  tiempo 13 min
            if (now.ToString("HH:mm").Trim() == "05:40")
            {
                DateTime dt = DateTime.Now;
                //DateTime dt = new DateTime(2021, 4, 5);
                if (dt.DayOfWeek == DayOfWeek.Monday)
                {
                    SqlConnection cn1 = conexion.conectar("BMSNayar");
                    SqlCommand comando1 = new SqlCommand("select mail.cod_estab,substring(replace(upper(establecimientos.nombre),'.',''),1,30) as nombre,mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ( '1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
                    SqlDataReader dr1 = comando1.ExecuteReader();
                    if (dr1.HasRows)
                    {
                        while (dr1.Read())
                        {
                            this.lblEstado.Text = "Enviando Reporte Comparativo semana-semana...";
                            this.EnviaMailGmail(this.Comparativo_semana_semana(dt, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), dr1["email"].ToString() + ",analista.comercial@mercadodeimportaciones.com");
                        }
                    }
                    if (dr1.IsClosed == false) { dr1.Close(); }
                    if (cn1.State == ConnectionState.Open) { cn1.Close(); }
                    this.lblEstado.Text = "Proceso de envio de Reporte Comparativo semana-semana - FINALIZADO EXITOSAMENTE";
                }
            }

            // 07/Jul/2021 - Noé García - Deshabilitado el envío de este reporte por solicitud de Vanessa Soto
            //Sorteo Dinamica tv  de 10 min a 5 
            //if (now.ToString("HH:mm").Trim() == "05:53")
            //{
            //    DateTime dt = DateTime.Now;
            //    SqlConnection cn1 = conexion.conectar("BMSNayar");
            //    SqlCommand comando1 = new SqlCommand("SELECT STUFF((SELECT ',' + mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '39', '43', '48', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005') order by CAST(mail.cod_estab as int) FOR XML PATH('')),1,1, '') as email", cn1);
            //    SqlDataReader dr1 = comando1.ExecuteReader();
            //    if (dr1.HasRows)
            //    {
            //        while (dr1.Read())
            //        {
            //            this.lblEstado.Text = "Enviando Reporte Dinamicas Por Temporada...";
            //            this.EnviaMailGmail(this.Boletos_Sorteo_Tv(dt), dr1["email"].ToString() + ",alberto.martinez@mercadodeimportaciones.com,luis.guerrero@mercadodeimportaciones.com,analista.comercial@mercadodeimportaciones.com,jesus.salazar@mercadodeimportaciones.com,jesus.suarez@mercadodeimportaciones.com");
            //        }
            //    }
            //    if (dr1.IsClosed == false) { dr1.Close(); }
            //    if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            //    this.lblEstado.Text = "Proceso de envio de Reporte Dinamicas Por Temporada - FINALIZADO EXITOSAMENTE";
            //}
            //Duracion 10 min

            /**
             * 
             * Reporte Acumulado De Ventas Mensual
             *  
             **/
            if (now.ToString("HH:mm").Trim() == "06:15")
            {
                DateTime dt = DateTime.Now;
                //DateTime dt = new DateTime(2021, 04, 1);

                SqlConnection cn1 = conexion.conectar("BMSNayar");
                SqlCommand comando1 = new SqlCommand("select mail.cod_estab,substring(replace(upper(establecimientos.nombre),'.',''),1,30) as nombre,mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ( '1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
                SqlDataReader dr1 = comando1.ExecuteReader();
                if (dr1.HasRows)
                {
                    while (dr1.Read())
                    {
                        this.lblEstado.Text = "Enviando Reporte Acumulado De Ventas Mensual...";
                        this.EnviaMailGmail(this.Acumulado_Ventas_Mensual(dt, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), dr1["email"].ToString() + ",analista.comercial@mercadodeimportaciones.com,mercadotecnia@mercadodeimportaciones.com");
                    }
                }
                if (dr1.IsClosed == false) { dr1.Close(); }
                if (cn1.State == ConnectionState.Open) { cn1.Close(); }

                this.lblEstado.Text = "Proceso de envio de Reporte Acumulado De Ventas Mensual - FINALIZADO EXITOSAMENTE";
            }

            /**
             * 
             * Reporte Señalizacion de Promociones
             *  
             **/
            if (now.ToString("HH:mm").Trim() == "06:30")
            {
                DateTime dt = DateTime.Now;

                if (dt.DayOfWeek == DayOfWeek.Monday)
                {
                    SqlConnection cn1 = conexion.conectar("BMSNayar");
                    SqlCommand comando1 = new SqlCommand("SELECT STUFF((SELECT ',' + mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int) FOR XML PATH('')),1,1, '') as email", cn1);
                    SqlDataReader dr1 = comando1.ExecuteReader();
                    if (dr1.HasRows)
                    {
                        while (dr1.Read())
                        {
                            this.lblEstado.Text = "Enviando Reporte Señalizacion de Promociones...";
                            this.EnviaMailGmail(this.Señalizacion_promociones(dt), dr1["email"].ToString() + ",alberto.martinez@mercadodeimportaciones.com,luis.guerrero@mercadodeimportaciones.com,analista.comercial@mercadodeimportaciones.com,jesus.salazar@mercadodeimportaciones.com,jesus.suarez@mercadodeimportaciones.com");
                        }
                    }
                    if (dr1.IsClosed == false) { dr1.Close(); }
                    if (cn1.State == ConnectionState.Open) { cn1.Close(); }

                    this.lblEstado.Text = "Proceso de envio de Reporte Señalizacion de Promociones - FINALIZADO EXITOSAMENTE";
                }

            }

            /**
             * 
             * Reporte Reporte Desplazamiento De Temporada
             *  
             **/
            if (now.ToString("HH:mm").Trim() == "06:40")
            {
                DateTime dt = DateTime.Now;
                //DateTime dt = new DateTime(2021, 4, 26);
                if (dt.DayOfWeek == DayOfWeek.Monday)
                {
                    SqlConnection cn1 = conexion.conectar("BMSNayar");
                    SqlCommand comando1 = new SqlCommand("SELECT STUFF((SELECT ',' + mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int) FOR XML PATH('')),1,1, '') as email", cn1);
                    SqlDataReader dr1 = comando1.ExecuteReader();
                    if (dr1.HasRows)
                    {
                        while (dr1.Read())
                        {
                            this.lblEstado.Text = "Enviando Reporte Desplazamiento De Temporada...";
                            this.EnviaMailGmail(this.Desplazamiento_Temporada(dt), dr1["email"].ToString() + ",alberto.martinez@mercadodeimportaciones.com,luis.guerrero@mercadodeimportaciones.com,analista.comercial@mercadodeimportaciones.com,jesus.salazar@mercadodeimportaciones.com,jesus.suarez@mercadodeimportaciones.com");
                        }
                    }
                    if (dr1.IsClosed == false) { dr1.Close(); }
                    if (cn1.State == ConnectionState.Open) { cn1.Close(); }

                    this.lblEstado.Text = "Proceso de envio de Reporte Reporte Desplazamiento De Temporada - FINALIZADO EXITOSAMENTE";
                }
            }

            /**
             * 
             * Reporte Articulos Menos Vendidos
             *  
             **/
            //luis y cordinadores 
            /*Menos Vendidos*/
            if (now.ToString("HH:mm").Trim() == "06:45")
            {
                DateTime dt = DateTime.Now;
                if (dt.DayOfWeek == DayOfWeek.Monday || dt.DayOfWeek == DayOfWeek.Friday)
                {
                    SqlConnection cn1 = conexion.conectar("BMSNayar");
                    SqlCommand comando1 = new SqlCommand("select mail.cod_estab, substring(replace(upper(establecimientos.nombre), '.', ''), 1, 30) as nombre, mail.email from BMSNayar.dbo.establecimientos inner join BMSNayar.dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '39', '43', '48', '79', '98', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
                    SqlDataReader dr1 = comando1.ExecuteReader();
                    if (dr1.HasRows)
                    {
                        while (dr1.Read())
                        {
                            this.lblEstado.Text = "Enviando Reporte Articulos Menos Vendidos...";
                            this.EnviaMailGmail(this.Menos_vendidos(dt, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), dr1["email"].ToString() + ",analista.comercial@mercadodeimportaciones.com");
                        }
                    }
                    if (dr1.IsClosed == false) { dr1.Close(); }
                    if (cn1.State == ConnectionState.Open) { cn1.Close(); }

                    this.lblEstado.Text = "Proceso de envio de Reporte Articulos Menos Vendidos - FINALIZADO EXITOSAMENTE";
                }

            }

            /**
             * 
             * Reporte de Cancelaciones y Devoluciones
             *  
             **/
            //luis y cordinadores 
            /*Cancelaciones y Devoluciones*/
            if (now.ToString("HH:mm").Trim() == "07:00")
            {
                DateTime dt = DateTime.Now;
                if (dt.DayOfWeek == DayOfWeek.Monday)
                {
                    SqlConnection cn1 = conexion.conectar("BMSNayar");
                    SqlCommand comando1 = new SqlCommand("select distinct mail.email_cordinador, coordinador from BMSNayar.dbo.establecimientos inner join BMSNayar.dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.[status] = 'V' and establecimientos.cod_estab not in ('1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by email_cordinador", cn1);
                    SqlDataReader dr1 = comando1.ExecuteReader();
                    if (dr1.HasRows)
                    {
                        while (dr1.Read())
                        {
                            this.lblEstado.Text = "Enviando Reporte Cancelaciones y Devoluciones...";
                            this.EnviaMailGmail(this.DevolucionesCancelacionesSucursal(dr1["email_cordinador"].ToString(), dr1["coordinador"].ToString(), dt), dr1["email_cordinador"].ToString() + ",luis.guerrero@mercadodeimportaciones.com,maferperezle01@gmail.com");
                        }
                    }
                    if (dr1.IsClosed == false) { dr1.Close(); }
                    if (cn1.State == ConnectionState.Open) { cn1.Close(); }

                    this.lblEstado.Text = "Proceso de envio de Reporte Cancelaciones y Devoluciones - FINALIZADO EXITOSAMENTE";
                }

            }

            #region "Reportes Ventas Bms" 

            //if ((now.ToString("HH:mm").Trim() == "12:00" || now.ToString("HH:mm").Trim() == "16:00" || now.ToString("HH:mm").Trim() == "18:30")) 
            //{

            //    //Manda informacion de sucursales a coordinadores
            //    SqlConnection cn1 = conexion.conectar("BMSNayar");
            //    SqlCommand comando1 = new SqlCommand("select  distinct  mail.email_cordinador from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '1001', '1002', '1003', '1004', '1005', '1006')", cn1);
            //    SqlDataReader dr1 = comando1.ExecuteReader();
            //    if (dr1.HasRows)
            //    {
            //        while (dr1.Read())
            //        {
            //            this.lblEstado.Text = "Enviando Reportes de ventas...";
            //            this.EnviaMail(this.Reporte_Ventas(DateTime.Now, ""), dr1["email_cordinador"].ToString());
            //            this.lblEstado.Text = "Proceso de envio de Reportes de ventas diaria - FINALIZADO EXITOSAMENTE";
            //        }
            //    }
            //    if (dr1.IsClosed == false) { dr1.Close(); }
            //    if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            //    //Manda informacion de cada sucursal 
            //    SqlConnection cn = conexion.conectar("BMSNayar");
            //    SqlCommand comando = new SqlCommand("select establecimientos.cod_estab,establecimientos.nombre,mail.email,mail.email_cordinador "
            //    + " from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab=mail.cod_estab"
            //    + " where establecimientos.status='V' and establecimientos.cod_estab not in ('1','1001','1002','1003','1004','1005','1006')"
            //    + " order by cast(establecimientos.cod_estab as int) asc", cn);
            //    SqlDataReader dr = comando.ExecuteReader();
            //    if (dr.HasRows)
            //    {
            //        while (dr.Read())
            //        {
            //            this.lblEstado.Text = "Enviando Reportes de ventas...";
            //            this.EnviaMail(this.Reporte_Ventas(DateTime.Now, dr["cod_estab"].ToString()), dr["email"].ToString());
            //            this.lblEstado.Text = "Proceso de envio de Reportes de ventas diaria - FINALIZADO EXITOSAMENTE";
            //        }
            //    }
            //    if (dr.IsClosed == false) { dr.Close(); }
            //    if (cn.State == ConnectionState.Open) { cn.Close(); }
            //}
            #endregion
        }

        private string obtenerNombreMesAbreviadoNumero(int numeroMes)
        {
            try
            {
                DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                string nombreMesAbreviado = formatoFecha.GetAbbreviatedMonthName(numeroMes);
                return nombreMesAbreviado;
            }
            catch
            {
                return "Desconocido";
            }
        }

        private string Reporte_Ventas(DateTime dia, string sucursal)
        {
            SqlConnection con = new SqlConnection(@"Server=server-cln,1433;Database=BMSNayar; User Id=eliasb;Password=23032004;Connection Timeout=0");
            SqlDataAdapter da = new SqlDataAdapter("F115_Datos", con);
            da.SelectCommand.CommandType = CommandType.StoredProcedure;
            da.SelectCommand.Parameters.Clear();

            da.SelectCommand.Parameters.Add("@FechaI", SqlDbType.DateTime).Value = DateTime.Now.ToString("dd/MM/yyyy") + " 00:00:00";
            da.SelectCommand.Parameters.Add("@FechaF", SqlDbType.DateTime).Value = DateTime.Now.ToString("dd/MM/yyyy") + " 23:59:00";
            da.SelectCommand.Parameters.Add("@Filtro", SqlDbType.NVarChar).Value = "";
            da.SelectCommand.Parameters.Add("@CondPago", SqlDbType.VarChar).Value = "";
            da.SelectCommand.Parameters.Add("@TipoAt", SqlDbType.VarChar).Value = "";
            da.SelectCommand.Parameters.Add("@FichaTec", SqlDbType.Bit).Value = 0;

            da.SelectCommand.Parameters.Add("@Sustitutiva", SqlDbType.Bit).Value = 1;
            da.SelectCommand.Parameters.Add("@IVAIncluido", SqlDbType.Bit).Value = 1;
            da.SelectCommand.Parameters.Add("@Todos", SqlDbType.Bit).Value = 0;
            da.SelectCommand.Parameters.Add("@Estab", SqlDbType.VarChar).Value = sucursal;
            da.SelectCommand.Parameters.Add("@GrupoEstab", SqlDbType.VarChar).Value = "";
            da.SelectCommand.Parameters.Add("@EstabInv", SqlDbType.Bit).Value = 0;
            da.SelectCommand.Parameters.Add("@Creditos", SqlDbType.Bit).Value = 1;

            da.SelectCommand.Parameters.Add("@IdPres", SqlDbType.SmallInt).Value = 12;
            da.SelectCommand.Parameters.Add("@NombreAgrup", SqlDbType.VarChar).Value = "Linea de Producto";
            da.SelectCommand.Parameters.Add("@FactorMoneda", SqlDbType.Decimal).Value = 1.0000;
            da.SelectCommand.Parameters.Add("@FactorVolumen", SqlDbType.Decimal).Value = 1.0000;
            da.SelectCommand.Parameters.Add("@FactorPeso", SqlDbType.Decimal).Value = 1.0000;
            da.SelectCommand.Parameters.Add("@Servicios", SqlDbType.Bit).Value = 1;
            da.SelectCommand.Parameters.Add("@SoloPqts", SqlDbType.Bit).Value = 0;
            da.SelectCommand.Parameters.Add("@Utilidad", SqlDbType.Bit).Value = 1;
            da.SelectCommand.Parameters.Add("@Usuario", SqlDbType.Char).Value = "50";

            System.Data.DataTable dt = new System.Data.DataTable();
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();

            da.Fill(dt);

            this.DG1.DataSource = dt;

            DG1.Columns[3].DefaultCellStyle.Format = "$ #,#0.00";
            DG1.Columns[4].DefaultCellStyle.Format = "$ #,#0.00";
            DG1.Columns[5].DefaultCellStyle.Format = "$ #,#0.00";

            DG1.Columns[7].DefaultCellStyle.Format = "#,#0";
            DG1.Columns[8].DefaultCellStyle.Format = "#,#0";
            DG1.Columns[9].DefaultCellStyle.Format = "#,#0";
            DG1.Columns[10].DefaultCellStyle.Format = "#,#0";
            DG1.Columns[11].DefaultCellStyle.Format = "#,#0";
            DG1.Columns[12].DefaultCellStyle.Format = "$ #,#0.00";
            DG1.Columns[13].DefaultCellStyle.Format = "$ #,#0.00";

            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Worksheets.Add();

            object objeto = DG1.GetClipboardContent();

            Microsoft.Office.Interop.Excel.Range rango;

            hoja.Name = "Resumen De Ventas";


            if (objeto != null)
            {

                hoja.Cells[1, 2] = "Resúmen De Ventas Dia " + DateTime.Now.ToString("dd/MM/yyyy");
                hoja.Cells[1, 2].Font.FontStyle = "Bold";

                hoja.Cells[3, 2] = "Linea Producto";
                hoja.Cells[3, 2].Font.FontStyle = "Bold";

                hoja.Cells[3, 3] = "Codigo";
                hoja.Cells[3, 3].Font.FontStyle = "Bold";



                hoja.Cells[3, 4] = "Nombre";
                hoja.Cells[3, 4].Font.FontStyle = "Bold";

                hoja.Cells[3, 5] = "Venta";
                hoja.Cells[3, 5].Font.FontStyle = "Bold";

                hoja.Cells[3, 6] = "Importe Sin Iva";
                hoja.Cells[3, 6].Font.FontStyle = "Bold";

                hoja.Cells[3, 7] = "Contribucion";
                hoja.Cells[3, 7].Font.FontStyle = "Bold";

                hoja.Cells[3, 8] = "Volumen";

                hoja.Cells[3, 8].Font.FontStyle = "Bold";

                hoja.Cells[3, 9] = "Peso";
                hoja.Cells[3, 9].Font.FontStyle = "Bold";
                hoja.Cells[3, 10] = "Unidades Totales";
                hoja.Cells[3, 10].Font.FontStyle = "Bold";
                hoja.Cells[3, 11] = "Presentaciones Totales";
                hoja.Cells[3, 11].Font.FontStyle = "Bold";
                hoja.Cells[3, 12] = "Unidades";
                hoja.Cells[3, 12].Font.FontStyle = "Bold";
                hoja.Cells[3, 13] = "Presentaciones";
                hoja.Cells[3, 13].Font.FontStyle = "Bold";
                hoja.Cells[3, 14] = "Gasto";
                hoja.Cells[3, 14].Font.FontStyle = "Bold";
                hoja.Cells[3, 15] = "Descuento Financiero";
                hoja.Cells[3, 15].Font.FontStyle = "Bold";
                hoja.Cells[3, 16] = "Nombre Grupo";
                hoja.Cells[3, 16].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("A2", "Q2");
                rango.Columns.AutoFit();

                Clipboard.SetDataObject(objeto);
                rango = (Range)hoja.Cells[4, 1];
                rango.Select();
                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            }


            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Rp_ventas_Sucursal_" + sucursal.Trim() + "_" + dia.ToString("dd MMMM yyyy") + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Rp_ventas_Sucursal_" + sucursal.Trim() + "_" + dia.ToString("dd MMMM yyyy") + ".xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Rp_ventas_Sucursal_" + sucursal.Trim() + "_" + dia.ToString("dd MMMM yyyy") + ".xlsx");
            g_Workbook.Close();
            excelApp.Quit();

            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();



            return System.Windows.Forms.Application.StartupPath + "\\Rp_ventas_Sucursal_" + sucursal.Trim() + "_" + dia.ToString("dd MMMM yyyy") + ".xlsx";
        }

        public static dynamic AsignaValorCelda(dynamic Celda1, dynamic Celda2) 
        {
            dynamic ValorCelda = 0;
            try
            {
                ValorCelda = Celda1 / Celda2;

                if (Double.IsNaN(ValorCelda))
                {
                    return 0;

                    //throw new DivideByZeroException();
                }

                return ValorCelda;
            }
            catch (DivideByZeroException)
            {
                return ValorCelda;
            }
        }


        private string Existencias_dias_ventas(DateTime dia)
        {
            string str;
            try
            {
                #region Resumen
                SqlConnection cn = conexion.conectar("BDIntegrador");
                SqlCommand sqlCommand = new SqlCommand()
                {
                    Connection = cn,
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "MI_ReporteDiasVenta",
                    CommandTimeout = 0
                };

                sqlCommand.Parameters.Clear();
                sqlCommand.Parameters.AddWithValue("@p_tipo", 1);

                SqlDataAdapter da = new SqlDataAdapter(sqlCommand);
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt;


                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
                Excel.Worksheet hoja = g_Workbook.Worksheets.Add(Before: g_Workbook.Worksheets[1]);
                hoja.Name = "Resumen";

                hoja.Columns["A"].ColumnWidth = 0;
                hoja.Columns["B"].ColumnWidth = 18;
                hoja.Columns["C"].ColumnWidth = 9;

                Microsoft.Office.Interop.Excel.Range rango;
                if (DG1.Rows.Count > 0)
                {
                    hoja.Cells[1, 2] = "Resúmen de Ventas al " + DateTime.Now.ToString("dd/MM/yyyy") + " (Según venta del día " + DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy") + ")";
                    hoja.Cells[1, 2].Font.FontStyle = "Bold";
                    hoja.Cells[2, 2] = "PRODUCTOS BASICOS y NO RESURTIBLES";
                    hoja.Cells[2, 2].Font.FontStyle = "Bold";

                    hoja.Cells[4, 2] = "Tipo";
                    hoja.Cells[4, 3] = "Unidades";
                    hoja.Cells[4, 4] = "Importe";

                    hoja.Cells[8, 2] = "Totales";
                    hoja.Cells[8, 3].Formula = "=SUM(C5:C7)";
                    hoja.Cells[8, 4] = "=SUM(D5:D7)";

                    hoja.Range["B4", "D8"].Borders.LineStyle = XlLineStyle.xlContinuous;


                    hoja.Cells[5, 3].NumberFormat = "#,###,##0";
                    hoja.Cells[6, 3].NumberFormat = "#,###,##0";
                    hoja.Cells[7, 3].NumberFormat = "#,###,##0";
                    hoja.Cells[8, 3].NumberFormat = "#,###,##0";

                    hoja.Cells[5, 4].NumberFormat = "$#,###,##0.00";
                    hoja.Cells[6, 4].NumberFormat = "$#,###,##0.00";
                    hoja.Cells[7, 4].NumberFormat = "$#,###,##0.00";
                    hoja.Cells[8, 4].NumberFormat = "$#,###,##0.00";

                    hoja.Cells[4, 2].Font.FontStyle = "Bold";
                    hoja.Cells[4, 3].Font.FontStyle = "Bold";
                    hoja.Cells[4, 4].Font.FontStyle = "Bold";

                    hoja.Cells[8, 2].Font.FontStyle = "Bold";
                    hoja.Cells[8, 3].Font.FontStyle = "Bold";
                    hoja.Cells[8, 4].Font.FontStyle = "Bold";
                    //Tipo
                    hoja.Cells[5, 2].Value = DG1.Rows[0].Cells[0].Value;
                    hoja.Cells[6, 2].Value = DG1.Rows[1].Cells[0].Value;
                    hoja.Cells[7, 2].Value = DG1.Rows[2].Cells[0].Value;
                    //Unidades
                    hoja.Cells[5, 3].Value = DG1.Rows[0].Cells[1].Value;
                    hoja.Cells[6, 3].Value = DG1.Rows[1].Cells[1].Value;
                    hoja.Cells[7, 3].Value = DG1.Rows[2].Cells[1].Value;
                    //Importe
                    hoja.Cells[5, 4].Value = DG1.Rows[0].Cells[2].Value;
                    hoja.Cells[6, 4].Value = DG1.Rows[1].Cells[2].Value;
                    hoja.Cells[7, 4].Value = DG1.Rows[2].Cells[2].Value;
                    rango = (Range)hoja.get_Range("A4", "D8");
                    rango.Columns.AutoFit();


                }
                #endregion

                #region SUPER BASICOS
                /*limpiar grid para cargar informacion  de  segunda consulta*/
                DG1.DataSource = null;
                DG1.Rows.Clear();
                DG1.Columns.Clear();

                SqlConnection cn2 = conexion.conectar("BDIntegrador");
                SqlCommand sqlCommand2 = new SqlCommand()
                {
                    Connection = cn2,
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "MI_ReporteDiasVenta",
                    CommandTimeout = 0
                };

                sqlCommand2.Parameters.Clear();
                sqlCommand2.Parameters.AddWithValue("@p_tipo", 2);

                SqlDataAdapter da2 = new SqlDataAdapter(sqlCommand2);
                System.Data.DataTable dt2 = new System.Data.DataTable();
                da2.Fill(dt2);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt2;
                Microsoft.Office.Interop.Excel.Range rango2;
                if (DG1.Rows.Count > 0)
                {

                    Excel.Worksheet hoja2 = g_Workbook.Worksheets.Add(After: g_Workbook.Worksheets[1]);

                    hoja2.Name = "SUPER BASICOS";
                    hoja2.Columns["A"].ColumnWidth = 13;
                    hoja2.Columns["B"].ColumnWidth = 18;
                    hoja2.Columns["C"].ColumnWidth = 9;
                    hoja2.Columns["P"].ColumnWidth = 18;

                    hoja2.Cells[1, 7] = " Reporte de existencia en dias venta  al " + DateTime.Now.ToString("dd/MM/yyyy") + " (Según venta del día " + DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy") + ")";
                    hoja2.Cells[2, 7] = "PRODUCTOS SUPERBASICOS DE HOGAR Y BELLEZA";
                    hoja2.Cells[3, 7] = "Ordenado por producto de mayor venta a menor venta";

                    hoja2.Cells[1, 7].Font.FontStyle = "Bold";
                    hoja2.Cells[2, 7].Font.FontStyle = "Bold";
                    hoja2.Cells[3, 7].Font.FontStyle = "Bold";

                    rango2 = (Range)hoja2.get_Range("G1", "L1");
                    rango2.Select();
                    rango2.Merge();
                    rango2.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango2 = (Range)hoja2.get_Range("G2", "L2");
                    rango2.Select();
                    rango2.Merge();
                    rango2.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango2 = (Range)hoja2.get_Range("G3", "L3");
                    rango2.Select();
                    rango2.Merge();
                    rango2.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    //Recorremos el DataGridView para colocar encabezados de columnas
                    for (int i = 0; i < DG1.ColumnCount; i++)
                    {
                        hoja2.Cells[5, i + 1] = DG1.Columns[i].HeaderText;
                        hoja2.Cells[5, i + 1].Font.FontStyle = "Bold";

                    }
                    int Fila = 6;
                    //Recorremos el DataGridView rellenando la hoja de trabajo 
                    foreach (DataGridViewRow item in DG1.Rows)
                    {
                        for (int i = 0; i < DG1.ColumnCount; i++)
                        {
                            hoja2.Cells[Fila, i + 1].Value = item.Cells[i].Value;
                        }
                        Fila++;
                    }

                    //System.Data.DataTable DT = (System.Data.DataTable)DG1.DataSource;


                    //for (int i = 0; i < DG1.ColumnCount; i++)
                    //{
                    //    hoja2.Cells[Fila, i + 1].Value = DT.AsEnumerable().Sum(r => r.Field<decimal>(dt.Columns[i]));
                    //}

                    int fila;
                    fila = DG1.Rows.Count + 5;
                    hoja2.Range["E6", "E" + fila.ToString()].NumberFormat = "0";
                    hoja2.Range["L6", "L" + fila.ToString()].NumberFormat = "0";

                    hoja2.Range["D6", "D" + fila.ToString()].NumberFormat = "0";

                    hoja2.Range["A5", "FA" + fila.ToString()].Borders.LineStyle = XlLineStyle.xlContinuous;
                    rango2 = (Range)hoja2.get_Range("A5", "FA" + DG1.Rows.Count.ToString());

                    //Final
                    hoja2.Cells[fila + 1, 1] = "TOTALES";
                    hoja2.Cells[fila + 1, 2].Font.FontStyle = "Bold";

                    //Importe Venta Mensual
                    hoja2.Cells[fila + 1, 15].Formula = "=SUM(O6:O" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 15].NumberFormat = "#,###,##0.00";
                    hoja2.Range["O6", "O" + fila.ToString()].NumberFormat = "#,###,##0.00";
                    //Contribución Vta Mensual
                    hoja2.Cells[fila + 1, 16].Formula = "=SUM(P6:P" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 16].NumberFormat = "$#,###,##0.00";
                    hoja2.Range["P6", "P" + fila.ToString()].NumberFormat = "$#,###,##0.00";
                    //Contribución Vta Mensual
                    hoja2.Cells[fila + 1, 17].Formula = "=SUM(Q6:Q" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 17].NumberFormat = "$#,###,##0.00";
                    hoja2.Range["Q6", "Q" + fila.ToString()].NumberFormat = "$#,###,##0.00";
                    //Unidades Vta Bimestral
                    hoja2.Cells[fila + 1, 18].Formula = "=SUM(R6:R" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 18].NumberFormat = "#,###,##0.00";
                    hoja2.Range["R6", "R" + fila.ToString()].NumberFormat = "#,###,##0.00";
                    //Importe Vta Bimestral
                    hoja2.Cells[fila + 1, 19].Formula = "=SUM(S6:S" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 19].NumberFormat = "$#,###,##0.00";
                    hoja2.Range["S6", "S" + fila.ToString()].NumberFormat = "$#,###,##0";
                    //Contribución Vta Bimestral
                    hoja2.Cells[fila + 1, 20].Formula = "=SUM(T6:T" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 20].NumberFormat = "$#,###,##0.00";
                    //Piezas Última Recepción
                    hoja2.Cells[fila + 1, 22].Formula = "=SUM(V6:V" + fila.ToString() + ")";
                    hoja2.Range["V6", "V" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 22].NumberFormat = "#,###,##0.00";

                    //Piezas vendidas desde U.R.
                    hoja2.Cells[fila + 1, 23].Formula = "=SUM(W6:W" + fila.ToString() + ")";
                    hoja2.Range["W6", "W" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 23].NumberFormat = "#,###,##0.00";
                    //%
                    hoja2.Range["X6", "X" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Existencia CEDIS
                    hoja2.Cells[fila + 1, 24].Formula = "=SUM(X6:X" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 24].NumberFormat = "#,###,##0.00";
                    hoja2.Range["X6", "X" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Existencia en Tiendas
                    hoja2.Cells[fila + 1, 25].Formula = "=SUM(Y6:Y" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 25].NumberFormat = "#,###,##0.00";
                    hoja2.Range["Y6", "Y" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja2.Cells[fila + 1, 26].Formula = "=SUM(Z6:Z" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 26].NumberFormat = "#,###,##0.00";
                    hoja2.Range["Z6", "Z" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja2.Cells[fila + 1, 27].Formula = "=SUM(AA6:AA" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 27].NumberFormat = "#,###,##0.00";
                    hoja2.Range["AA6", "AA" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja2.Cells[fila + 1, 28].Formula = "=SUM(AB6:AB" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 28].NumberFormat = "#,###,##0.00";
                    hoja2.Range["AB6", "AB" + fila.ToString()].NumberFormat = "#,###,##0";
                    int ad;
                    ad = fila + 1;
                    //Dias Venta Mensual
                    hoja2.Cells[fila + 1, 29].NumberFormat = "#,###,##0.00";
                    hoja2.Cells[fila + 1, 29].Formula = "=AB" + ad.ToString() + " / O" + ad.ToString();

                    hoja2.Cells[fila + 1, 30].Formula = "=SUM(AD6:AD" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 30].NumberFormat = "#,###,##0.00";
                    hoja2.Range["AD6", "AD" + fila.ToString()].NumberFormat = "#,###,##0";
                    //cod_prod
                    hoja2.Cells[fila + 1, 31].Formula = "=AB" + ad.ToString() + "/(R" + ad.ToString() + "/30)";
                    hoja2.Cells[fila + 1, 31].NumberFormat = "#,###,##0.00";
                    hoja2.Range["AD6", "AD" + fila.ToString()].NumberFormat = "#,###,##0";

                    // COD_PRO 32

                    hoja2.Range["AF6", "AF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 32].Formula = "=SUM(AF6:AF" + fila.ToString() + ")";
                    //
                    hoja2.Range["AG6", "AG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 33].Formula = "=SUM(AG6:AG" + fila.ToString() + ")";
                    //
                    hoja2.Cells[fila + 1, 34].NumberFormat = "#,###,##0.00";
                    hoja2.Range["AH6", "AH" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Dias Venta 3
                    hoja2.Cells[fila + 1, 34].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 32].Value, hoja2.Cells[fila + 1, 33].Value);
                    /**/
                    hoja2.Range["AI6", "AI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 35].Formula = "=SUM(AI6:AI" + fila.ToString() + ")";

                    hoja2.Range["AJ6", "AJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 36].Formula = "=SUM(AJ6:AJ" + fila.ToString() + ")";
                    //Dias Venta 4
                    hoja2.Cells[fila + 1, 37].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 35].Value, hoja2.Cells[fila + 1, 36].Value);
                    /**/
                    hoja2.Range["AL6", "AL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 38].Formula = "=SUM(AL6:AL" + fila.ToString() + ")";
                    /**/
                    hoja2.Range["AM6", "AM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 39].Formula = "=SUM(AM6:AM" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 40].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 38].Value, hoja2.Cells[fila + 1, 39].Value);

                    /**/
                    hoja2.Range["AO6", "AO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 41].Formula = "=SUM(AO6:AO" + fila.ToString() + ")";
                    /**/
                    hoja2.Range["AP6", "AP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 42].Formula = "=SUM(AP6:AP" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 43].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 41].Value, hoja2.Cells[fila + 1, 42].Value);
                    /**/
                    hoja2.Range["AR6", "AR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 44].Formula = "=SUM(AR6:AR" + fila.ToString() + ")";

                    hoja2.Range["AS6", "AS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 45].Formula = "=SUM(AS6:AS" + fila.ToString() + ")";

                    hoja2.Cells[fila + 1, 46].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 44].Value, hoja2.Cells[fila + 1, 45].Value);
                    /**/
                    hoja2.Range["AU6", "AU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 47].Formula = "=SUM(AU6:AU" + fila.ToString() + ")";

                    hoja2.Range["AV6", "AV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 48].Formula = "=SUM(AV6:AV" + fila.ToString() + ")";

                    hoja2.Cells[fila + 1, 49].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 47].Value, hoja2.Cells[fila + 1, 48].Value);
                    /**/
                    hoja2.Range["AX6", "AX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 50].Formula = "=SUM(AX6:AX" + fila.ToString() + ")";
                    hoja2.Range["AY6", "AY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 51].Formula = "=SUM(AY6:AY" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 52].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 50].Value, hoja2.Cells[fila + 1, 51].Value);
                    /**/
                    hoja2.Range["BA6", "BA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 53].Formula = "=SUM(BA6:BA" + fila.ToString() + ")";

                    hoja2.Range["BB6", "BB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 54].Formula = "=SUM(BB6:BB" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 55].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 53].Value, hoja2.Cells[fila + 1, 54].Value);
                    /**/
                    hoja2.Range["BD6", "BD" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 56].Formula = "=SUM(BD6:BD" + fila.ToString() + ")";
                    hoja2.Range["BE6", "BE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 57].Formula = "=SUM(BE6:BE" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 58].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 56].Value, hoja2.Cells[fila + 1, 57].Value);
                    /**/
                    hoja2.Range["BG6", "BG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 59].Formula = "=SUM(BG6:BG" + fila.ToString() + ")";
                    hoja2.Range["BH6", "BH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 60].Formula = "=SUM(BH6:BH" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 61].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 59].Value, hoja2.Cells[fila + 1, 60].Value);
                    /**/
                    hoja2.Range["BJ6", "BJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 62].Formula = "=SUM(BJ6:BJ" + fila.ToString() + ")";
                    hoja2.Range["BK6", "BK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 63].Formula = "=SUM(BK6:BK" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 64].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 62].Value, hoja2.Cells[fila + 1, 63].Value);
                    /**/
                    hoja2.Cells[fila + 1, 65].Formula = "=SUM(BM6:BM" + fila.ToString() + ")";
                    hoja2.Range["BM6", "BM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 66].Formula = "=SUM(BN6:BN" + fila.ToString() + ")";
                    hoja2.Range["BN6", "BN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 67].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 65].Value, hoja2.Cells[fila + 1, 66].Value);
                    /**/
                    hoja2.Range["BP6", "BP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 68].Formula = "=SUM(BP6:BP" + fila.ToString() + ")";
                    hoja2.Range["BQ6", "BQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 69].Formula = "=SUM(BQ6:BQ" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 70].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 68].Value, hoja2.Cells[fila + 1, 69].Value);
                    /**/
                    hoja2.Range["BS6", "BS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 71].Formula = "=SUM(BS6:BS" + fila.ToString() + ")";
                    hoja2.Range["BT6", "BT" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 72].Formula = "=SUM(BT6:BT" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 73].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 71].Value, hoja2.Cells[fila + 1, 72].Value);
                    /**/
                    hoja2.Range["BV6", "BV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 74].Formula = "=SUM(BV6:BV" + fila.ToString() + ")";
                    hoja2.Range["BW6", "BW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 75].Formula = "=SUM(BW6:BW" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 76].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 74].Value, hoja2.Cells[fila + 1, 75].Value);
                    /**/
                    hoja2.Range["BY6", "BY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 77].Formula = "=SUM(BY6:BY" + fila.ToString() + ")";
                    hoja2.Range["BZ6", "BZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 78].Formula = "=SUM(BZ6:BZ" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 79].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 77].Value, hoja2.Cells[fila + 1, 78].Value);
                    /**/
                    hoja2.Range["CB6", "CB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 80].Formula = "=SUM(CB6:CB" + fila.ToString() + ")";
                    hoja2.Range["CC6", "CC" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 81].Formula = "=SUM(CC6:CC" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 82].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 80].Value, hoja2.Cells[fila + 1, 81].Value);
                    /**/
                    hoja2.Range["CE6", "CE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 83].Formula = "=SUM(CE6:CE" + fila.ToString() + ")";
                    hoja2.Range["CF6", "CF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 84].Formula = "=SUM(CF6:CF" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 85].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 83].Value, hoja2.Cells[fila + 1, 84].Value);
                    /**/
                    hoja2.Range["CH6", "CH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 86].Formula = "=SUM(CH6:CH" + fila.ToString() + ")";
                    hoja2.Range["CI6", "CI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 87].Formula = "=SUM(CI6:CI" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 88].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 86].Value, hoja2.Cells[fila + 1, 87].Value);
                    /**/
                    hoja2.Range["CK6", "CK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 89].Formula = "=SUM(CK6:CK" + fila.ToString() + ")";
                    hoja2.Range["CL6", "CL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 90].Formula = "=SUM(CL6:CL" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 91].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 89].Value, hoja2.Cells[fila + 1, 90].Value);
                    /**/
                    hoja2.Range["CN6", "CN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 92].Formula = "=SUM(CN6:CN" + fila.ToString() + ")";
                    hoja2.Range["CO6", "CO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 93].Formula = "=SUM(CO6:CO" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 94].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 92].Value, hoja2.Cells[fila + 1, 93].Value);
                    /**/
                    hoja2.Range["CQ6", "CQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 95].Formula = "=SUM(CQ6:CQ" + fila.ToString() + ")";
                    hoja2.Range["CR6", "CR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 96].Formula = "=SUM(CR6:CR" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 97].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 95].Value, hoja2.Cells[fila + 1, 96].Value);
                    /**/
                    hoja2.Range["CT6", "CT" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 98].Formula = "=SUM(CT6:CT" + fila.ToString() + ")";
                    hoja2.Range["CU6", "CU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 99].Formula = "=SUM(CU6:CU" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 100].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 98].Value, hoja2.Cells[fila + 1, 99].Value);
                    /**/
                    hoja2.Range["CW6", "CW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 101].Formula = "=SUM(CW6:CW" + fila.ToString() + ")";
                    hoja2.Range["CX6", "CX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 102].Formula = "=SUM(CX6:CX" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 103].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 101].Value, hoja2.Cells[fila + 1, 102].Value);
                    //
                    hoja2.Range["CZ6", "CZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 104].Formula = "=SUM(CZ6:CZ" + fila.ToString() + ")";
                    hoja2.Range["DA6", "DA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 105].Formula = "=SUM(DA6:DA" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 106].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 104].Value, hoja2.Cells[fila + 1, 105].Value);
                    /**/
                    hoja2.Range["DC6", "DC" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 107].Formula = "=SUM(DC6:DC" + fila.ToString() + ")";
                    hoja2.Range["DD6", "DD" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 108].Formula = "=SUM(DD6:DD" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 109].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 107].Value, hoja2.Cells[fila + 1, 108].Value);
                    //
                    hoja2.Range["DF6", "DF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 110].Formula = "=SUM(DF6:DF" + fila.ToString() + ")";
                    hoja2.Range["DG6", "DG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 111].Formula = "=SUM(DG6:DG" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 112].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 110].Value, hoja2.Cells[fila + 1, 111].Value);
                    /**/
                    hoja2.Range["DI6", "DI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 113].Formula = "=SUM(DI6:DI" + fila.ToString() + ")";    /**/
                    hoja2.Range["DJ6", "DJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 114].Formula = "=SUM(DJ6:DJ" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 115].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 113].Value, hoja2.Cells[fila + 1, 114].Value);
                    /**/
                    hoja2.Range["DL6", "DL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 116].Formula = "=SUM(DL6:DL" + fila.ToString() + ")";
                    hoja2.Range["DM6", "DM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 117].Formula = "=SUM(DM6:DM" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 118].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 116].Value, hoja2.Cells[fila + 1, 117].Value);
                    /**/
                    hoja2.Range["DO6", "DO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 119].Formula = "=SUM(DO6:DO" + fila.ToString() + ")";
                    hoja2.Range["DP6", "DP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 120].Formula = "=SUM(DP6:DP" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 121].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 119].Value, hoja2.Cells[fila + 1, 120].Value);
                    /**/
                    hoja2.Range["DR6", "DR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 122].Formula = "=SUM(DR6:DR" + fila.ToString() + ")";
                    hoja2.Range["DS6", "DS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 123].Formula = "=SUM(DS6:DS" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 124].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 122].Value, hoja2.Cells[fila + 1, 123].Value);
                    /**/
                    hoja2.Range["DU6", "DU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 125].Formula = "=SUM(DU6:DU" + fila.ToString() + ")"; /**/
                    hoja2.Range["DV6", "DV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 126].Formula = "=SUM(DV6:DV" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 127].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 125].Value, hoja2.Cells[fila + 1, 126].Value);
                    /**/
                    hoja2.Range["DX6", "DX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 128].Formula = "=SUM(DX6:DX" + fila.ToString() + ")";
                    hoja2.Range["DY6", "DY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 129].Formula = "=SUM(DY6:DY" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 130].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 128].Value, hoja2.Cells[fila + 1, 129].Value);
                    /**/
                    hoja2.Range["EA6", "EA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 131].Formula = "=SUM(EA6:EA" + fila.ToString() + ")";
                    hoja2.Range["EB6", "EB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 132].Formula = "=SUM(EB6:EB" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 133].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 131].Value, hoja2.Cells[fila + 1, 132].Value);
                    /**/
                    hoja2.Range["ED6", "ED" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 134].Formula = "=SUM(ED6:ED" + fila.ToString() + ")";
                    hoja2.Range["EE6", "EE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 135].Formula = "=SUM(EE6:EE" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 136].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 134].Value, hoja2.Cells[fila + 1, 135].Value);
                    /**/
                    hoja2.Range["EG6", "EG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 137].Formula = "=SUM(EG6:EG" + fila.ToString() + ")";
                    hoja2.Range["EH6", "EH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 138].Formula = "=SUM(EH6:EH" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 139].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 137].Value, hoja2.Cells[fila + 1, 138].Value);
                    /**/
                    hoja2.Range["EJ6", "EJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 140].Formula = "=SUM(EJ6:EJ" + fila.ToString() + ")";
                    hoja2.Range["EK6", "EK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 141].Formula = "=SUM(EK6:EK" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 142].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 140].Value, hoja2.Cells[fila + 1, 141].Value);
                    /**/
                    hoja2.Range["EM6", "EM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 143].Formula = "=SUM(EM6:EM" + fila.ToString() + ")";
                    hoja2.Range["EN6", "EN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 144].Formula = "=SUM(EN6:EN" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 145].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 143].Value, hoja2.Cells[fila + 1, 144].Value);
                    /**/
                    hoja2.Range["EP6", "EP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 146].Formula = "=SUM(EP6:EP" + fila.ToString() + ")";
                    hoja2.Range["EQ6", "EQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 147].Formula = "=SUM(EQ6:EQ" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 148].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 146].Value, hoja2.Cells[fila + 1, 147].Value);
                    /**/
                    hoja2.Range["ES6", "ES" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 149].Formula = "=SUM(ES6:ES" + fila.ToString() + ")";
                    hoja2.Range["ET6", "ET" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 150].Formula = "=SUM(ET6:ET" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 151].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 149].Value, hoja2.Cells[fila + 1, 150].Value);
                    /**/
                    hoja2.Range["EV6", "EV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 152].Formula = "=SUM(EV6:EV" + fila.ToString() + ")";
                    hoja2.Range["EW6", "EW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 153].Formula = "=SUM(EW6:EW" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 154].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 152].Value, hoja2.Cells[fila + 1, 153].Value);
                    /**/
                    hoja2.Range["EY6", "EY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 155].Formula = "=SUM(EY6:EY" + fila.ToString() + ")";
                    hoja2.Range["EZ6", "EZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja2.Cells[fila + 1, 156].Formula = "=SUM(EZ6:EZ" + fila.ToString() + ")";
                    hoja2.Cells[fila + 1, 157].Value = AsignaValorCelda(hoja2.Cells[fila + 1, 155].Value, hoja2.Cells[fila + 1, 156].Value);

                    /**/
                    hoja2.Range["A" + ad, "FA" + ad.ToString()].Borders.LineStyle = XlLineStyle.xlContinuous;

                    hoja2.Range["A" + ad, "FA" + ad.ToString()].Font.FontStyle = "Bold";
                    hoja2.Range["C6", "C" + fila.ToString()].NumberFormat = "0";

                    rango2.Columns.AutoFit();
                }
                #endregion

                #region BASICOS
                SqlConnection cn3 = conexion.conectar("BDIntegrador");
                SqlCommand sqlCommand3 = new SqlCommand()
                {
                    Connection = cn3,
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "MI_ReporteDiasVenta",
                    CommandTimeout = 0
                };

                sqlCommand3.Parameters.Clear();
                sqlCommand3.Parameters.AddWithValue("@p_tipo", 3);

                SqlDataAdapter da3 = new SqlDataAdapter(sqlCommand3);
                System.Data.DataTable dt3 = new System.Data.DataTable();
                da3.Fill(dt3);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt3;

                Microsoft.Office.Interop.Excel.Range rango3;
                if (DG1.Rows.Count > 0)
                {
                    Excel.Worksheet hoja3 = g_Workbook.Worksheets.Add(After: g_Workbook.Worksheets[2]);

                    hoja3.Name = "BASICOS";
                    hoja3.Columns["A"].ColumnWidth = 0;
                    hoja3.Columns["B"].ColumnWidth = 18;
                    hoja3.Columns["C"].ColumnWidth = 9;

                    hoja3.Cells[1, 5] = "Reporte de existencia en dias venta al " + DateTime.Now.ToString("dd/MM/yyyy") + " (Según venta del día " + DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy") + ")";
                    hoja3.Cells[2, 5] = "PRODUCTOS BASICOS";
                    hoja3.Cells[3, 5] = "Ordenado por producto de mayor venta a menor venta";

                    hoja3.Cells[1, 5].Font.FontStyle = "Bold";
                    hoja3.Cells[2, 5].Font.FontStyle = "Bold";
                    hoja3.Cells[3, 5].Font.FontStyle = "Bold";

                    rango3 = (Range)hoja3.get_Range("E1", "J1");
                    rango3.Select();
                    rango3.Merge();
                    rango3.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango3 = (Range)hoja3.get_Range("E2", "J2");
                    rango3.Select();
                    rango3.Merge();
                    rango3.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango3 = (Range)hoja3.get_Range("E3", "J3");
                    rango3.Select();
                    rango3.Merge();
                    rango3.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //Recorremos el DataGridView para colocar encabezados de columnas

                    for (int i = 0; i < DG1.ColumnCount; i++)
                    {
                        hoja3.Cells[5, i + 1] = DG1.Columns[i].HeaderText;
                        hoja3.Cells[5, i + 1].Font.FontStyle = "Bold";

                    }
                    int Fila = 6;
                    //Recorremos el DataGridView rellenando la hoja de trabajo 
                    foreach (DataGridViewRow item in DG1.Rows)
                    {
                        for (int i = 0; i < DG1.ColumnCount; i++)
                        {
                            hoja3.Cells[Fila, i + 1].Value = item.Cells[i].Value;
                        }
                        Fila++;
                    }

                    int fila;
                    fila = DG1.Rows.Count + 5;
                    hoja3.Range["E6", "E" + fila.ToString()].NumberFormat = "0";
                    hoja3.Range["L6", "L" + fila.ToString()].NumberFormat = "0";

                    hoja3.Range["D6", "D" + fila.ToString()].NumberFormat = "0";


                    hoja3.Range["A5", "FA" + fila.ToString()].Borders.LineStyle = XlLineStyle.xlContinuous;
                    rango2 = (Range)hoja3.get_Range("A5", "FA" + DG1.Rows.Count.ToString());

                    //Final
                    hoja3.Cells[fila + 1, 1] = "TOTALES";
                    hoja3.Cells[fila + 1, 2].Font.FontStyle = "Bold";

                    //Importe Venta Mensual
                    hoja3.Cells[fila + 1, 15].Formula = "=SUM(O6:O" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 15].NumberFormat = "#,###,##0.00";
                    hoja3.Range["O6", "O" + fila.ToString()].NumberFormat = "#,###,##0.00";
                    //Contribución Vta Mensual
                    hoja3.Cells[fila + 1, 16].Formula = "=SUM(P6:P" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 16].NumberFormat = "$#,###,##0.00";
                    hoja3.Range["P6", "P" + fila.ToString()].NumberFormat = "$#,###,##0.00";
                    //Contribución Vta Mensual
                    hoja3.Cells[fila + 1, 17].Formula = "=SUM(Q6:Q" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 17].NumberFormat = "$#,###,##0.00";
                    hoja3.Range["Q6", "Q" + fila.ToString()].NumberFormat = "$#,###,##0.00";
                    //Unidades Vta Bimestral
                    hoja3.Cells[fila + 1, 18].Formula = "=SUM(R6:R" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 18].NumberFormat = "#,###,##0.00";
                    hoja3.Range["R6", "R" + fila.ToString()].NumberFormat = "#,###,##0.00";
                    //Importe Vta Bimestral
                    hoja3.Cells[fila + 1, 19].Formula = "=SUM(S6:S" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 19].NumberFormat = "$#,###,##0.00";
                    hoja3.Range["S6", "S" + fila.ToString()].NumberFormat = "$#,###,##0";
                    //Contribución Vta Bimestral
                    hoja3.Cells[fila + 1, 20].Formula = "=SUM(T6:T" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 20].NumberFormat = "$#,###,##0.00";
                    //Piezas Última Recepción
                    hoja3.Cells[fila + 1, 22].Formula = "=SUM(V6:V" + fila.ToString() + ")";
                    hoja3.Range["V6", "V" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 22].NumberFormat = "#,###,##0.00";

                    //Piezas vendidas desde U.R.
                    hoja3.Cells[fila + 1, 23].Formula = "=SUM(W6:W" + fila.ToString() + ")";
                    hoja3.Range["W6", "W" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 23].NumberFormat = "#,###,##0.00";
                    //%
                    hoja3.Range["X6", "X" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Existencia CEDIS
                    hoja3.Cells[fila + 1, 24].Formula = "=SUM(X6:X" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 24].NumberFormat = "#,###,##0.00";
                    hoja3.Range["X6", "X" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Existencia en Tiendas
                    hoja3.Cells[fila + 1, 25].Formula = "=SUM(Y6:Y" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 25].NumberFormat = "#,###,##0.00";
                    hoja3.Range["Y6", "Y" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja3.Cells[fila + 1, 26].Formula = "=SUM(Z6:Z" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 26].NumberFormat = "#,###,##0.00";
                    hoja3.Range["Z6", "Z" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja3.Cells[fila + 1, 27].Formula = "=SUM(AA6:AA" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 27].NumberFormat = "#,###,##0.00";
                    hoja3.Range["AA6", "AA" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja3.Cells[fila + 1, 28].Formula = "=SUM(AB6:AB" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 28].NumberFormat = "#,###,##0.00";
                    hoja3.Range["AB6", "AB" + fila.ToString()].NumberFormat = "#,###,##0";
                    int ad;
                    ad = fila + 1;
                    //Dias Venta Mensual
                    hoja3.Cells[fila + 1, 29].NumberFormat = "#,###,##0.00";
                    hoja3.Cells[fila + 1, 29].Formula = "=AB" + ad.ToString() + " / O" + ad.ToString();

                    hoja3.Cells[fila + 1, 30].Formula = "=SUM(AD6:AD" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 30].NumberFormat = "#,###,##0.00";
                    hoja3.Range["AD6", "AD" + fila.ToString()].NumberFormat = "#,###,##0";
                    //cod_prod
                    hoja3.Cells[fila + 1, 31].Formula = "=AB" + ad.ToString() + "/(R" + ad.ToString() + "/30)";
                    hoja3.Cells[fila + 1, 31].NumberFormat = "#,###,##0.00";
                    hoja3.Range["AD6", "AD" + fila.ToString()].NumberFormat = "#,###,##0";

                    // COD_PRO 32

                    hoja3.Range["AF6", "AF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 32].Formula = "=SUM(AF6:AF" + fila.ToString() + ")";
                    //
                    hoja3.Range["AG6", "AG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 33].Formula = "=SUM(AG6:AG" + fila.ToString() + ")";
                    //
                    hoja3.Cells[fila + 1, 34].NumberFormat = "#,###,##0.00";
                    hoja3.Range["AH6", "AH" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Dias Venta 3
                    hoja3.Cells[fila + 1, 34].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 32].Value, hoja3.Cells[fila + 1, 33].Value);
                    /**/
                    hoja3.Range["AI6", "AI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 35].Formula = "=SUM(AI6:AI" + fila.ToString() + ")";

                    hoja3.Range["AJ6", "AJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 36].Formula = "=SUM(AJ6:AJ" + fila.ToString() + ")";
                    //Dias Venta 4
                    hoja3.Cells[fila + 1, 37].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 35].Value, hoja3.Cells[fila + 1, 36].Value);
                    /**/
                    hoja3.Range["AL6", "AL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 38].Formula = "=SUM(AL6:AL" + fila.ToString() + ")";
                    /**/
                    hoja3.Range["AM6", "AM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 39].Formula = "=SUM(AM6:AM" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 40].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 38].Value, hoja3.Cells[fila + 1, 39].Value);

                    /**/
                    hoja3.Range["AO6", "AO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 41].Formula = "=SUM(AO6:AO" + fila.ToString() + ")";
                    /**/
                    hoja3.Range["AP6", "AP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 42].Formula = "=SUM(AP6:AP" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 43].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 41].Value, hoja3.Cells[fila + 1, 42].Value);
                    /**/
                    hoja3.Range["AR6", "AR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 44].Formula = "=SUM(AR6:AR" + fila.ToString() + ")";

                    hoja3.Range["AS6", "AS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 45].Formula = "=SUM(AS6:AS" + fila.ToString() + ")";

                    hoja3.Cells[fila + 1, 46].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 44].Value, hoja3.Cells[fila + 1, 45].Value);
                    /**/
                    hoja3.Range["AU6", "AU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 47].Formula = "=SUM(AU6:AU" + fila.ToString() + ")";

                    hoja3.Range["AV6", "AV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 48].Formula = "=SUM(AV6:AV" + fila.ToString() + ")";

                    hoja3.Cells[fila + 1, 49].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 47].Value, hoja3.Cells[fila + 1, 48].Value);
                    /**/
                    hoja3.Range["AX6", "AX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 50].Formula = "=SUM(AX6:AX" + fila.ToString() + ")";
                    hoja3.Range["AY6", "AY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 51].Formula = "=SUM(AY6:AY" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 52].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 50].Value, hoja3.Cells[fila + 1, 51].Value);
                    /**/
                    hoja3.Range["BA6", "BA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 53].Formula = "=SUM(BA6:BA" + fila.ToString() + ")";

                    hoja3.Range["BB6", "BB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 54].Formula = "=SUM(BB6:BB" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 55].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 53].Value, hoja3.Cells[fila + 1, 54].Value);
                    /**/
                    hoja3.Range["BD6", "BD" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 56].Formula = "=SUM(BD6:BD" + fila.ToString() + ")";
                    hoja3.Range["BE6", "BE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 57].Formula = "=SUM(BE6:BE" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 58].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 56].Value, hoja3.Cells[fila + 1, 57].Value);
                    /**/
                    hoja3.Range["BG6", "BG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 59].Formula = "=SUM(BG6:BG" + fila.ToString() + ")";
                    hoja3.Range["BH6", "BH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 60].Formula = "=SUM(BH6:BH" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 61].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 59].Value, hoja3.Cells[fila + 1, 60].Value);
                    /**/
                    hoja3.Range["BJ6", "BJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 62].Formula = "=SUM(BJ6:BJ" + fila.ToString() + ")";
                    hoja3.Range["BK6", "BK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 63].Formula = "=SUM(BK6:BK" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 64].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 62].Value, hoja3.Cells[fila + 1, 63].Value);
                    /**/
                    hoja3.Cells[fila + 1, 65].Formula = "=SUM(BM6:BM" + fila.ToString() + ")";
                    hoja3.Range["BM6", "BM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 66].Formula = "=SUM(BN6:BN" + fila.ToString() + ")";
                    hoja3.Range["BN6", "BN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 67].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 65].Value, hoja3.Cells[fila + 1, 66].Value);
                    /**/
                    hoja3.Range["BP6", "BP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 68].Formula = "=SUM(BP6:BP" + fila.ToString() + ")";
                    hoja3.Range["BQ6", "BQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 69].Formula = "=SUM(BQ6:BQ" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 70].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 68].Value, hoja3.Cells[fila + 1, 69].Value);
                    /**/
                    hoja3.Range["BS6", "BS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 71].Formula = "=SUM(BS6:BS" + fila.ToString() + ")";
                    hoja3.Range["BT6", "BT" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 72].Formula = "=SUM(BT6:BT" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 73].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 71].Value, hoja3.Cells[fila + 1, 72].Value);
                    /**/
                    hoja3.Range["BV6", "BV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 74].Formula = "=SUM(BV6:BV" + fila.ToString() + ")";
                    hoja3.Range["BW6", "BW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 75].Formula = "=SUM(BW6:BW" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 76].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 74].Value, hoja3.Cells[fila + 1, 75].Value);
                    /**/
                    hoja3.Range["BY6", "BY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 77].Formula = "=SUM(BY6:BY" + fila.ToString() + ")";
                    hoja3.Range["BZ6", "BZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 78].Formula = "=SUM(BZ6:BZ" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 79].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 77].Value, hoja3.Cells[fila + 1, 78].Value);
                    /**/
                    hoja3.Range["CB6", "CB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 80].Formula = "=SUM(CB6:CB" + fila.ToString() + ")";
                    hoja3.Range["CC6", "CC" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 81].Formula = "=SUM(CC6:CC" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 82].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 80].Value, hoja3.Cells[fila + 1, 81].Value);
                    /**/
                    hoja3.Range["CE6", "CE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 83].Formula = "=SUM(CE6:CE" + fila.ToString() + ")";
                    hoja3.Range["CF6", "CF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 84].Formula = "=SUM(CF6:CF" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 85].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 83].Value, hoja3.Cells[fila + 1, 84].Value);
                    /**/
                    hoja3.Range["CH6", "CH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 86].Formula = "=SUM(CH6:CH" + fila.ToString() + ")";
                    hoja3.Range["CI6", "CI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 87].Formula = "=SUM(CI6:CI" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 88].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 86].Value, hoja3.Cells[fila + 1, 87].Value);
                    /**/
                    hoja3.Range["CK6", "CK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 89].Formula = "=SUM(CK6:CK" + fila.ToString() + ")";
                    hoja3.Range["CL6", "CL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 90].Formula = "=SUM(CL6:CL" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 91].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 89].Value, hoja3.Cells[fila + 1, 90].Value);
                    /**/
                    hoja3.Range["CN6", "CN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 92].Formula = "=SUM(CN6:CN" + fila.ToString() + ")";
                    hoja3.Range["CO6", "CO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 93].Formula = "=SUM(CO6:CO" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 94].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 92].Value, hoja3.Cells[fila + 1, 93].Value);
                    /**/
                    hoja3.Range["CQ6", "CQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 95].Formula = "=SUM(CQ6:CQ" + fila.ToString() + ")";
                    hoja3.Range["CR6", "CR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 96].Formula = "=SUM(CR6:CR" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 97].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 95].Value, hoja3.Cells[fila + 1, 96].Value);
                    /**/
                    hoja3.Range["CT6", "CT" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 98].Formula = "=SUM(CT6:CT" + fila.ToString() + ")";
                    hoja3.Range["CU6", "CU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 99].Formula = "=SUM(CU6:CU" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 100].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 98].Value, hoja3.Cells[fila + 1, 99].Value);
                    /**/
                    hoja3.Range["CW6", "CW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 101].Formula = "=SUM(CW6:CW" + fila.ToString() + ")";
                    hoja3.Range["CX6", "CX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 102].Formula = "=SUM(CX6:CX" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 103].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 101].Value, hoja3.Cells[fila + 1, 102].Value);
                    //
                    hoja3.Range["CZ6", "CZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 104].Formula = "=SUM(CZ6:CZ" + fila.ToString() + ")";
                    hoja3.Range["DA6", "DA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 105].Formula = "=SUM(DA6:DA" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 106].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 104].Value, hoja3.Cells[fila + 1, 105].Value);
                    /**/
                    hoja3.Range["DC6", "DC" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 107].Formula = "=SUM(DC6:DC" + fila.ToString() + ")";
                    hoja3.Range["DD6", "DD" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 108].Formula = "=SUM(DD6:DD" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 109].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 107].Value, hoja3.Cells[fila + 1, 108].Value);
                    //
                    hoja3.Range["DF6", "DF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 110].Formula = "=SUM(DF6:DF" + fila.ToString() + ")";
                    hoja3.Range["DG6", "DG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 111].Formula = "=SUM(DG6:DG" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 112].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 110].Value, hoja3.Cells[fila + 1, 111].Value);
                    /**/
                    hoja3.Range["DI6", "DI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 113].Formula = "=SUM(DI6:DI" + fila.ToString() + ")";    /**/
                    hoja3.Range["DJ6", "DJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 114].Formula = "=SUM(DJ6:DJ" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 115].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 113].Value, hoja3.Cells[fila + 1, 114].Value);
                    /**/
                    hoja3.Range["DL6", "DL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 116].Formula = "=SUM(DL6:DL" + fila.ToString() + ")";
                    hoja3.Range["DM6", "DM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 117].Formula = "=SUM(DM6:DM" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 118].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 116].Value, hoja3.Cells[fila + 1, 117].Value);
                    /**/
                    hoja3.Range["DO6", "DO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 119].Formula = "=SUM(DO6:DO" + fila.ToString() + ")";
                    hoja3.Range["DP6", "DP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 120].Formula = "=SUM(DP6:DP" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 121].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 119].Value, hoja3.Cells[fila + 1, 120].Value);
                    /**/
                    hoja3.Range["DR6", "DR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 122].Formula = "=SUM(DR6:DR" + fila.ToString() + ")";
                    hoja3.Range["DS6", "DS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 123].Formula = "=SUM(DS6:DS" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 124].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 122].Value, hoja3.Cells[fila + 1, 123].Value);
                    /**/
                    hoja3.Range["DU6", "DU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 125].Formula = "=SUM(DU6:DU" + fila.ToString() + ")"; /**/
                    hoja3.Range["DV6", "DV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 126].Formula = "=SUM(DV6:DV" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 127].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 125].Value, hoja3.Cells[fila + 1, 126].Value);
                    /**/
                    hoja3.Range["DX6", "DX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 128].Formula = "=SUM(DX6:DX" + fila.ToString() + ")";
                    hoja3.Range["DY6", "DY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 129].Formula = "=SUM(DY6:DY" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 130].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 128].Value, hoja3.Cells[fila + 1, 129].Value);
                    /**/
                    hoja3.Range["EA6", "EA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 131].Formula = "=SUM(EA6:EA" + fila.ToString() + ")";
                    hoja3.Range["EB6", "EB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 132].Formula = "=SUM(EB6:EB" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 133].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 131].Value, hoja3.Cells[fila + 1, 132].Value);
                    /**/
                    hoja3.Range["ED6", "ED" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 134].Formula = "=SUM(ED6:ED" + fila.ToString() + ")";
                    hoja3.Range["EE6", "EE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 135].Formula = "=SUM(EE6:EE" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 136].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 134].Value, hoja3.Cells[fila + 1, 135].Value);
                    /**/
                    hoja3.Range["EG6", "EG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 137].Formula = "=SUM(EG6:EG" + fila.ToString() + ")";
                    hoja3.Range["EH6", "EH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 138].Formula = "=SUM(EH6:EH" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 139].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 137].Value, hoja3.Cells[fila + 1, 138].Value);
                    /**/
                    hoja3.Range["EJ6", "EJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 140].Formula = "=SUM(EJ6:EJ" + fila.ToString() + ")";
                    hoja3.Range["EK6", "EK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 141].Formula = "=SUM(EK6:EK" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 142].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 140].Value, hoja3.Cells[fila + 1, 141].Value);
                    /**/
                    hoja3.Range["EM6", "EM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 143].Formula = "=SUM(EM6:EM" + fila.ToString() + ")";
                    hoja3.Range["EN6", "EN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 144].Formula = "=SUM(EN6:EN" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 145].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 143].Value, hoja3.Cells[fila + 1, 144].Value);
                    /**/
                    hoja3.Range["EP6", "EP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 146].Formula = "=SUM(EP6:EP" + fila.ToString() + ")";
                    hoja3.Range["EQ6", "EQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 147].Formula = "=SUM(EQ6:EQ" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 148].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 146].Value, hoja3.Cells[fila + 1, 147].Value);
                    /**/
                    hoja3.Range["ES6", "ES" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 149].Formula = "=SUM(ES6:ES" + fila.ToString() + ")";
                    hoja3.Range["ET6", "ET" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 150].Formula = "=SUM(ET6:ET" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 151].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 149].Value, hoja3.Cells[fila + 1, 150].Value);
                    /**/
                    hoja3.Range["EV6", "EV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 152].Formula = "=SUM(EV6:EV" + fila.ToString() + ")";
                    hoja3.Range["EW6", "EW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 153].Formula = "=SUM(EW6:EW" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 154].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 152].Value, hoja3.Cells[fila + 1, 153].Value);
                    /**/
                    hoja3.Range["EY6", "EY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 155].Formula = "=SUM(EY6:EY" + fila.ToString() + ")";
                    hoja3.Range["EZ6", "EZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja3.Cells[fila + 1, 156].Formula = "=SUM(EZ6:EZ" + fila.ToString() + ")";
                    hoja3.Cells[fila + 1, 157].Value = AsignaValorCelda(hoja3.Cells[fila + 1, 155].Value, hoja3.Cells[fila + 1, 156].Value);

                    /**/
                    hoja3.Range["A" + ad, "FA" + ad.ToString()].Borders.LineStyle = XlLineStyle.xlContinuous;

                    hoja3.Range["A" + ad, "FA" + ad.ToString()].Font.FontStyle = "Bold";
                    hoja3.Range["C6", "C" + fila.ToString()].NumberFormat = "0";

                    rango2.Columns.AutoFit();

                }
                #endregion

                #region NO RESULTIBLES 
                /*nueva consulta NO RESURTIBLES*/
                SqlConnection cn4 = conexion.conectar("BDIntegrador");
                SqlCommand sqlCommand4 = new SqlCommand()
                {
                    Connection = cn4,
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "MI_ReporteDiasVenta",
                    CommandTimeout = 0
                };

                sqlCommand4.Parameters.Clear();
                sqlCommand4.Parameters.AddWithValue("@p_tipo", 4);

                SqlDataAdapter da4 = new SqlDataAdapter(sqlCommand4);
                System.Data.DataTable dt4 = new System.Data.DataTable();
                da4.Fill(dt4);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt4;

                Microsoft.Office.Interop.Excel.Range rango4;
                if (DG1.Rows.Count > 0)
                {
                    Excel.Worksheet hoja4 = g_Workbook.Worksheets.Add(After: g_Workbook.Worksheets[3]);
                    hoja4.Name = "NO RESURTIBLES";
                    hoja4.Columns["A"].ColumnWidth = 0;
                    hoja4.Columns["B"].ColumnWidth = 18;
                    hoja4.Columns["C"].ColumnWidth = 9;

                    hoja4.Cells[1, 7] = "Reporte de existencia en dias venta al " + DateTime.Now.ToString("dd/MM/yyyy") + " (Según venta del día " + DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy") + ")";
                    hoja4.Cells[2, 7] = "PRODUCTOS NO RESURTIBLES";
                    hoja4.Cells[3, 7] = "Ordenado por producto de mayor venta a menor venta";

                    hoja4.Cells[1, 7].Font.FontStyle = "Bold";
                    hoja4.Cells[2, 7].Font.FontStyle = "Bold";
                    hoja4.Cells[3, 7].Font.FontStyle = "Bold";

                    rango4 = (Range)hoja4.get_Range("G1", "M1");
                    rango4.Select();
                    rango4.Merge();
                    rango4.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango4 = (Range)hoja4.get_Range("G2", "M2");
                    rango4.Select();
                    rango4.Merge();
                    rango4.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rango4 = (Range)hoja4.get_Range("G3", "M3");
                    rango4.Select();
                    rango4.Merge();
                    rango4.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    //Recorremos el DataGridView para colocar encabezados de columnas

                    for (int i = 0; i < DG1.ColumnCount; i++)
                    {
                        hoja4.Cells[5, i + 1] = DG1.Columns[i].HeaderText;
                        hoja4.Cells[5, i + 1].Font.FontStyle = "Bold";

                    }
                    int Fila = 6;
                    //Recorremos el DataGridView rellenando la hoja de trabajo 
                    foreach (DataGridViewRow item in DG1.Rows)
                    {
                        for (int i = 0; i < DG1.ColumnCount; i++)
                        {
                            hoja4.Cells[Fila, i + 1].Value = item.Cells[i].Value;
                        }
                        Fila++;
                    }

                    int fila;
                    fila = DG1.Rows.Count + 5;
                    hoja4.Range["E6", "E" + fila.ToString()].NumberFormat = "0";
                    hoja4.Range["L6", "L" + fila.ToString()].NumberFormat = "0";

                    hoja4.Range["D6", "D" + fila.ToString()].NumberFormat = "0";


                    hoja4.Range["A5", "FA" + fila.ToString()].Borders.LineStyle = XlLineStyle.xlContinuous;
                    rango2 = (Range)hoja4.get_Range("A5", "FA" + DG1.Rows.Count.ToString());

                    //Final
                    hoja4.Cells[fila + 1, 1] = "TOTALES";
                    hoja4.Cells[fila + 1, 2].Font.FontStyle = "Bold";

                    //Importe Venta Mensual
                    hoja4.Cells[fila + 1, 15].Formula = "=SUM(O6:O" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 15].NumberFormat = "#,###,##0.00";
                    hoja4.Range["O6", "O" + fila.ToString()].NumberFormat = "#,###,##0.00";
                    //Contribución Vta Mensual
                    hoja4.Cells[fila + 1, 16].Formula = "=SUM(P6:P" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 16].NumberFormat = "$#,###,##0.00";
                    hoja4.Range["P6", "P" + fila.ToString()].NumberFormat = "$#,###,##0.00";
                    //Contribución Vta Mensual
                    hoja4.Cells[fila + 1, 17].Formula = "=SUM(Q6:Q" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 17].NumberFormat = "$#,###,##0.00";
                    hoja4.Range["Q6", "Q" + fila.ToString()].NumberFormat = "$#,###,##0.00";
                    //Unidades Vta Bimestral
                    hoja4.Cells[fila + 1, 18].Formula = "=SUM(R6:R" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 18].NumberFormat = "#,###,##0.00";
                    hoja4.Range["R6", "R" + fila.ToString()].NumberFormat = "#,###,##0.00";
                    //Importe Vta Bimestral
                    hoja4.Cells[fila + 1, 19].Formula = "=SUM(S6:S" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 19].NumberFormat = "$#,###,##0.00";
                    hoja4.Range["S6", "S" + fila.ToString()].NumberFormat = "$#,###,##0";
                    //Contribución Vta Bimestral
                    hoja4.Cells[fila + 1, 20].Formula = "=SUM(T6:T" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 20].NumberFormat = "$#,###,##0.00";
                    //Piezas Última Recepción
                    hoja4.Cells[fila + 1, 22].Formula = "=SUM(V6:V" + fila.ToString() + ")";
                    hoja4.Range["V6", "V" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 22].NumberFormat = "#,###,##0.00";

                    //Piezas vendidas desde U.R.
                    hoja4.Cells[fila + 1, 23].Formula = "=SUM(W6:W" + fila.ToString() + ")";
                    hoja4.Range["W6", "W" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 23].NumberFormat = "#,###,##0.00";
                    //%
                    hoja4.Range["X6", "X" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Existencia CEDIS
                    hoja4.Cells[fila + 1, 24].Formula = "=SUM(X6:X" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 24].NumberFormat = "#,###,##0.00";
                    hoja4.Range["X6", "X" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Existencia en Tiendas
                    hoja4.Cells[fila + 1, 25].Formula = "=SUM(Y6:Y" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 25].NumberFormat = "#,###,##0.00";
                    hoja4.Range["Y6", "Y" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja4.Cells[fila + 1, 26].Formula = "=SUM(Z6:Z" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 26].NumberFormat = "#,###,##0.00";
                    hoja4.Range["Z6", "Z" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja4.Cells[fila + 1, 27].Formula = "=SUM(AA6:AA" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 27].NumberFormat = "#,###,##0.00";
                    hoja4.Range["AA6", "AA" + fila.ToString()].NumberFormat = "#,###,##0";
                    //
                    hoja4.Cells[fila + 1, 28].Formula = "=SUM(AB6:AB" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 28].NumberFormat = "#,###,##0.00";
                    hoja4.Range["AB6", "AB" + fila.ToString()].NumberFormat = "#,###,##0";
                    int ad;
                    ad = fila + 1;
                    //Dias Venta Mensual
                    hoja4.Cells[fila + 1, 29].NumberFormat = "#,###,##0.00";
                    hoja4.Cells[fila + 1, 29].Formula = "=AB" + ad.ToString() + " / O" + ad.ToString();

                    hoja4.Cells[fila + 1, 30].Formula = "=SUM(AD6:AD" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 30].NumberFormat = "#,###,##0.00";
                    hoja4.Range["AD6", "AD" + fila.ToString()].NumberFormat = "#,###,##0";
                    //cod_prod
                    hoja4.Cells[fila + 1, 31].Formula = "=AB" + ad.ToString() + "/(R" + ad.ToString() + "/30)";
                    hoja4.Cells[fila + 1, 31].NumberFormat = "#,###,##0.00";
                    hoja4.Range["AD6", "AD" + fila.ToString()].NumberFormat = "#,###,##0";

                    // COD_PRO 32

                    hoja4.Range["AF6", "AF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 32].Formula = "=SUM(AF6:AF" + fila.ToString() + ")";
                    //
                    hoja4.Range["AG6", "AG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 33].Formula = "=SUM(AG6:AG" + fila.ToString() + ")";
                    //
                    hoja4.Cells[fila + 1, 34].NumberFormat = "#,###,##0.00";
                    hoja4.Range["AH6", "AH" + fila.ToString()].NumberFormat = "#,###,##0";
                    //Dias Venta 3
                    hoja4.Cells[fila + 1, 34].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 32].Value, hoja4.Cells[fila + 1, 33].Value);
                    /**/
                    hoja4.Range["AI6", "AI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 35].Formula = "=SUM(AI6:AI" + fila.ToString() + ")";

                    hoja4.Range["AJ6", "AJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 36].Formula = "=SUM(AJ6:AJ" + fila.ToString() + ")";
                    //Dias Venta 4
                    hoja4.Cells[fila + 1, 37].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 35].Value, hoja4.Cells[fila + 1, 36].Value);
                    /**/
                    hoja4.Range["AL6", "AL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 38].Formula = "=SUM(AL6:AL" + fila.ToString() + ")";
                    /**/
                    hoja4.Range["AM6", "AM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 39].Formula = "=SUM(AM6:AM" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 40].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 38].Value, hoja4.Cells[fila + 1, 39].Value);

                    /**/
                    hoja4.Range["AO6", "AO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 41].Formula = "=SUM(AO6:AO" + fila.ToString() + ")";
                    /**/
                    hoja4.Range["AP6", "AP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 42].Formula = "=SUM(AP6:AP" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 43].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 41].Value, hoja4.Cells[fila + 1, 42].Value);
                    /**/
                    hoja4.Range["AR6", "AR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 44].Formula = "=SUM(AR6:AR" + fila.ToString() + ")";

                    hoja4.Range["AS6", "AS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 45].Formula = "=SUM(AS6:AS" + fila.ToString() + ")";

                    hoja4.Cells[fila + 1, 46].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 44].Value, hoja4.Cells[fila + 1, 45].Value);
                    /**/
                    hoja4.Range["AU6", "AU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 47].Formula = "=SUM(AU6:AU" + fila.ToString() + ")";

                    hoja4.Range["AV6", "AV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 48].Formula = "=SUM(AV6:AV" + fila.ToString() + ")";

                    hoja4.Cells[fila + 1, 49].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 47].Value, hoja4.Cells[fila + 1, 48].Value);
                    /**/
                    hoja4.Range["AX6", "AX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 50].Formula = "=SUM(AX6:AX" + fila.ToString() + ")";
                    hoja4.Range["AY6", "AY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 51].Formula = "=SUM(AY6:AY" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 52].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 50].Value, hoja4.Cells[fila + 1, 51].Value);
                    /**/
                    hoja4.Range["BA6", "BA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 53].Formula = "=SUM(BA6:BA" + fila.ToString() + ")";

                    hoja4.Range["BB6", "BB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 54].Formula = "=SUM(BB6:BB" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 55].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 53].Value, hoja4.Cells[fila + 1, 54].Value);
                    /**/
                    hoja4.Range["BD6", "BD" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 56].Formula = "=SUM(BD6:BD" + fila.ToString() + ")";
                    hoja4.Range["BE6", "BE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 57].Formula = "=SUM(BE6:BE" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 58].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 56].Value, hoja4.Cells[fila + 1, 57].Value);
                    /**/
                    hoja4.Range["BG6", "BG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 59].Formula = "=SUM(BG6:BG" + fila.ToString() + ")";
                    hoja4.Range["BH6", "BH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 60].Formula = "=SUM(BH6:BH" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 61].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 59].Value, hoja4.Cells[fila + 1, 60].Value);
                    /**/
                    hoja4.Range["BJ6", "BJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 62].Formula = "=SUM(BJ6:BJ" + fila.ToString() + ")";
                    hoja4.Range["BK6", "BK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 63].Formula = "=SUM(BK6:BK" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 64].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 62].Value, hoja4.Cells[fila + 1, 63].Value);
                    /**/
                    hoja4.Cells[fila + 1, 65].Formula = "=SUM(BM6:BM" + fila.ToString() + ")";
                    hoja4.Range["BM6", "BM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 66].Formula = "=SUM(BN6:BN" + fila.ToString() + ")";
                    hoja4.Range["BN6", "BN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 67].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 65].Value, hoja4.Cells[fila + 1, 66].Value);
                    /**/
                    hoja4.Range["BP6", "BP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 68].Formula = "=SUM(BP6:BP" + fila.ToString() + ")";
                    hoja4.Range["BQ6", "BQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 69].Formula = "=SUM(BQ6:BQ" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 70].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 68].Value, hoja4.Cells[fila + 1, 69].Value);
                    /**/
                    hoja4.Range["BS6", "BS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 71].Formula = "=SUM(BS6:BS" + fila.ToString() + ")";
                    hoja4.Range["BT6", "BT" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 72].Formula = "=SUM(BT6:BT" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 73].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 71].Value, hoja4.Cells[fila + 1, 72].Value);
                    /**/
                    hoja4.Range["BV6", "BV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 74].Formula = "=SUM(BV6:BV" + fila.ToString() + ")";
                    hoja4.Range["BW6", "BW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 75].Formula = "=SUM(BW6:BW" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 76].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 74].Value, hoja4.Cells[fila + 1, 75].Value);
                    /**/
                    hoja4.Range["BY6", "BY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 77].Formula = "=SUM(BY6:BY" + fila.ToString() + ")";
                    hoja4.Range["BZ6", "BZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 78].Formula = "=SUM(BZ6:BZ" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 79].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 77].Value, hoja4.Cells[fila + 1, 78].Value);
                    /**/
                    hoja4.Range["CB6", "CB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 80].Formula = "=SUM(CB6:CB" + fila.ToString() + ")";
                    hoja4.Range["CC6", "CC" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 81].Formula = "=SUM(CC6:CC" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 82].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 80].Value, hoja4.Cells[fila + 1, 81].Value);
                    /**/
                    hoja4.Range["CE6", "CE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 83].Formula = "=SUM(CE6:CE" + fila.ToString() + ")";
                    hoja4.Range["CF6", "CF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 84].Formula = "=SUM(CF6:CF" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 85].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 83].Value, hoja4.Cells[fila + 1, 84].Value);
                    /**/
                    hoja4.Range["CH6", "CH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 86].Formula = "=SUM(CH6:CH" + fila.ToString() + ")";
                    hoja4.Range["CI6", "CI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 87].Formula = "=SUM(CI6:CI" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 88].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 86].Value, hoja4.Cells[fila + 1, 87].Value);
                    /**/
                    hoja4.Range["CK6", "CK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 89].Formula = "=SUM(CK6:CK" + fila.ToString() + ")";
                    hoja4.Range["CL6", "CL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 90].Formula = "=SUM(CL6:CL" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 91].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 89].Value, hoja4.Cells[fila + 1, 90].Value);
                    /**/
                    hoja4.Range["CN6", "CN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 92].Formula = "=SUM(CN6:CN" + fila.ToString() + ")";
                    hoja4.Range["CO6", "CO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 93].Formula = "=SUM(CO6:CO" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 94].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 92].Value, hoja4.Cells[fila + 1, 93].Value);
                    /**/
                    hoja4.Range["CQ6", "CQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 95].Formula = "=SUM(CQ6:CQ" + fila.ToString() + ")";
                    hoja4.Range["CR6", "CR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 96].Formula = "=SUM(CR6:CR" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 97].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 95].Value, hoja4.Cells[fila + 1, 96].Value);
                    /**/
                    hoja4.Range["CT6", "CT" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 98].Formula = "=SUM(CT6:CT" + fila.ToString() + ")";
                    hoja4.Range["CU6", "CU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 99].Formula = "=SUM(CU6:CU" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 100].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 98].Value, hoja4.Cells[fila + 1, 99].Value);
                    /**/
                    hoja4.Range["CW6", "CW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 101].Formula = "=SUM(CW6:CW" + fila.ToString() + ")";
                    hoja4.Range["CX6", "CX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 102].Formula = "=SUM(CX6:CX" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 103].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 101].Value, hoja4.Cells[fila + 1, 102].Value);
                    //
                    hoja4.Range["CZ6", "CZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 104].Formula = "=SUM(CZ6:CZ" + fila.ToString() + ")";
                    hoja4.Range["DA6", "DA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 105].Formula = "=SUM(DA6:DA" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 106].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 104].Value, hoja4.Cells[fila + 1, 105].Value);
                    /**/
                    hoja4.Range["DC6", "DC" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 107].Formula = "=SUM(DC6:DC" + fila.ToString() + ")";
                    hoja4.Range["DD6", "DD" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 108].Formula = "=SUM(DD6:DD" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 109].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 107].Value, hoja4.Cells[fila + 1, 108].Value);
                    //
                    hoja4.Range["DF6", "DF" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 110].Formula = "=SUM(DF6:DF" + fila.ToString() + ")";
                    hoja4.Range["DG6", "DG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 111].Formula = "=SUM(DG6:DG" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 112].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 110].Value, hoja4.Cells[fila + 1, 111].Value);
                    /**/
                    hoja4.Range["DI6", "DI" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 113].Formula = "=SUM(DI6:DI" + fila.ToString() + ")";    /**/
                    hoja4.Range["DJ6", "DJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 114].Formula = "=SUM(DJ6:DJ" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 115].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 113].Value, hoja4.Cells[fila + 1, 114].Value);
                    /**/
                    hoja4.Range["DL6", "DL" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 116].Formula = "=SUM(DL6:DL" + fila.ToString() + ")";
                    hoja4.Range["DM6", "DM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 117].Formula = "=SUM(DM6:DM" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 118].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 116].Value, hoja4.Cells[fila + 1, 117].Value);
                    /**/
                    hoja4.Range["DO6", "DO" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 119].Formula = "=SUM(DO6:DO" + fila.ToString() + ")";
                    hoja4.Range["DP6", "DP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 120].Formula = "=SUM(DP6:DP" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 121].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 119].Value, hoja4.Cells[fila + 1, 120].Value);
                    /**/
                    hoja4.Range["DR6", "DR" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 122].Formula = "=SUM(DR6:DR" + fila.ToString() + ")";
                    hoja4.Range["DS6", "DS" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 123].Formula = "=SUM(DS6:DS" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 124].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 122].Value, hoja4.Cells[fila + 1, 123].Value);
                    /**/
                    hoja4.Range["DU6", "DU" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 125].Formula = "=SUM(DU6:DU" + fila.ToString() + ")"; /**/
                    hoja4.Range["DV6", "DV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 126].Formula = "=SUM(DV6:DV" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 127].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 125].Value, hoja4.Cells[fila + 1, 126].Value);
                    /**/
                    hoja4.Range["DX6", "DX" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 128].Formula = "=SUM(DX6:DX" + fila.ToString() + ")";
                    hoja4.Range["DY6", "DY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 129].Formula = "=SUM(DY6:DY" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 130].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 128].Value, hoja4.Cells[fila + 1, 129].Value);
                    /**/
                    hoja4.Range["EA6", "EA" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 131].Formula = "=SUM(EA6:EA" + fila.ToString() + ")";
                    hoja4.Range["EB6", "EB" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 132].Formula = "=SUM(EB6:EB" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 133].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 131].Value, hoja4.Cells[fila + 1, 132].Value);
                    /**/
                    hoja4.Range["ED6", "ED" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 134].Formula = "=SUM(ED6:ED" + fila.ToString() + ")";
                    hoja4.Range["EE6", "EE" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 135].Formula = "=SUM(EE6:EE" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 136].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 134].Value, hoja4.Cells[fila + 1, 135].Value);
                    /**/
                    hoja4.Range["EG6", "EG" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 137].Formula = "=SUM(EG6:EG" + fila.ToString() + ")";
                    hoja4.Range["EH6", "EH" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 138].Formula = "=SUM(EH6:EH" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 139].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 137].Value, hoja4.Cells[fila + 1, 138].Value);
                    /**/
                    hoja4.Range["EJ6", "EJ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 140].Formula = "=SUM(EJ6:EJ" + fila.ToString() + ")";
                    hoja4.Range["EK6", "EK" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 141].Formula = "=SUM(EK6:EK" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 142].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 140].Value, hoja4.Cells[fila + 1, 141].Value);
                    /**/
                    hoja4.Range["EM6", "EM" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 143].Formula = "=SUM(EM6:EM" + fila.ToString() + ")";
                    hoja4.Range["EN6", "EN" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 144].Formula = "=SUM(EN6:EN" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 145].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 143].Value, hoja4.Cells[fila + 1, 144].Value);
                    /**/
                    hoja4.Range["EP6", "EP" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 146].Formula = "=SUM(EP6:EP" + fila.ToString() + ")";
                    hoja4.Range["EQ6", "EQ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 147].Formula = "=SUM(EQ6:EQ" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 148].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 146].Value, hoja4.Cells[fila + 1, 147].Value);
                    /**/
                    hoja4.Range["ES6", "ES" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 149].Formula = "=SUM(ES6:ES" + fila.ToString() + ")";
                    hoja4.Range["ET6", "ET" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 150].Formula = "=SUM(ET6:ET" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 151].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 149].Value, hoja4.Cells[fila + 1, 150].Value);
                    /**/
                    hoja4.Range["EV6", "EV" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 152].Formula = "=SUM(EV6:EV" + fila.ToString() + ")";
                    hoja4.Range["EW6", "EW" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 153].Formula = "=SUM(EW6:EW" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 154].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 152].Value, hoja4.Cells[fila + 1, 153].Value);
                    /**/
                    hoja4.Range["EY6", "EY" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 155].Formula = "=SUM(EY6:EY" + fila.ToString() + ")";
                    hoja4.Range["EZ6", "EZ" + fila.ToString()].NumberFormat = "#,###,##0";
                    hoja4.Cells[fila + 1, 156].Formula = "=SUM(EZ6:EZ" + fila.ToString() + ")";
                    hoja4.Cells[fila + 1, 157].Value = AsignaValorCelda(hoja4.Cells[fila + 1, 155].Value, hoja4.Cells[fila + 1, 156].Value);

                    /**/
                    hoja4.Range["A" + ad, "FA" + ad.ToString()].Borders.LineStyle = XlLineStyle.xlContinuous;

                    hoja4.Range["A" + ad, "FA" + ad.ToString()].Font.FontStyle = "Bold";
                    hoja4.Range["C6", "C" + fila.ToString()].NumberFormat = "0";

                    rango2.Columns.AutoFit();

                }
                if (g_Workbook.Worksheets.Count > 4)
                {
                    g_Workbook.Worksheets[5].Delete();
                    if (g_Workbook.Worksheets.Count > 5)
                    {
                        g_Workbook.Worksheets[6].Delete();
                    }
                }
                #endregion

                if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Existencias_Dia_Venta_" + dia.ToString("dd MMMM yyyy") + ".xlsx"))
                {
                    File.Delete(System.Windows.Forms.Application.StartupPath + "\\Existencias_Dia_Venta_" + dia.ToString("dd MMMM yyyy") + ".xlsx");
                }
                g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Existencias_Dia_Venta_" + dia.ToString("dd MMMM yyyy") + ".xlsx");
                g_Workbook.Close();
                excelApp.Quit();
                str = System.Windows.Forms.Application.StartupPath + "\\Existencias_Dia_Venta_" + dia.ToString("dd MMMM yyyy") + ".xlsx";
            }
            catch (Exception e)
            {
                lblError.Text = "Hubo en error: " + e.Message.ToString();
                return str = "";
            }
            return str;
        }

        private string Ventas_Articulos_30(DateTime dt, string cod_estab, string nombre)
        {
            DateTime fi_anterior; DateTime ff_anterior; DateTime fi_actual; DateTime ff_actual;

            fi_actual = dt.AddDays(-30);
            ff_actual = dt.AddDays(-1);
            fi_anterior = (dt.AddDays(-7)).AddDays(-30);
            ff_anterior = (dt.AddDays(-7)).AddDays(-1);

            Microsoft.Office.Interop.Excel.Range rango;

            SqlConnection cn1 = conexion.conectar("BDIntegrador");
            SqlCommand sqlCommand1 = new SqlCommand()
            {
                Connection = cn1,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Ventas_Articulos_30Dias",
                CommandTimeout = 0
            };
            //ok
            sqlCommand1.Parameters.Clear();
            sqlCommand1.Parameters.AddWithValue("@p_dt", dt.ToString("yyyyMMdd"));

            SqlDataAdapter da1 = new SqlDataAdapter(sqlCommand1);
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt1;
            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            DataObject dataObj = DG1.GetClipboardContent();
            excelApp.Visible = false;
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);
            /*tenemos 3 hojas iniciales + la agregada 4*/

            hoja = (Worksheet)g_Workbook.Sheets.get_Item(1);
            hoja.Activate();
            hoja.Name = "VENTAS ARTICULOS 30 DIAS";
            hoja.Columns["A"].ColumnWidth = 3;

            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);

                rango = (Range)hoja.get_Range("F1", "I1");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[1, 6] = "VENTAS DE ARTICULOS A 30 DIAS";
                hoja.Cells[1, 6].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("F2", "I2");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 6] = "PRODUCTOS BASICOS y SUPER BASICOS ";
                hoja.Cells[2, 6].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B4", "C4");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                rango = (Range)hoja.get_Range("D4", "H4");
                rango.Select();
                rango.Merge();
                hoja.Cells[4, 4] = "Semana Anterior  De " + fi_anterior.ToString("yyyy-MM-dd") + " a " + ff_anterior.ToString("yyyy-MM-dd");
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("I4", "N4");
                rango.Select();
                rango.Merge();
                hoja.Cells[4, 9] = "Semana Actual De " + fi_actual.ToString("yyyy-MM-dd") + " a " + ff_actual.ToString("yyyy-MM-dd");
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";

                rango = (Range)hoja.Cells[6, 1];
                rango.Select();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                hoja.Cells[4, 3].Font.FontStyle = "Bold";

                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[5, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[5, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[5, j + 2].Font.FontStyle = "Bold";
                }

                int registros;
                registros = DG1.Rows.Count + 4;
                rango = (Range)hoja.get_Range("K6", "N" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();

                hoja.Columns["B"].ColumnWidth = 35;
                hoja.Columns["C"].ColumnWidth = 15;

                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                rango = (Range)hoja.get_Range("B6", "N" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                hoja.Range["D6", "D" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["E6", "E" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["F6", "F" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["G6", "G" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["H6", "H" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["I6", "I" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["J6", "J" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["K6", "K" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["L6", "L" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["M6", "M" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["N6", "N" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
            }

            int posicion = 1;
            g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);
            posicion = posicion + 1;

            SqlConnection cn3 = conexion.conectar("BDIntegrador");
            SqlCommand sqlCommand3 = new SqlCommand()
            {
                Connection = cn3,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Ventas_Articulos_30DPromocion",
                CommandTimeout = 0
            };

            sqlCommand3.Parameters.Clear();
            sqlCommand3.Parameters.AddWithValue("@p_dt", dt.ToString("yyyyMMdd"));
            sqlCommand3.Parameters.AddWithValue("@p_estab", cod_estab.Trim());

            SqlDataAdapter da3 = new SqlDataAdapter(sqlCommand3);
            System.Data.DataTable dt3 = new System.Data.DataTable();
            da3.Fill(dt3);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt3;
            this.DG1.SelectAll();

            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);

                hoja = (Worksheet)g_Workbook.Sheets.get_Item(posicion);
                hoja.Activate();
                hoja.Name = nombre;

                /*INSERTAR DATOS DE SUCURSAL*/

                rango = (Range)hoja.get_Range("K2", "U2");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 11] = "VENTAS DE ARTICULOS A 30 DIAS";
                hoja.Cells[2, 11].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("K3", "U3");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[3, 11] = "PRODUCTOS BASICOS Y SUPER BASICOS";
                hoja.Cells[3, 11].Font.FontStyle = "Bold";


                //celda ''
                rango = (Range)hoja.get_Range("B4", "J4");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                rango = (Range)hoja.get_Range("K4", "O4");
                rango.Select();
                rango.Merge();

                hoja.Cells[4, 11] = "Semana Anterior  De " + fi_anterior.ToString("yyyy-MM-dd") + " a " + ff_anterior.ToString("yyyy-MM-dd");
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("P4", "U4");
                rango.Select();
                rango.Merge();
                hoja.Cells[4, 16] = "Semana Actual De " + fi_actual.ToString("yyyy-MM-dd") + " a " + ff_actual.ToString("yyyy-MM-dd");
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("V4", "W4");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";


                hoja.Columns["A"].ColumnWidth = 3;

                rango = (Range)hoja.Cells[6, 1];
                rango.Select();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                hoja.Cells[4, 3].Font.FontStyle = "Bold";

                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[5, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[5, 2 + j].Font.FontStyle = "Bold";
                    hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[5, j + 2].Font.FontStyle = "Bold";
                }
                int registros = 0;
                registros = DG1.Rows.Count + 4;
                rango = (Range)hoja.get_Range("B6", "W" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();

                hoja.Columns["B"].ColumnWidth = 24;
                hoja.Columns["C"].ColumnWidth = 6.5;
                hoja.Columns["D"].ColumnWidth = 48;
                hoja.Columns["E"].ColumnWidth = 15;
                hoja.Columns["F"].ColumnWidth = 15;
                hoja.Columns["G"].ColumnWidth = 15;
                hoja.Columns["H"].ColumnWidth = 21.3;
                hoja.Columns["I"].ColumnWidth = 7;
                hoja.Columns["J"].ColumnWidth = 27;

                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                //double dias_venta;
                ///*Aplicar cambios a la columna piezas a transferir/ pedir */
                //for (int j = 0; j < DG1.Rows.Count-1; j++)
                //{
                //    hoja.Cells[6 + j, 22] = "=IF(T" + (6 + j) + " > 90, (Q" + (6 + j )+ " - ((R" + (6 + j) + " / 30) * 90)) * -1, IF(T" + (6 + j) + " < 60, ((R" + (6 + j )+ " / 30) * 60) - Q" + (6 + j) + ",))";
                //    hoja.Cells[6 + j, 23] = "=ROUND(IF(V" + (6 + j) + "<0,(Q" + (6 + j) + "+V" + (6 + j) + ")/R" + (6 + j) + "*30,(V" + (6 + j) + "+Q" + (6 + j) + ")/R" + (6 + j) + "*30),0)";


                //    if (hoja.Cells[6 + j, 20].value != null) {

                //        dias_venta = hoja.Cells[6 + j, 20].value;
                //        if (dias_venta < 60 || dias_venta > 90)
                //        {
                //            hoja.Cells[6 + j, 20].Interior.Color = System.Drawing.Color.Red;
                //        }
                //    }

                //}

                hoja.Range["K6", "K" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["L6", "L" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["M6", "M" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["N6", "N" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["O6", "O" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["P6", "P" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["Q6", "Q" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["R6", "R" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["S6", "S" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["T6", "T" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["U6", "U" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["V6", "V" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["W6", "W" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
            }

            if (cod_estab.ToString().Trim() == "9")
            {


                SqlConnection cn4 = conexion.conectar("BDIntegrador");
                SqlCommand sqlCommand4 = new SqlCommand()
                {
                    Connection = cn4,
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "[MI_Ventas_Articulos_30DPromocion]",
                    CommandTimeout = 0
                };

                sqlCommand4.Parameters.Clear();
                sqlCommand4.Parameters.AddWithValue("@p_dt", dt.ToString("yyyyMMdd"));
                sqlCommand4.Parameters.AddWithValue("@p_estab", "99");

                SqlDataAdapter da4 = new SqlDataAdapter(sqlCommand4);
                System.Data.DataTable dt4 = new System.Data.DataTable();
                da4.Fill(dt4);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt4;
                this.DG1.SelectAll();

                dataObj = DG1.GetClipboardContent();
                if (dataObj != null)
                {
                    Clipboard.SetDataObject(dataObj);

                    hoja = (Worksheet)g_Workbook.Sheets.get_Item(3);
                    hoja.Activate();
                    hoja.Name = "IMP CULIACAN LINEA";
                    /*INSERTAR DATOS DE SUCURSAL*/

                    rango = (Range)hoja.get_Range("K2", "U2");
                    rango.Select();
                    rango.Merge();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    hoja.Cells[2, 11] = "VENTAS DE ARTICULOS A 30 DIAS";
                    hoja.Cells[2, 11].Font.FontStyle = "Bold";

                    rango = (Range)hoja.get_Range("K3", "U3");
                    rango.Select();
                    rango.Merge();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    hoja.Cells[3, 11] = "PRODUCTOS BASICOS Y SUPER BASICOS";
                    hoja.Cells[3, 11].Font.FontStyle = "Bold";

                    //celda ''
                    rango = (Range)hoja.get_Range("B4", "J4");
                    rango.Select();
                    rango.Merge();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                    rango = (Range)hoja.get_Range("K4", "O4");
                    rango.Select();
                    rango.Merge();

                    hoja.Cells[4, 11] = "Semana Anterior  De " + fi_anterior.ToString("yyyy-MM-dd") + " a " + ff_anterior.ToString("yyyy-MM-dd");
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Font.FontStyle = "Bold";

                    rango = (Range)hoja.get_Range("P4", "U4");
                    rango.Select();
                    rango.Merge();
                    hoja.Cells[4, 16] = "Semana Actual De " + fi_actual.ToString("yyyy-MM-dd") + " a " + ff_actual.ToString("yyyy-MM-dd");
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Font.FontStyle = "Bold";


                    rango = (Range)hoja.get_Range("V4", "W4");
                    rango.Select();
                    rango.Merge();

                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Font.FontStyle = "Bold";

                    hoja.Columns["A"].ColumnWidth = 3;

                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    hoja.Cells[4, 3].Font.FontStyle = "Bold";

                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                    for (int j = 0; j < DG1.Columns.Count; j++)
                    {
                        hoja.Cells[5, 2 + j] = DG1.Columns[j].HeaderText;
                        hoja.Cells[5, 2 + j].Font.FontStyle = "Bold";
                        hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        hoja.Cells[5, j + 2].Font.FontStyle = "Bold";
                    }

                    //double dias_venta;
                    ///*Aplicar cambios a la columna piezas a transferir/ pedir */
                    //for (int j = 0; j < DG1.Rows.Count - 1; j++)
                    //{
                    //    hoja.Cells[6 + j, 22] = "=IF(T" + (6 + j) + " > 90, (Q" + (6 + j) + " - ((R" + (6 + j) + " / 30) * 90)) * -1, IF(T" + (6 + j) + " < 60, ((R" + (6 + j) + " / 30) * 60) - Q" + (6 + j) + ",))";
                    //    hoja.Cells[6 + j, 23] = "=ROUND(IF(V" + (6 + j) + "<0,(Q" + (6 + j) + "+V" + (6 + j) + ")/R" + (6 + j) + "*30,(V" + (6 + j) + "+Q" + (6 + j) + ")/R" + (6 + j) + "*30),0)";


                    //    if (hoja.Cells[6 + j, 20].value != null)
                    //    {

                    //        dias_venta = hoja.Cells[6 + j, 20].value;
                    //        if (dias_venta < 60 || dias_venta > 90)
                    //        {
                    //            hoja.Cells[6 + j, 20].Interior.Color = System.Drawing.Color.Red;
                    //        }
                    //    }

                    //}


                    int registros = 0;
                    registros = DG1.Rows.Count + 4;
                    rango = (Range)hoja.get_Range("B6", "W" + registros.ToString());
                    rango.Select();
                    rango.EntireColumn.AutoFit();

                    hoja.Columns["B"].ColumnWidth = 24;
                    hoja.Columns["C"].ColumnWidth = 6.5;
                    hoja.Columns["D"].ColumnWidth = 48;
                    hoja.Columns["E"].ColumnWidth = 15;
                    hoja.Columns["F"].ColumnWidth = 15;
                    hoja.Columns["G"].ColumnWidth = 15;
                    hoja.Columns["H"].ColumnWidth = 21.3;
                    hoja.Columns["I"].ColumnWidth = 7;
                    hoja.Columns["J"].ColumnWidth = 27;

                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                    hoja.Range["K6", "K" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["L6", "L" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["M6", "M" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["N6", "N" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                    hoja.Range["O6", "O" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["P6", "P" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["Q6", "Q" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["R6", "R" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["S6", "S" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                    hoja.Range["T6", "T" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["U6", "U" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["V6", "V" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["W6", "W" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";


                }
            }

            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\ExistenciasDiaVenta30_Suc_" + cod_estab.ToString().Trim() + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\ExistenciasDiaVenta30_Suc_" + cod_estab.ToString().Trim() + ".xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\ExistenciasDiaVenta30_Suc_" + cod_estab.ToString().Trim() + ".xlsx");
            g_Workbook.Close();
            excelApp.Quit();

            return System.Windows.Forms.Application.StartupPath + "\\ExistenciasDiaVenta30_Suc_" + cod_estab.ToString().Trim() + ".xlsx";
        }

        private string Comparativo_semana_semana(DateTime dt, string cod_estab, string nombre)
        {

            DateTime fi_actual;
            DateTime ff_actual;
            DateTime inv_actual;
            DateTime inv_anterior;

            fi_actual = dt.AddDays(-7);
            ff_actual = dt.AddDays(-1);
            inv_actual = dt;
            inv_anterior = fi_actual;

            DateTime fi_anterior;
            DateTime ff_anterior;

            ff_anterior = fi_actual.AddDays(-1);
            fi_anterior = fi_actual.AddDays(-7);


            //MessageBox.Show("fi_actual->" + fi_actual.ToString() + "ff_actual->" + ff_actual.ToString()+"inv_actual->"+inv_actual.ToString());
            //MessageBox.Show("fi_anterior->" + fi_anterior.ToString() + "ff_anterior->" + ff_anterior.ToString()+"inventario_anterior->"+ inv_anterior.ToString());

            Microsoft.Office.Interop.Excel.Range rango;

            SqlConnection cn1 = conexion.conectar("BDIntegrador");
            SqlCommand sqlCommand1 = new SqlCommand()
            {
                Connection = cn1,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Ventas_Articulos_7Dias",
                CommandTimeout = 0
            };

            sqlCommand1.Parameters.Clear();
            sqlCommand1.Parameters.AddWithValue("@p_dt", dt.ToString("yyyyMMdd"));
            sqlCommand1.Parameters.AddWithValue("@p_estab", cod_estab);

            SqlDataAdapter da1 = new SqlDataAdapter(sqlCommand1);
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt1;
            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            DataObject dataObj = DG1.GetClipboardContent();
            excelApp.Visible = false;
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);
            /*tenemos 3 hojas iniciales + la agregada 4*/

            hoja = (Worksheet)g_Workbook.Sheets.get_Item(1);
            hoja.Activate();
            hoja.Name = nombre;
            hoja.Columns["A"].ColumnWidth = 3;

            dataObj = DG1.GetClipboardContent();


            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);

                rango = (Range)hoja.get_Range("F1", "I1");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[1, 6] = "VENTAS DE ARTICULOS A 7 DIAS";
                hoja.Cells[1, 6].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("F2", "I2");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 6] = "PRODUCTOS BASICOS y SUPER BASICOS ";
                hoja.Cells[2, 6].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B4", "H4");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;


                rango = (Range)hoja.get_Range("I4", "K4");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                hoja.Cells[4, 9] = "Semana Anterior  De " + fi_anterior.ToString("yyyy-MM-dd") + " a " + ff_anterior.ToString("yyyy-MM-dd");
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("L4", "N4");
                rango.Select();
                rango.Merge();
                hoja.Cells[4, 12] = "Semana Actual De " + fi_actual.ToString("yyyy-MM-dd") + " a " + ff_actual.ToString("yyyy-MM-dd");
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";


                rango = (Range)hoja.get_Range("O4", "Q4");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";



                rango = (Range)hoja.Cells[6, 1];
                rango.Select();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                hoja.Cells[4, 3].Font.FontStyle = "Bold";

                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[5, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[5, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[5, j + 2].Font.FontStyle = "Bold";
                }

                int registros;
                registros = DG1.Rows.Count + 4;
                rango = (Range)hoja.get_Range("K6", "N" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();

                hoja.Columns["B"].ColumnWidth = 35;
                hoja.Columns["C"].ColumnWidth = 15;

                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                rango = (Range)hoja.get_Range("B6", "Q" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                hoja.Range["I6", "I" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["J6", "J" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["K6", "K" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["L6", "L" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["M6", "M" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["N6", "N" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["O6", "O" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["P6", "P" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["Q6", "Q" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";

            }

            if (cod_estab.ToString().Trim() == "9")
            {

                SqlConnection cn4 = conexion.conectar("BDIntegrador");
                SqlCommand sqlCommand4 = new SqlCommand()
                {
                    Connection = cn4,
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "MI_Ventas_Articulos_7Dias",
                    CommandTimeout = 0
                };

                sqlCommand4.Parameters.Clear();
                sqlCommand4.Parameters.AddWithValue("@p_dt", dt.ToString("yyyyMMdd"));
                sqlCommand4.Parameters.AddWithValue("@p_estab", "99");

                SqlDataAdapter da4 = new SqlDataAdapter(sqlCommand4);
                System.Data.DataTable dt4 = new System.Data.DataTable();
                da4.Fill(dt4);
                this.DG1.DataSource = null;
                this.DG1.Rows.Clear();
                this.DG1.Columns.Clear();
                this.DG1.DataSource = dt4;
                this.DG1.SelectAll();

                dataObj = DG1.GetClipboardContent();
                if (dataObj != null)
                {
                    Clipboard.SetDataObject(dataObj);

                    hoja = (Worksheet)g_Workbook.Sheets.get_Item(2);
                    hoja.Activate();
                    hoja.Name = "IMP CULIACAN LINEA";
                    /*INSERTAR DATOS DE SUCURSAL*/
                    rango = (Range)hoja.get_Range("F1", "I1");
                    rango.Select();
                    rango.Merge();

                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    hoja.Cells[1, 6] = "VENTAS DE ARTICULOS A 7 DIAS";
                    hoja.Cells[1, 6].Font.FontStyle = "Bold";

                    rango = (Range)hoja.get_Range("F2", "I2");
                    rango.Select();
                    rango.Merge();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    hoja.Cells[2, 6] = "PRODUCTOS BASICOS y SUPER BASICOS ";
                    hoja.Cells[2, 6].Font.FontStyle = "Bold";

                    rango = (Range)hoja.get_Range("B4", "H4");
                    rango.Select();
                    rango.Merge();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;


                    rango = (Range)hoja.get_Range("I4", "K4");
                    rango.Select();
                    rango.Merge();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                    hoja.Cells[4, 9] = "Semana Anterior  De " + fi_anterior.ToString("yyyy-MM-dd") + " a " + ff_anterior.ToString("yyyy-MM-dd");
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Font.FontStyle = "Bold";

                    rango = (Range)hoja.get_Range("L4", "N4");
                    rango.Select();
                    rango.Merge();
                    hoja.Cells[4, 12] = "Semana Actual De " + fi_actual.ToString("yyyy-MM-dd") + " a " + ff_actual.ToString("yyyy-MM-dd");
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Font.FontStyle = "Bold";


                    rango = (Range)hoja.get_Range("O4", "Q4");
                    rango.Select();
                    rango.Merge();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    rango.Font.FontStyle = "Bold";



                    rango = (Range)hoja.Cells[6, 1];
                    rango.Select();
                    rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    hoja.Cells[4, 3].Font.FontStyle = "Bold";

                    hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                    for (int j = 0; j < DG1.Columns.Count; j++)
                    {
                        hoja.Cells[5, 2 + j] = DG1.Columns[j].HeaderText;
                        hoja.Cells[5, 2 + j].Font.FontStyle = "Bold";

                        hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        hoja.Cells[5, j + 2].Font.FontStyle = "Bold";
                    }

                    int registros;
                    registros = DG1.Rows.Count + 4;
                    rango = (Range)hoja.get_Range("K6", "N" + registros.ToString());
                    rango.Select();
                    rango.EntireColumn.AutoFit();

                    hoja.Columns["B"].ColumnWidth = 35;
                    hoja.Columns["C"].ColumnWidth = 15;

                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                    rango = (Range)hoja.get_Range("B6", "Q" + registros.ToString());
                    rango.Select();
                    rango.EntireColumn.AutoFit();
                    rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                    hoja.Range["I6", "I" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["J6", "J" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["K6", "K" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                    hoja.Range["L6", "L" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["M6", "M" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["N6", "N" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";
                    hoja.Range["O6", "O" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["P6", "P" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                    hoja.Range["Q6", "Q" + DG1.Rows.Count + 5.ToString()].NumberFormat = "$#,###,##0.00";

                }
            }

            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Comp_Semana_Semana_Suc_" + cod_estab.ToString().Trim() + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Comp_Semana_Semana_Suc_" + cod_estab.ToString().Trim() + ".xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Comp_Semana_Semana_Suc_" + cod_estab.ToString().Trim() + ".xlsx");
            g_Workbook.Close();
            excelApp.Quit();

            return System.Windows.Forms.Application.StartupPath + "\\Comp_Semana_Semana_Suc_" + cod_estab.ToString().Trim() + ".xlsx";
        }

        private string Boletos_Sorteo_Tv(DateTime dt)
        {

            DateTime fi_actual;
            DateTime ff_actual;

            /*Rango de fecha */
            fi_actual = dt.AddDays(-7);
            ff_actual = dt.AddDays(-1);

            Microsoft.Office.Interop.Excel.Range rango;
            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand sqlCommand1 = new SqlCommand()
            {
                Connection = cn1,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Sorteo_Tv",
                CommandTimeout = 0
            };

            sqlCommand1.Parameters.Clear();

            sqlCommand1.Parameters.AddWithValue("@fecha_inicio", fi_actual);
            sqlCommand1.Parameters.AddWithValue("@fecha_fin", ff_actual);

            SqlDataAdapter da1 = new SqlDataAdapter(sqlCommand1);
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt1;
            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            DataObject dataObj = DG1.GetClipboardContent();
            excelApp.Visible = false;
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);
            /*tenemos 3 hojas iniciales + la agregada 4*/

            hoja = (Worksheet)g_Workbook.Sheets.get_Item(1);
            hoja.Activate();
            hoja.Name = "BOLETAJE SORTEO";
            hoja.Columns["A"].ColumnWidth = 3;

            dataObj = DG1.GetClipboardContent();


            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);

                rango = (Range)hoja.get_Range("B1", "J1");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[1, 2] = "BOLETOS DINAMICA SORTEO";
                hoja.Cells[1, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B2", "J2");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 2] = "(DINAMICAS POR TEMPORADA)";
                hoja.Cells[2, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B4", "B4");
                rango.Select();
                rango.Merge();

                hoja.Cells[4, 2] = "Semana " + fi_actual.ToString("yyyy-MM-dd") + " a " + ff_actual.ToString("yyyy-MM-dd");
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                rango.Font.FontStyle = "Bold";

                rango = (Range)hoja.Cells[6, 1];
                rango.Select();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                //hoja.Cells[4, 3].Font.FontStyle = "Bold";

                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[5, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[5, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[5, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[5, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    //hoja.Cells[5, j + 2].Font.FontStyle = "Bold";
                }

                int registros;
                registros = DG1.Rows.Count + 6;

                rango = (Range)hoja.get_Range("B4", "J" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;


                /*suma de valores en grid*/
                double dia_1 = 0; double dia_2 = 0; double dia_3 = 0;
                double dia_4 = 0; double dia_5 = 0; double dia_6 = 0;
                double dia_7 = 0; double dia_8 = 0;
                foreach (DataGridViewRow row in DG1.Rows)
                {
                    dia_1 += Convert.ToDouble(row.Cells[1].Value);
                    dia_2 += Convert.ToDouble(row.Cells[2].Value);
                    dia_3 += Convert.ToDouble(row.Cells[3].Value);
                    dia_4 += Convert.ToDouble(row.Cells[4].Value);
                    dia_5 += Convert.ToDouble(row.Cells[5].Value);
                    dia_6 += Convert.ToDouble(row.Cells[6].Value);
                    dia_7 += Convert.ToDouble(row.Cells[7].Value);
                    dia_8 += Convert.ToDouble(row.Cells[8].Value);
                }

                hoja.Cells[DG1.Rows.Count + 6, 2] = "TOTAL"; hoja.Cells[DG1.Rows.Count + 6, 2].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 3] = Convert.ToString(dia_1); hoja.Cells[DG1.Rows.Count + 6, 3].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 4] = Convert.ToString(dia_2); hoja.Cells[DG1.Rows.Count + 6, 4].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 5] = Convert.ToString(dia_3); hoja.Cells[DG1.Rows.Count + 6, 5].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 6] = Convert.ToString(dia_4); hoja.Cells[DG1.Rows.Count + 6, 6].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 7] = Convert.ToString(dia_5); hoja.Cells[DG1.Rows.Count + 6, 7].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 8] = Convert.ToString(dia_6); hoja.Cells[DG1.Rows.Count + 6, 8].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 9] = Convert.ToString(dia_7); hoja.Cells[DG1.Rows.Count + 6, 9].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 6, 10] = Convert.ToString(dia_8); hoja.Cells[DG1.Rows.Count + 6, 10].Font.FontStyle = "Bold";

                /*Asignar formatos a celdas*/
                hoja.Range["C6", "C" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["D6", "D" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["E6", "E" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["F6", "F" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["G6", "G" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["H6", "H" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["I6", "I" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["J6", "J" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";


            }



            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Dinamica_Sorteo" + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Dinamica_Sorteo" + ".xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Dinamica_Sorteo" + ".xlsx");
            g_Workbook.Close();
            excelApp.Quit();

            return System.Windows.Forms.Application.StartupPath + "\\Dinamica_Sorteo" + ".xlsx";
        }

        private string Acumulado_Ventas_Mensual(DateTime dt, string cod_estab, string nombre)
        {

            /*Agregar validaciones de fecha*/
            DateTime fecha_Inicio;
            DateTime fecha_Final;

            int num_dia = Int32.Parse(dt.ToString("dd"));

            //Falta Agregar Validacion en caso de que  dia lunes sea primero agarrar mes anterior 
            if (num_dia == 1)
            {
                fecha_Inicio = dt.AddMonths(-1);
                fecha_Final = dt.AddDays(-1);
            }
            else
            {
                fecha_Inicio = new DateTime(dt.Year, dt.Month, 1);
                fecha_Final = dt.AddDays(-1);
            }

            //Validacion consultar  de acuerdo al mes actual
            //En caso de enero- febreo  se compara con año anterior, En caso de Marzo se compara con 2 años hacia atras ,En caso de Abril a julio con 2 años hacia atras 
            switch (Int32.Parse(fecha_Inicio.ToString("MM")))
            {
                case 1:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 2:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 3:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 4:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 5:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 6:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 7:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 8:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 9:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 10:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 11:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 12:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                default:
                    break;
            }

            //Imprimir valores de fechas
            //MessageBox.Show("Fecha Inicio A->" + fecha_Inicio + "Fecha Final actual-->" + fecha_Final);
            //MessageBox.Show("Fecha Inicio Anterior->" + fi_anterior + "Fecha Final Anterior -->" + ff_anterior);

            Microsoft.Office.Interop.Excel.Range rango;
            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand sqlCommand1 = new SqlCommand()
            {
                Connection = cn1,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Acumulado_Venta_Mensual",
                CommandTimeout = 0
            };

            sqlCommand1.Parameters.Clear();
            sqlCommand1.Parameters.AddWithValue("@dia_actual", dt.ToString("yyyyMMdd"));
            sqlCommand1.Parameters.AddWithValue("@cod_estab", cod_estab);

            SqlDataAdapter da1 = new SqlDataAdapter(sqlCommand1);
            System.Data.DataTable dt1 = new System.Data.DataTable();

            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            Clipboard.Clear();
            da1.Fill(dt1);
            this.DG1.DataSource = dt1;
            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            DataObject objeto = DG1.GetClipboardContent();
            excelApp.Visible = false;
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);
            //tenemos 3 hojas iniciales + la agregada 4

            hoja = (Worksheet)g_Workbook.Sheets.get_Item(1);
            hoja.Activate();
            hoja.Name = nombre.ToString();
            hoja.Columns["A"].ColumnWidth = 3;

            //dataObj = DG1.GetClipboardContent();
            if (objeto != null)
            {
                Clipboard.SetDataObject(objeto);

                rango = (Range)hoja.get_Range("B1", "J1");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[1, 2] = "ACUMULADO DE VENTAS MENSUAL";
                hoja.Cells[1, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B2", "J2");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 2] = "ACUMULADO " + fecha_Inicio.ToString("dd") + "-" + fecha_Final.ToString("dd") + " " + obtenerNombreMesAbreviadoNumero(Int32.Parse(fecha_Inicio.ToString("MM"))) + " " + fecha_Inicio.ToString("yyyy") + " - " + fi_anterior.ToString("dd") + "-" + ff_anterior.ToString("dd") + " " + obtenerNombreMesAbreviadoNumero(Int32.Parse(fi_anterior.ToString("MM"))) + " " + fi_anterior.ToString("yyyy");
                hoja.Cells[2, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.Cells[4, 1];
                rango.Select();


                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[3, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[3, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[3, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[3, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[3, j + 2].Font.FontStyle = "Bold";
                }



                int registros = 0;
                registros = DG1.Rows.Count + 3;

                rango = (Range)hoja.get_Range("B3", "Q" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                //Asignar formatos a celdas

                int final_registros = 0;
                final_registros = DG1.Rows.Count + 6;

                hoja.Range["B4", "B" + final_registros.ToString()].EntireColumn.NumberFormat = "DD/MM/YYYY";
                hoja.Range["C4", "C" + final_registros.ToString()].NumberFormat = "@";
                hoja.Range["D4", "D" + final_registros.ToString()].NumberFormat = "@";
                hoja.Range["E4", "E" + final_registros.ToString()].NumberFormat = "@";
                hoja.Range["F4", "F" + final_registros.ToString()].NumberFormat = "@";
                hoja.Range["G4", "G" + final_registros.ToString()].NumberFormat = "@";
                hoja.Range["H4", "H" + final_registros.ToString()].NumberFormat = "@";
                hoja.Range["I4", "I" + final_registros.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["J4", "J" + final_registros.ToString()].NumberFormat = "#,###,##0";

                hoja.Range["K4", "K" + final_registros.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["L4", "L" + final_registros.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["M4", "M" + final_registros.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["N4", "N" + final_registros.ToString()].NumberFormat = "$#,###,##0.00";

                hoja.Range["O4", "O" + final_registros.ToString()].NumberFormat = "#,###,##0";

                hoja.Range["P4", "P" + final_registros.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["Q4", "Q" + final_registros.ToString()].NumberFormat = "$#,###,##0.00";

            }

            objeto = null;

            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Acumulado_ventas_suc_" + cod_estab.ToString().Trim() + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Acumulado_ventas_suc_" + cod_estab.ToString().Trim() + ".xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Acumulado_ventas_suc_" + cod_estab.ToString().Trim() + ".xlsx");
            g_Workbook.Close();
            excelApp.Quit();

            return System.Windows.Forms.Application.StartupPath + "\\Acumulado_ventas_suc_" + cod_estab.ToString().Trim() + ".xlsx";
        }

        private string Señalizacion_promociones(DateTime dt)
        {

            DateTime fi_actual;
            DateTime ff_actual;

            /*Rango de fecha */
            fi_actual = dt.AddDays(-7);
            ff_actual = dt.AddDays(-1);

            string mes_actual = dt.AddDays(-1).ToString("MM");


            string dia_fi = dt.AddDays(-7).ToString("dd");
            string dia_ff = dt.AddDays(-1).ToString("dd");

            int num_mes = Int32.Parse(mes_actual);
            string nom_mes = "";

            nom_mes = obtenerNombreMesAbreviadoNumero(num_mes);

            Microsoft.Office.Interop.Excel.Range rango;
            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand sqlCommand1 = new SqlCommand()
            {
                Connection = cn1,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Revision_Promociones",
                CommandTimeout = 0
            };

            sqlCommand1.Parameters.Clear();

            sqlCommand1.Parameters.AddWithValue("@1FI", fi_actual.ToString("yyyyMMdd") + " 00:00:00");
            sqlCommand1.Parameters.AddWithValue("@2FF", ff_actual.ToString("yyyyMMdd") + " 23:59:59");

            SqlDataAdapter da1 = new SqlDataAdapter(sqlCommand1);
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt1;
            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            DataObject dataObj = DG1.GetClipboardContent();
            excelApp.Visible = false;
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);


            hoja = (Worksheet)g_Workbook.Sheets.get_Item(1);
            hoja.Activate();
            hoja.Name = "LISTADO PROMOCIONES";
            hoja.Columns["A"].ColumnWidth = 3;

            dataObj = DG1.GetClipboardContent();


            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);

                rango = (Range)hoja.get_Range("B1", "M1");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                hoja.Cells[1, 2] = "ACUMULADO VENTA DE ARTICULOS EN PROMOCION";
                hoja.Cells[1, 2].Font.FontStyle = "Bold";


                rango = (Range)hoja.get_Range("B2", "M2");
                rango.Select();
                rango.Merge();
                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                hoja.Cells[2, 2] = "Semana ".ToUpper() + dia_fi.ToString() + " - " + dia_ff.ToString() + " " + nom_mes.ToString().ToUpper();
                hoja.Cells[2, 2].Font.FontStyle = "Bold";


                rango = (Range)hoja.Cells[5, 1];
                rango.Select();


                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[4, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[4, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[4, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[4, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;
                    hoja.Cells[4, j + 2].Font.FontStyle = "Bold";
                }

                int registros;
                registros = DG1.Rows.Count + 4;

                rango = (Range)hoja.get_Range("B4", "M" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;


                double p_1 = 0; double p_2 = 0; double p_3 = 0;
                double p_4 = 0; double p_5 = 0; double p_6 = 0;
                double p_7 = 0; double p_8 = 0; double p_9 = 0;
                double p_10 = 0; double p_11 = 0;
                foreach (DataGridViewRow row in DG1.Rows)
                {
                    p_1 += Convert.ToDouble(row.Cells[1].Value);
                    p_2 += Convert.ToDouble(row.Cells[2].Value);
                    p_3 += Convert.ToDouble(row.Cells[3].Value);
                    p_4 += Convert.ToDouble(row.Cells[4].Value);
                    p_5 += Convert.ToDouble(row.Cells[5].Value);
                    p_6 += Convert.ToDouble(row.Cells[6].Value);
                    p_7 += Convert.ToDouble(row.Cells[7].Value);
                    p_8 += Convert.ToDouble(row.Cells[8].Value);
                    p_9 += Convert.ToDouble(row.Cells[9].Value);
                    p_10 += Convert.ToDouble(row.Cells[10].Value);
                    p_11 += Convert.ToDouble(row.Cells[11].Value);
                }

                hoja.Cells[DG1.Rows.Count + 4, 2] = "TOTAL"; hoja.Cells[DG1.Rows.Count + 4, 2].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 3] = Convert.ToString(p_1); hoja.Cells[DG1.Rows.Count + 4, 3].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 4] = Convert.ToString(p_2); hoja.Cells[DG1.Rows.Count + 4, 4].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 5] = Convert.ToString(p_3); hoja.Cells[DG1.Rows.Count + 4, 5].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 6] = Convert.ToString(p_4); hoja.Cells[DG1.Rows.Count + 4, 6].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 7] = Convert.ToString(p_5); hoja.Cells[DG1.Rows.Count + 4, 7].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 8] = Convert.ToString(p_6); hoja.Cells[DG1.Rows.Count + 4, 8].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 9] = Convert.ToString(p_7); hoja.Cells[DG1.Rows.Count + 4, 9].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 10] = Convert.ToString(p_8); hoja.Cells[DG1.Rows.Count + 4, 10].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 11] = Convert.ToString(p_9); hoja.Cells[DG1.Rows.Count + 4, 11].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 12] = Convert.ToString(p_10); hoja.Cells[DG1.Rows.Count + 4, 12].Font.FontStyle = "Bold";
                hoja.Cells[DG1.Rows.Count + 4, 13] = Convert.ToString(p_11); hoja.Cells[DG1.Rows.Count + 4, 13].Font.FontStyle = "Bold";

                hoja.Range["C6", "C" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["D6", "D" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["E6", "E" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["F6", "F" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["G6", "G" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["H6", "H" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["I6", "I" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["J6", "J" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["K6", "K" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["L6", "L" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["M6", "M" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["N6", "N" + DG1.Rows.Count + 5.ToString()].NumberFormat = "#,###,##0";

            }



            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\señalizacion_promociones" + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\señalizacion_promociones" + ".xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\señalizacion_promociones" + ".xlsx");
            g_Workbook.Close();

            return System.Windows.Forms.Application.StartupPath + "\\señalizacion_promociones" + ".xlsx";

        }

        private string Desplazamiento_Temporada(DateTime dt)
        {
            DateTime fecha_Inicio;
            DateTime fecha_Final;

            int num_dia = Int32.Parse(dt.ToString("dd"));

            //Falta Agregar Validacion en caso de que  dia lunes sea primero agarrar mes anterior 
            if (num_dia == 1)
            {
                fecha_Inicio = dt.AddMonths(-1);
                fecha_Final = dt.AddDays(-1);
            }
            else
            {
                fecha_Inicio = new DateTime(dt.Year, dt.Month, 1);
                fecha_Final = dt;
            }

            string nom_mes_actual = "";
            int num_mes_actual = Int32.Parse(fecha_Inicio.ToString("MM"));
            nom_mes_actual = obtenerNombreMesAbreviadoNumero(num_mes_actual);

            //Validacion consultar  de acuerdo al mes actual
            //En caso de enero- febreo  se compara con año anterior, En caso de Marzo se compara con 2 años hacia atras ,En caso de Abril a julio con 2 años hacia atras 
            switch (Int32.Parse(fecha_Inicio.ToString("MM")))
            {
                case 1:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 2:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 3:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 4:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 5:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 6:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 7:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 2);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 8:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 9:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 10:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 11:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                case 12:
                    diferencia = Int32.Parse(fecha_Inicio.ToString("yyyy")) - (Int32.Parse(fecha_Inicio.ToString("yyyy")) - 1);
                    fi_anterior = fecha_Inicio.AddYears(-diferencia);
                    ff_anterior = fecha_Final.AddYears(-diferencia);
                    break;
                default:
                    break;
            }


            Microsoft.Office.Interop.Excel.Range rango;
            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand sqlCommand1 = new SqlCommand()
            {
                Connection = cn1,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Desplazamiento_Temporada",
                CommandTimeout = 0
            };

            sqlCommand1.Parameters.Clear();
            sqlCommand1.Parameters.AddWithValue("@dia_actual", fecha_Final);
            sqlCommand1.Parameters.AddWithValue("@ano", diferencia);
            sqlCommand1.Parameters.AddWithValue("@agrupado", "L");

            SqlDataAdapter da1 = new SqlDataAdapter(sqlCommand1);
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt1;
            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            DataObject dataObj = DG1.GetClipboardContent();
            excelApp.Visible = false;
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);
            //tenemos 3 hojas iniciales + la agregada 4

            hoja = (Worksheet)g_Workbook.Sheets.get_Item(1);
            hoja.Activate();
            hoja.Name = "Despl Por Linea";
            hoja.Columns["A"].ColumnWidth = 3;

            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);

                rango = (Range)hoja.get_Range("B1", "N1");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[1, 2] = "DESPLAZAMIENTO DE TEMPORADA POR LINEA";
                hoja.Cells[1, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B2", "N2");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 2] = "ACUMULADO  [M1] " + fecha_Inicio.ToString("dd") + " - " + fecha_Final.ToString("dd") + " " + nom_mes_actual + " " + fecha_Inicio.ToString("yyyy") + "    [M2] " + fi_anterior.ToString("dd") + " - " + ff_anterior.ToString("dd") + " " + nom_mes_actual + " " + fi_anterior.ToString("yyyy");
                hoja.Cells[2, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.Cells[4, 1];
                rango.Select();

                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[3, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[3, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[3, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[3, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[3, j + 2].Font.FontStyle = "Bold";
                }

                int registros = 0;
                registros = DG1.Rows.Count + 2;

                rango = (Range)hoja.get_Range("B3", "N" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;
                //Asignar formatos a celdas
                hoja.Range["B4", "B" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["C4", "C" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["D4", "D" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["E4", "E" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["F4", "F" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["G4", "G" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["H4", "H" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["I4", "I" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["J4", "J" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["K4", "K" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["L4", "L" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["M4", "M" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0.00";
                hoja.Range["N4", "N" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0.00";
            }

            SqlConnection cn2 = conexion.conectar("BMSNayar");
            SqlCommand sqlCommand2 = new SqlCommand()
            {
                Connection = cn2,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Desplazamiento_Temporada",
                CommandTimeout = 0
            };

            sqlCommand2.Parameters.Clear();
            sqlCommand2.Parameters.AddWithValue("@dia_actual", fecha_Final);
            sqlCommand2.Parameters.AddWithValue("@ano", diferencia);
            sqlCommand2.Parameters.AddWithValue("@agrupado", "C");

            SqlDataAdapter da2 = new SqlDataAdapter(sqlCommand2);
            System.Data.DataTable dt2 = new System.Data.DataTable();
            da2.Fill(dt2);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt2;
            this.DG1.SelectAll();

            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);
                hoja = (Worksheet)g_Workbook.Sheets.get_Item(2);
                hoja.Activate();
                hoja.Name = "Despl Por Clasificacion";
                hoja.Columns["A"].ColumnWidth = 3;
                //INSERTAR DATOS DE SUCURSAL

                rango = (Range)hoja.get_Range("B1", "P1");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[1, 2] = "DESPLAZAMIENTO DE TEMPORADA POR CLASIFICACION";
                hoja.Cells[1, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B2", "P2");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 2] = "ACUMULADO  [M1] " + fecha_Inicio.ToString("dd") + " - " + fecha_Final.ToString("dd") + " " + nom_mes_actual + " " + fecha_Inicio.ToString("yyyy") + "   [M2] " + fi_anterior.ToString("dd") + " - " + ff_anterior.ToString("dd") + " " + nom_mes_actual + " " + fi_anterior.ToString("yyyy");
                hoja.Cells[2, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.Cells[4, 1];
                rango.Select();

                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[3, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[3, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[3, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[3, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[3, j + 2].Font.FontStyle = "Bold";
                }

                int registros = 0;
                registros = DG1.Rows.Count + 2;

                rango = (Range)hoja.get_Range("B3", "P" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                //Asignar formatos a celdas
                hoja.Range["B4", "B" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["C4", "C" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["D4", "D" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["E4", "E" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";

                hoja.Range["F4", "F" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["G4", "G" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["H4", "H" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["I4", "I" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["H4", "H" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["I4", "I" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["J4", "J" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["K4", "K" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";

                hoja.Range["L4", "L" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["M4", "M" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["N4", "N" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["O4", "O" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0.00";
                hoja.Range["P4", "P" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0.00";
            }

            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Desplazamiento_Temporada_Suc.xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Desplazamiento_Temporada_Suc.xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Desplazamiento_Temporada_Suc.xlsx");
            g_Workbook.Close();
            excelApp.Quit();


            return System.Windows.Forms.Application.StartupPath + "\\Desplazamiento_Temporada_Suc.xlsx";
        }

        private string Menos_vendidos(DateTime dt, string cod_estab, string nombre)
        {
            DateTime fi_actual;
            DateTime ff_actual;

            /*Rango de fecha */
            fi_actual = dt.AddDays(-30);
            ff_actual = dt.AddDays(-1);

            Microsoft.Office.Interop.Excel.Range rango;
            SqlConnection cn1 = conexion.conectar("BDIntegrador");
            SqlCommand sqlCommand1 = new SqlCommand()
            {
                Connection = cn1,
                CommandType = CommandType.StoredProcedure,
                CommandText = "MI_Articulos_Menos_Venta",
                CommandTimeout = 0
            };

            sqlCommand1.Parameters.Clear();
            sqlCommand1.Parameters.AddWithValue("@dia_actual", dt);
            sqlCommand1.Parameters.AddWithValue("@cod_estab", cod_estab);

            SqlDataAdapter da1 = new SqlDataAdapter(sqlCommand1);
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            this.DG1.DataSource = null;
            this.DG1.Rows.Clear();
            this.DG1.Columns.Clear();
            this.DG1.DataSource = dt1;
            this.DG1.SelectAll();

            Excel.Application excelApp = new Excel.Application();
            DataObject dataObj = DG1.GetClipboardContent();
            excelApp.Visible = false;
            Excel.Workbook g_Workbook = excelApp.Application.Workbooks.Add();
            Excel.Worksheet hoja = g_Workbook.Sheets.Add(After: g_Workbook.Sheets[g_Workbook.Sheets.Count]);
            //tenemos 3 hojas iniciales + la agregada 4

            hoja = (Worksheet)g_Workbook.Sheets.get_Item(1);
            hoja.Activate();
            hoja.Name = nombre.ToString();
            hoja.Columns["A"].ColumnWidth = 3;

            dataObj = DG1.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);

                rango = (Range)hoja.get_Range("B1", "J1");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[1, 2] = "ARTICULOS MENOS VENDIDOS";
                hoja.Cells[1, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.get_Range("B2", "J2");
                rango.Select();
                rango.Merge();

                rango.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                hoja.Cells[2, 2] = "PERIODO " + fi_actual.ToString("dd-MM-yyyy") + " AL " + ff_actual.ToString("dd-MM-yyyy");
                hoja.Cells[2, 2].Font.FontStyle = "Bold";

                rango = (Range)hoja.Cells[4, 1];
                rango.Select();


                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                for (int j = 0; j < DG1.Columns.Count; j++)
                {
                    hoja.Cells[3, 2 + j] = DG1.Columns[j].HeaderText;
                    hoja.Cells[3, 2 + j].Font.FontStyle = "Bold";

                    hoja.Cells[3, j + 2].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    hoja.Cells[3, j + 2].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    hoja.Cells[3, j + 2].Font.FontStyle = "Bold";
                }

                int registros = 0;
                registros = DG1.Rows.Count + 3;

                rango = (Range)hoja.get_Range("B3", "L" + registros.ToString());
                rango.Select();
                rango.EntireColumn.AutoFit();
                rango.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rango.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                //Asignar formatos a celdas

                hoja.Range["B4", "B" + DG1.Rows.Count + 6.ToString()].EntireColumn.NumberFormat = "@";
                hoja.Range["C4", "C" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["D4", "D" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["E4", "E" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["F4", "F" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";
                hoja.Range["G4", "G" + DG1.Rows.Count + 6.ToString()].NumberFormat = "@";

                hoja.Range["H4", "H" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0"; ;
                hoja.Range["I4", "I" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0"; ;
                hoja.Range["J4", "J" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0";
                hoja.Range["K4", "K" + DG1.Rows.Count + 6.ToString()].NumberFormat = "$#,###,##0.00";
                hoja.Range["L4", "L" + DG1.Rows.Count + 6.ToString()].NumberFormat = "#,###,##0.00";
            }

            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Articulos_Menos_Vendidos_suc" + cod_estab.ToString() + ".xlsx"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Articulos_Menos_Vendidos_suc" + cod_estab.ToString() + ".xlsx");
            }

            g_Workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Articulos_Menos_Vendidos_suc" + cod_estab.ToString() + ".xlsx");
            g_Workbook.Close();
            excelApp.Quit();

            return System.Windows.Forms.Application.StartupPath + "\\Articulos_Menos_Vendidos_suc" + cod_estab.ToString() + ".xlsx";
        }

        private string sCol(int nCol)
        {
            switch (nCol)
            {
                case 1:
                    return "A";
                case 2:
                    return "B";
                case 3:
                    return "C";
                case 4:
                    return "D";
                case 5:
                    return "E";
                case 6:
                    return "F";
                case 7:
                    return "G";
                case 8:
                    return "H";
                case 9:
                    return "I";
                case 10:
                    return "J";
                case 11:
                    return "K";
                case 12:
                    return "L";
                case 13:
                    return "M";
                case 14:
                    return "N";
                case 15:
                    return "O";
                case 16:
                    return "P";
                case 17:
                    return "Q";
                case 18:
                    return "R";
                case 19:
                    return "S";
                case 20:
                    return "T";
                case 21:
                    return "U";
                case 22:
                    return "V";
                case 23:
                    return "W";
                case 24:
                    return "X";
                case 25:
                    return "Y";
                case 26:
                    return "Z";
                case 27:
                    return "AA";
                case 28:
                    return "AB";
                case 29:
                    return "AC";
                case 30:
                    return "AD";
                case 31:
                    return "AE";
                case 32:
                    return "AF";
                case 33:
                    return "AG";
                case 34:
                    return "AH";
                case 35:
                    return "AI";
                case 36:
                    return "AJ";
                case 37:
                    return "AK";
                case 38:
                    return "AL";
                case 39:
                    return "AM";
                case 40:
                    return "AN";
                case 41:
                    return "AO";
                case 42:
                    return "AP";
                case 43:
                    return "AQ";
                case 44:
                    return "AR";
                case 45:
                    return "AS";
                case 46:
                    return "AT";
                case 47:
                    return "AU";
                case 48:
                    return "AV";
                case 49:
                    return "AW";
                case 50:
                    return "AX";
                case 51:
                    return "AY";
                case 52:
                    return "AZ";
                case 53:
                    return "BA";
                case 54:
                    return "BB";
                case 55:
                    return "BC";
                case 56:
                    return "BD";
                case 57:
                    return "BE";
                case 58:
                    return "BF";
                case 59:
                    return "BG";
                case 60:
                    return "BH";
                case 61:
                    return "BI";
                case 62:
                    return "BJ";
                case 63:
                    return "BK";
                case 64:
                    return "BL";
                case 65:
                    return "BM";
                case 66:
                    return "BN";
                case 67:
                    return "BO";
                case 68:
                    return "BP";
                case 69:
                    return "BQ";
                case 70:
                    return "BR";
                case 71:
                    return "BS";
                case 72:
                    return "BT";
                case 73:
                    return "BU";
                case 74:
                    return "BV";
                case 75:
                    return "BW";
                case 76:
                    return "BX";
                case 77:
                    return "BY";
                case 78:
                    return "BZ";
                case 79:
                    return "CA";
                case 80:
                    return "CB";
                case 81:
                    return "CC";
                case 82:
                    return "CD";
                case 83:
                    return "CE";
                case 84:
                    return "CF";
                case 85:
                    return "CG";
                case 86:
                    return "CH";
                case 87:
                    return "CI";
                case 88:
                    return "CJ";
                case 89:
                    return "CK";
                case 90:
                    return "CL";
                case 91:
                    return "CM";
                case 92:
                    return "CN";
                case 93:
                    return "CO";
                case 94:
                    return "CP";
                case 95:
                    return "CQ";
                case 96:
                    return "CR";
                case 97:
                    return "CS";
                case 98:
                    return "CT";
                case 99:
                    return "CU";
                case 100:
                    return "CV";
                case 101:
                    return "CW";
                case 102:
                    return "CX";
                case 103:
                    return "CY";
                case 104:
                    return "CZ";
                default:
                    return "";

            }
        }

        private string DevolucionesCancelacionesSucursal(string coordinador, string nombre, DateTime FechaHora)
        {
            DateTime fi;
            DateTime ff;

            fi = FechaHora.AddDays(-7);  // Lunes de la semana pasada
            ff = FechaHora.AddDays(-1);  // Ayer domingo

            Microsoft.Office.Interop.Excel.Application excel;
            excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Application.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook libro;
            libro = excel.Workbooks.Add();

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("select mail.cod_estab, substring(replace(upper(establecimientos.nombre), '.', ''), 1, 30) as nombre, mail.email_cordinador from BMSNayar.dbo.establecimientos inner join BMSNayar.dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') and mail.email_cordinador = '" + coordinador + "' order by CAST(mail.cod_estab as int) DESC", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    SqlConnection cn = conexion.conectar("BMSNayar");
                    SqlCommand comando = new SqlCommand();
                    comando.Connection = cn;
                    comando.CommandType = CommandType.StoredProcedure;
                    comando.CommandText = "MI_CancelYDevol";
                    comando.CommandTimeout = 240;
                    Worksheet hoja = new Worksheet();
                    Microsoft.Office.Interop.Excel.Range rango;
                    int nreportes = 2;
                    for (int i = 1; i <= nreportes; i++)
                    {
                        comando.Parameters.Clear();
                        libro.Worksheets.Add();
                        hoja = (Worksheet)libro.Worksheets.get_Item(1);
                        //libro.Worksheets.Add(Type.Missing,libro.Worksheets[libro.Worksheets.Count], 1, -4167);
                        //hoja = (Worksheet)libro.Worksheets.get_Item(libro.Worksheets.Count);
                        switch (i)
                        {
                            case 1:
                                comando.Parameters.AddWithValue("@3TipoMov", 1);
                                comando.Parameters.AddWithValue("@1FechaIni", fi);
                                comando.Parameters.AddWithValue("@2FechaFin", ff);
                                comando.Parameters.AddWithValue("@4cod_estab", dr1["cod_estab"].ToString());
                                hoja.Name = "DEVOLUCIONES ESTAB " + dr1["cod_estab"].ToString();
                                break;
                            case 2:
                                comando.Parameters.AddWithValue("@3TipoMov", 2);
                                comando.Parameters.AddWithValue("@1FechaIni", fi);
                                comando.Parameters.AddWithValue("@2FechaFin", ff);
                                comando.Parameters.AddWithValue("@4cod_estab", dr1["cod_estab"].ToString());
                                hoja.Name = "CANCELACIONES ESTAB " + dr1["cod_estab"].ToString();
                                break;
                        }
                        SqlDataAdapter da = new SqlDataAdapter(comando);
                        System.Data.DataTable dt = new System.Data.DataTable();
                        da.Fill(dt);
                        DG1.DataSource = null;
                        DG1.Rows.Clear();
                        DG1.Columns.Clear();
                        DG1.DataSource = dt;
                        DG1.SelectAll();
                        object objeto = DG1.GetClipboardContent();

                        try
                        {

                            if (objeto != null)
                            {
                                Clipboard.SetDataObject(objeto);
                                foreach (DataGridViewColumn columna in DG1.Columns)
                                {
                                    hoja.Cells[1, columna.Index + 2] = columna.Name.ToString().ToUpper();
                                }
                                rango = (Range)hoja.Cells[2, 1];
                                rango.Select();
                                hoja.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                            }
                            Clipboard.Clear();
                            objeto = null;
                        }
                        catch (Exception e)
                        {
                            // Si no se hace nada cuando pasa un error sino que solo se muestra en la etiqueta informativa, el proceso de envío del reporte continúa...
                            //MessageBox.Show(e.Message);
                            lblEstado.Text = e.Message;
                        }

                    }

                    if (cn.State == ConnectionState.Open) { cn.Close(); }
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\Cancelaciones y Devoluciones del " + fi.ToString("dd-MMM-yyyy") + " al " + ff.ToString("dd-MMM-yyyy") + " para " + nombre.Trim() + ".xlsb"))
            {
                File.Delete(System.Windows.Forms.Application.StartupPath + "\\Cancelaciones y Devoluciones del " + fi.ToString("dd-MMM-yyyy") + " al " + ff.ToString("dd-MMM-yyyy") + " para " + nombre.Trim() + ".xlsb");
            }

            libro.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Cancelaciones y Devoluciones del " + fi.ToString("dd-MMM-yyyy") + " al " + ff.ToString("dd-MMM-yyyy") + " para " + nombre.Trim() + ".xlsb", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12);
            libro.Close();
            excel.Quit();
            return System.Windows.Forms.Application.StartupPath + "\\Cancelaciones y Devoluciones del " + fi.ToString("dd-MMM-yyyy") + " al " + ff.ToString("dd-MMM-yyyy") + " para " + nombre.Trim() + ".xlsb";
        }

        #region Botones de Reportes

        private void btn_ReporteApertura_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte de Apertura";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("luis.guerrero@mercadodeimportaciones.com");
            emailList.Add("alberto.martinez@mercadodeimportaciones.com");
            emailList.Add("prevencion@mercadodeimportaciones.com");
            emailList.Add("francisco.ontiveros@mercadodeimportaciones.com");
            emailList.Add("Gerencia.Sistemas@mercadodeimportaciones.com");
            emailList.Add("maferperezle01@gmail.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_ReporteApertura.Enabled = false;
            EnviaMailGmail(ReporteApertura(fechaReporte), destinatarios);
            this.lblEstado.Text = "Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_ReporteApertura.Enabled = true;
        }

        private void btn_ReporteVentasDescuento_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte de Ventas con Descuento";

            // destinatarios
            emailList.Add("monica.perez @mercadodeimportaciones.com");
            emailList.Add("jassiel.perez @mercadodeimportaciones.com");
            emailList.Add("alberto.martinez@mercadodeimportaciones.com");
            emailList.Add("francisco.ontiveros@mercadodeimportaciones.com");
            emailList.Add("luis.guerrero@mercadodeimportaciones.com");
            emailList.Add("mercadotecniaauxiliar@mercadodeimportaciones.com");
            emailList.Add("luis.cota@mercadodeimportaciones.com");
            emailList.Add("mario.serrano@mercadodeimportaciones.com");
            emailList.Add("annacelia.soto@mercadodeimportaciones.com");
            emailList.Add("asofia.mercadodeimportaciones@gmail.com");
            emailList.Add("Gerencia.Sistemas@mercadodeimportaciones.com");
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            emailList.Add("hibrajid.lara@mercadodeimportaciones.com");

            string destinatarios = string.Join(",", emailList);
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_ReporteVentasDescuento.Enabled = false;
            EnviaMailGmail(ReporteDescuento(fechaReporte), destinatarios);
            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_ReporteVentasDescuento.Enabled = true;
        }
        
        private void btn_ReporteIncidencias_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte de Incidencias";

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_ReporteIncidencias.Enabled = false;
            this.Incidencias();
            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_ReporteIncidencias.Enabled = true;
        }
        
        private void btn_Top80_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Top 80";

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_Top80.Enabled = false;
            this.Top80();
            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_Top80.Enabled = true;
        }

        private void btn_ComparativoPresupuesto_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Comparativo vs Presupuesto";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("director@mercadodeimportaciones.com");
            emailList.Add("guillermina.cervantes@mercadodeimportaciones.com");
            emailList.Add("culiacan@mercadodeimportaciones.com");
            emailList.Add("guamuchilcentro@mercadodeimportaciones.com");
            emailList.Add("guamuchilplaza@mercadodeimportaciones.com");
            emailList.Add("guasavecentro@mercadodeimportaciones.com");
            emailList.Add("guaymas@mercadodeimportaciones.com");
            emailList.Add("hermosillocentro@mercadodeimportaciones.com");
            emailList.Add("hermosillosendero@mercadodeimportaciones.com");
            emailList.Add("mochis@mercadodeimportaciones.com");
            emailList.Add("navojoa@mercadodeimportaciones.com");
            emailList.Add("navolato@mercadodeimportaciones.com");
            emailList.Add("obregon@mercadodeimportaciones.com");
            emailList.Add("guasaveplaza@mercadodeimportaciones.com");
            emailList.Add("luis.guerrero@mercadodeimportaciones.com");
            emailList.Add("karmina.alcala@mercadodeimportaciones.com");
            emailList.Add("carrasco@mercadodeimportaciones.com");
            emailList.Add("escobedo@mercadodeimportaciones.com");
            emailList.Add("luis.cota@mercadodeimportaciones.com");
            emailList.Add("mario.serrano@mercadodeimportaciones.com");
            emailList.Add("eduardope01@hotmail.com");
            emailList.Add("annacelia.soto@mercadodeimportaciones.com");
            emailList.Add("alberto.martinez@mercadodeimportaciones.com");
            emailList.Add("recepcion@mercadodeimportaciones.com");
            emailList.Add("rubi@mercadodeimportaciones.com");
            emailList.Add("hermosillomonterrey@mercadodeimportaciones.com");
            emailList.Add("ladiesobregon@mercadodeimportaciones.com");
            emailList.Add("guaymas.serdan@mercadodeimportaciones.com");
            emailList.Add("monica.perez@mercadodeimportaciones.com");
            emailList.Add("terranova@mercadodeimportaciones.com");
            emailList.Add("asofia.mercadodeimportaciones@gmail.com");
            emailList.Add("cabosendero@mercadodeimportaciones.com");
            emailList.Add("jassiel.perez@mercadodeimportaciones.com");
            emailList.Add("soporte.sinaloa@mercadodeimportaciones.com");
            emailList.Add("mazatlanelmar@mercadodeimportaciones.com");
            emailList.Add("jesus.suarez@mercadodeimportaciones.com");
            emailList.Add("senderos.culiacan@mercadodeimportaciones.com");
            emailList.Add("abastos@mercadodeimportaciones.com");
            emailList.Add("mazatlanaquiles@mercadodeimportaciones.com");
            emailList.Add("mercadotecnia@mercadodeimportaciones.com");
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            emailList.Add("barrancos@mercadodeimportaciones.com");
            emailList.Add("mochissendero@mercadodeimportaciones.com");
            emailList.Add("obregonsendero@mercadodeimportaciones.com");
            emailList.Add("san.isidro@mercadodeimportaciones.com");
            emailList.Add("hermosillomatamoros@mercadodeimportaciones.com");
            emailList.Add("jesus.salazar@mercadodeimportaciones.com");
            emailList.Add("huatabampo@mercadodeimportaciones.com");
            emailList.Add("silvia.bojorquez@mercadodeimportaciones.com");
            emailList.Add("cabosendero2@mercadodeimportaciones.com");
            emailList.Add("auxiliar.sistemas@mercadodeimportaciones.com");
            emailList.Add("maferperezle01@gmail.com");
            emailList.Add("auxiliar.ventas@mercadodeimportaciones.com");
            emailList.Add("lapaz@mercadodeimportaciones.com");
            emailList.Add("obregonplaza@mercadodeimportaciones.com");
            emailList.Add("gerencia.sistemas@mercadodeimportaciones.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_ComparativoPresupuesto.Enabled = false;
            this.EnviaMailGmail(this.ComparativoPresupuesto(fechaReporte), destinatarios);
            //this.EnviaMail(this.ComparativoPresupuesto(fechaReporte), destinatarios);
            btn_ComparativoPresupuesto.Enabled = true;
            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
        }

        private void btn_CierreCedis_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Cierre CEDIS";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("guillermina.cervantes@mercadodeimportaciones.com");
            emailList.Add("cedis@mercadodeimportaciones.com");
            emailList.Add("cedis.recibo@mercadodeimportaciones.com");
            emailList.Add("cedis.inventarios@mercadodeimportaciones.com");
            emailList.Add("monica.perez@mercadodeimportaciones.com");
            emailList.Add("jassiel.perez@mercadodeimportaciones.com");
            emailList.Add("gilberto.govea@mercadodeimportaciones.com");
            emailList.Add("hibrajid.lara@mercadodeimportaciones.com");
            emailList.Add("gerencia.sistemas@mercadodeimportaciones.com");

            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generacion
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_CierreCedis.Enabled = false;
            this.EnviaMailGmail(this.CierreCedis(fechaReporte), destinatarios);
            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_CierreCedis.Enabled = true;
        }

        private void btn_IndicadorPresupuesto_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Indicador de Presupuesto";

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generacion
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_IndicadorPresupuesto.Enabled = false;
            this.Presupuesto();
            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_IndicadorPresupuesto.Enabled = true;
        }

        private void btn_ExistenciasDiasVentas_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Existencias En Dias Ventas";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            emailList.Add("guillermina.cervantes@mercadodeimportaciones.com");
            emailList.Add("annacelia.soto@mercadodeimportaciones.com");
            emailList.Add("luis.cota@mercadodeimportaciones.com");
            emailList.Add("mario.serrano@mercadodeimportaciones.com");
            emailList.Add("Gerencia.Sistemas@mercadodeimportaciones.com");
            emailList.Add("monica.perez@mercadodeimportaciones.com");
            emailList.Add("jassiel.perez@mercadodeimportaciones.com");
            emailList.Add("gilberto.govea@mercadodeimportaciones.com");
            emailList.Add("alberto.martinez@mercadodeimportaciones.com");
            emailList.Add("luis.guerrero@mercadodeimportaciones.com");
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            emailList.Add("jesus.suarez@mercadodeimportaciones.com");
            emailList.Add("hibrajid.lara@mercadodeimportaciones.com");
            emailList.Add("maferperezle01@gmail.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_ExistenciasDiasVentas.Enabled = false;
            this.EnviaMailGmail(this.Existencias_dias_ventas(fechaReporte), destinatarios);
            btn_ExistenciasDiasVentas.Enabled = true;
            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
        }

        private void btn_VentaArticulos30Dias_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Venta De Articulos 30 Dias Venta";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;


            // destinatarios
            emailList.Clear();
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Este Reporte se genera únicamente los días LUNES.\n ¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_VentaArticulos30Dias.Enabled = false;

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("select mail.cod_estab,substring(replace(upper(establecimientos.nombre),'.',''),1,30) as nombre,mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ( '1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    this.EnviaMailGmail(this.Ventas_Articulos_30(fechaReporte, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), $"{dr1["email"].ToString()}, {destinatarios}");
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_VentaArticulos30Dias.Enabled = true;

        }

        private void btn_ComparativoSemanaSemana_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Comparativo semana-semana";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Este Reporte se genera únicamente los días LUNES.\n ¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_ComparativoSemanaSemana.Enabled = false;

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("select mail.cod_estab,substring(replace(upper(establecimientos.nombre),'.',''),1,30) as nombre,mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ( '1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    this.EnviaMailGmail(this.Comparativo_semana_semana(fechaReporte, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), $"{dr1["email"].ToString()}, {destinatarios}");
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_ComparativoSemanaSemana.Enabled = true;

        }

        private void btn_AcumuladoVentasMensual_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Acumulado De Ventas Mensual";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            emailList.Add("mercadotecnia@mercadodeimportaciones.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_AcumuladoVentasMensual.Enabled = false;

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("select mail.cod_estab,substring(replace(upper(establecimientos.nombre),'.',''),1,30) as nombre,mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ( '1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    this.EnviaMailGmail(this.Acumulado_Ventas_Mensual(fechaReporte, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), $"{dr1["email"].ToString()}, {destinatarios}");
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_AcumuladoVentasMensual.Enabled = true;
        }

        private void btn_SenalizacionPromociones_Click(object sender, EventArgs e)
        {

            string nombreReporte = "Reporte Señalizacion de Promociones";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("alberto.martinez@mercadodeimportaciones.com");
            emailList.Add("luis.guerrero@mercadodeimportaciones.com");
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            emailList.Add("jesus.salazar@mercadodeimportaciones.com");
            emailList.Add("jesus.suarez@mercadodeimportaciones.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Este Reporte se genera únicamente los días LUNES.\n ¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_SenalizacionPromociones.Enabled = false;

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("SELECT STUFF((SELECT ',' + mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int) FOR XML PATH('')),1,1, '') as email", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    this.EnviaMailGmail(this.Señalizacion_promociones(fechaReporte), $"{dr1["email"].ToString()}, {destinatarios}");
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_SenalizacionPromociones.Enabled = true;

        }

        private void btn_DesplazamientoTemporada_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Desplazamiento De Temporada";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("alberto.martinez@mercadodeimportaciones.com");
            emailList.Add("luis.guerrero@mercadodeimportaciones.com");
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            emailList.Add("jesus.salazar@mercadodeimportaciones.com");
            emailList.Add("jesus.suarez@mercadodeimportaciones.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Este Reporte se genera únicamente los días LUNES.\n ¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_DesplazamientoTemporada.Enabled = false;

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("SELECT STUFF((SELECT ',' + mail.email from establecimientos inner join dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int) FOR XML PATH('')),1,1, '') as email", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    this.EnviaMailGmail(this.Desplazamiento_Temporada(fechaReporte), $"{dr1["email"].ToString()}, {destinatarios}");
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_DesplazamientoTemporada.Enabled = true;
        }

        private void btn_ArticulosMenosVendidos_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte Articulos Menos Vendidos";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("analista.comercial@mercadodeimportaciones.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Este Reporte se genera únicamente los días LUNES y VIERNES.\n ¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_ArticulosMenosVendidos.Enabled = false;

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("select mail.cod_estab, substring(replace(upper(establecimientos.nombre), '.', ''), 1, 30) as nombre, mail.email from BMSNayar.dbo.establecimientos inner join BMSNayar.dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.status = 'V' and establecimientos.cod_estab not in ('1', '39', '43', '48', '79', '98', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by CAST(mail.cod_estab as int)", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    this.EnviaMailGmail(this.Menos_vendidos(fechaReporte, dr1["cod_estab"].ToString(), dr1["nombre"].ToString()), $"{dr1["email"].ToString()}, {destinatarios}");
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_ArticulosMenosVendidos.Enabled = true;
        }

        private void btn_CancelacionesDevoluciones_Click(object sender, EventArgs e)
        {
            string nombreReporte = "Reporte de Cancelaciones y Devoluciones";
            DateTime fechaReporte = dtFechaReporte.Value.Date + dtFechaReporte.Value.TimeOfDay;

            // destinatarios
            emailList.Clear();
            emailList.Add("luis.guerrero@mercadodeimportaciones.com");
            emailList.Add("maferperezle01@gmail.com");
            string destinatarios = string.Join(",", emailList);

            // confirmacion
            DialogResult dialogResult = MessageBox.Show($"Este Reporte se genera únicamente los días LUNES.\n ¿Desea generar {nombreReporte} con fecha {dtFechaReporte.Value.ToString("dddd, dd MMMM yyyy")} ?", "Confirmación", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) return;

            // generar reporte
            this.lblEstado.Text = $"Enviando {nombreReporte}...";
            btn_CancelacionesDevoluciones.Enabled = false;

            SqlConnection cn1 = conexion.conectar("BMSNayar");
            SqlCommand comando1 = new SqlCommand("select distinct mail.email_cordinador, coordinador from BMSNayar.dbo.establecimientos inner join BMSNayar.dbo.MI_Estab_Mail() as mail on establecimientos.cod_estab = mail.cod_estab where establecimientos.[status] = 'V' and establecimientos.cod_estab not in ('1', '10', '27', '37', '39', '43', '48', '72', '79', '101', '102', '104', '105', '106', '107', '108', '1001', '1002', '1003', '1004', '1005', '1006') order by email_cordinador", cn1);
            SqlDataReader dr1 = comando1.ExecuteReader();
            if (dr1.HasRows)
            {
                while (dr1.Read())
                {
                    this.EnviaMailGmail(this.DevolucionesCancelacionesSucursal(dr1["email_cordinador"].ToString(), dr1["coordinador"].ToString(), fechaReporte), $"{dr1["email_cordinador"].ToString()}, {destinatarios}");
                }
            }
            if (dr1.IsClosed == false) { dr1.Close(); }
            if (cn1.State == ConnectionState.Open) { cn1.Close(); }

            this.lblEstado.Text = $"Proceso de envío de {nombreReporte} - FINALIZADO EXITOSAMENTE.";
            btn_CancelacionesDevoluciones.Enabled = true;

        }

        #endregion
    }
}




namespace EnviaReportes
{
    partial class FrmPrimcipal
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPrimcipal));
            this.DG1 = new System.Windows.Forms.DataGridView();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.lblEstado = new System.Windows.Forms.Label();
            this.btn_ReporteApertura = new System.Windows.Forms.Button();
            this.btn_ReporteVentasDescuento = new System.Windows.Forms.Button();
            this.btn_ReporteIncidencias = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_Top80 = new System.Windows.Forms.Button();
            this.btn_ComparativoPresupuesto = new System.Windows.Forms.Button();
            this.btn_CierreCedis = new System.Windows.Forms.Button();
            this.lblError = new System.Windows.Forms.Label();
            this.btn_IndicadorPresupuesto = new System.Windows.Forms.Button();
            this.btn_ExistenciasDiasVentas = new System.Windows.Forms.Button();
            this.btn_VentaArticulos30Dias = new System.Windows.Forms.Button();
            this.btn_ComparativoSemanaSemana = new System.Windows.Forms.Button();
            this.btn_AcumuladoVentasMensual = new System.Windows.Forms.Button();
            this.btn_SenalizacionPromociones = new System.Windows.Forms.Button();
            this.btn_DesplazamientoTemporada = new System.Windows.Forms.Button();
            this.btn_ArticulosMenosVendidos = new System.Windows.Forms.Button();
            this.btn_CancelacionesDevoluciones = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.groupBox14 = new System.Windows.Forms.GroupBox();
            this.dtFechaReporte = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.DG1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.groupBox12.SuspendLayout();
            this.groupBox13.SuspendLayout();
            this.groupBox14.SuspendLayout();
            this.SuspendLayout();
            // 
            // DG1
            // 
            this.DG1.AllowUserToAddRows = false;
            this.DG1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.DG1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DG1.Location = new System.Drawing.Point(403, 17);
            this.DG1.Name = "DG1";
            this.DG1.ReadOnly = true;
            this.DG1.Size = new System.Drawing.Size(659, 483);
            this.DG1.TabIndex = 0;
            this.DG1.Visible = false;
            // 
            // timer1
            // 
            this.timer1.Interval = 60000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // lblEstado
            // 
            this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblEstado.Location = new System.Drawing.Point(12, 519);
            this.lblEstado.Name = "lblEstado";
            this.lblEstado.Size = new System.Drawing.Size(1050, 23);
            this.lblEstado.TabIndex = 2;
            // 
            // btn_ReporteApertura
            // 
            this.btn_ReporteApertura.Location = new System.Drawing.Point(10, 19);
            this.btn_ReporteApertura.Name = "btn_ReporteApertura";
            this.btn_ReporteApertura.Size = new System.Drawing.Size(164, 23);
            this.btn_ReporteApertura.TabIndex = 3;
            this.btn_ReporteApertura.Text = "Apertura y Cierre";
            this.btn_ReporteApertura.UseVisualStyleBackColor = true;
            this.btn_ReporteApertura.Click += new System.EventHandler(this.btn_ReporteApertura_Click);
            // 
            // btn_ReporteVentasDescuento
            // 
            this.btn_ReporteVentasDescuento.Location = new System.Drawing.Point(10, 48);
            this.btn_ReporteVentasDescuento.Name = "btn_ReporteVentasDescuento";
            this.btn_ReporteVentasDescuento.Size = new System.Drawing.Size(164, 23);
            this.btn_ReporteVentasDescuento.TabIndex = 4;
            this.btn_ReporteVentasDescuento.Text = "Ventas con Descuento";
            this.btn_ReporteVentasDescuento.UseVisualStyleBackColor = true;
            this.btn_ReporteVentasDescuento.Click += new System.EventHandler(this.btn_ReporteVentasDescuento_Click);
            // 
            // btn_ReporteIncidencias
            // 
            this.btn_ReporteIncidencias.Location = new System.Drawing.Point(10, 77);
            this.btn_ReporteIncidencias.Name = "btn_ReporteIncidencias";
            this.btn_ReporteIncidencias.Size = new System.Drawing.Size(164, 23);
            this.btn_ReporteIncidencias.TabIndex = 5;
            this.btn_ReporteIncidencias.Text = "Incidencias del Establecimiento";
            this.btn_ReporteIncidencias.UseVisualStyleBackColor = true;
            this.btn_ReporteIncidencias.Click += new System.EventHandler(this.btn_ReporteIncidencias_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dtFechaReporte);
            this.groupBox1.Location = new System.Drawing.Point(12, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(222, 60);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Fecha Generación de Reportes";
            // 
            // btn_Top80
            // 
            this.btn_Top80.Location = new System.Drawing.Point(8, 18);
            this.btn_Top80.Name = "btn_Top80";
            this.btn_Top80.Size = new System.Drawing.Size(164, 23);
            this.btn_Top80.TabIndex = 7;
            this.btn_Top80.Text = "Top 80";
            this.btn_Top80.UseVisualStyleBackColor = true;
            this.btn_Top80.Click += new System.EventHandler(this.btn_Top80_Click);
            // 
            // btn_ComparativoPresupuesto
            // 
            this.btn_ComparativoPresupuesto.Location = new System.Drawing.Point(6, 18);
            this.btn_ComparativoPresupuesto.Name = "btn_ComparativoPresupuesto";
            this.btn_ComparativoPresupuesto.Size = new System.Drawing.Size(164, 23);
            this.btn_ComparativoPresupuesto.TabIndex = 8;
            this.btn_ComparativoPresupuesto.Text = "Comparativo vs Presupuesto";
            this.btn_ComparativoPresupuesto.UseVisualStyleBackColor = true;
            this.btn_ComparativoPresupuesto.Click += new System.EventHandler(this.btn_ComparativoPresupuesto_Click);
            // 
            // btn_CierreCedis
            // 
            this.btn_CierreCedis.Location = new System.Drawing.Point(6, 16);
            this.btn_CierreCedis.Name = "btn_CierreCedis";
            this.btn_CierreCedis.Size = new System.Drawing.Size(164, 23);
            this.btn_CierreCedis.TabIndex = 9;
            this.btn_CierreCedis.Text = "Cierre CEDIS";
            this.btn_CierreCedis.UseVisualStyleBackColor = true;
            this.btn_CierreCedis.Click += new System.EventHandler(this.btn_CierreCedis_Click);
            // 
            // lblError
            // 
            this.lblError.AutoSize = true;
            this.lblError.ForeColor = System.Drawing.Color.Red;
            this.lblError.Location = new System.Drawing.Point(15, 390);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(0, 13);
            this.lblError.TabIndex = 10;
            // 
            // btn_IndicadorPresupuesto
            // 
            this.btn_IndicadorPresupuesto.Location = new System.Drawing.Point(6, 16);
            this.btn_IndicadorPresupuesto.Name = "btn_IndicadorPresupuesto";
            this.btn_IndicadorPresupuesto.Size = new System.Drawing.Size(164, 23);
            this.btn_IndicadorPresupuesto.TabIndex = 11;
            this.btn_IndicadorPresupuesto.Text = "Indicador de Presupuesto";
            this.btn_IndicadorPresupuesto.UseVisualStyleBackColor = true;
            this.btn_IndicadorPresupuesto.Click += new System.EventHandler(this.btn_IndicadorPresupuesto_Click);
            // 
            // btn_ExistenciasDiasVentas
            // 
            this.btn_ExistenciasDiasVentas.Location = new System.Drawing.Point(6, 19);
            this.btn_ExistenciasDiasVentas.Name = "btn_ExistenciasDiasVentas";
            this.btn_ExistenciasDiasVentas.Size = new System.Drawing.Size(164, 23);
            this.btn_ExistenciasDiasVentas.TabIndex = 12;
            this.btn_ExistenciasDiasVentas.Text = "Existencias En Dias Ventas";
            this.btn_ExistenciasDiasVentas.UseVisualStyleBackColor = true;
            this.btn_ExistenciasDiasVentas.Click += new System.EventHandler(this.btn_ExistenciasDiasVentas_Click);
            // 
            // btn_VentaArticulos30Dias
            // 
            this.btn_VentaArticulos30Dias.Location = new System.Drawing.Point(10, 15);
            this.btn_VentaArticulos30Dias.Name = "btn_VentaArticulos30Dias";
            this.btn_VentaArticulos30Dias.Size = new System.Drawing.Size(164, 23);
            this.btn_VentaArticulos30Dias.TabIndex = 13;
            this.btn_VentaArticulos30Dias.Text = "Venta Articulos 30 Dias";
            this.btn_VentaArticulos30Dias.UseVisualStyleBackColor = true;
            this.btn_VentaArticulos30Dias.Click += new System.EventHandler(this.btn_VentaArticulos30Dias_Click);
            // 
            // btn_ComparativoSemanaSemana
            // 
            this.btn_ComparativoSemanaSemana.Location = new System.Drawing.Point(6, 18);
            this.btn_ComparativoSemanaSemana.Name = "btn_ComparativoSemanaSemana";
            this.btn_ComparativoSemanaSemana.Size = new System.Drawing.Size(164, 23);
            this.btn_ComparativoSemanaSemana.TabIndex = 14;
            this.btn_ComparativoSemanaSemana.Text = "Comparativo semana-semana";
            this.btn_ComparativoSemanaSemana.UseVisualStyleBackColor = true;
            this.btn_ComparativoSemanaSemana.Click += new System.EventHandler(this.btn_ComparativoSemanaSemana_Click);
            // 
            // btn_AcumuladoVentasMensual
            // 
            this.btn_AcumuladoVentasMensual.Location = new System.Drawing.Point(6, 14);
            this.btn_AcumuladoVentasMensual.Name = "btn_AcumuladoVentasMensual";
            this.btn_AcumuladoVentasMensual.Size = new System.Drawing.Size(164, 23);
            this.btn_AcumuladoVentasMensual.TabIndex = 15;
            this.btn_AcumuladoVentasMensual.Text = "Acumulado Ventas Mensual";
            this.btn_AcumuladoVentasMensual.UseVisualStyleBackColor = true;
            this.btn_AcumuladoVentasMensual.Click += new System.EventHandler(this.btn_AcumuladoVentasMensual_Click);
            // 
            // btn_SenalizacionPromociones
            // 
            this.btn_SenalizacionPromociones.Location = new System.Drawing.Point(6, 15);
            this.btn_SenalizacionPromociones.Name = "btn_SenalizacionPromociones";
            this.btn_SenalizacionPromociones.Size = new System.Drawing.Size(164, 23);
            this.btn_SenalizacionPromociones.TabIndex = 16;
            this.btn_SenalizacionPromociones.Text = "Señalizacion Promociones";
            this.btn_SenalizacionPromociones.UseVisualStyleBackColor = true;
            this.btn_SenalizacionPromociones.Click += new System.EventHandler(this.btn_SenalizacionPromociones_Click);
            // 
            // btn_DesplazamientoTemporada
            // 
            this.btn_DesplazamientoTemporada.Location = new System.Drawing.Point(6, 15);
            this.btn_DesplazamientoTemporada.Name = "btn_DesplazamientoTemporada";
            this.btn_DesplazamientoTemporada.Size = new System.Drawing.Size(164, 23);
            this.btn_DesplazamientoTemporada.TabIndex = 17;
            this.btn_DesplazamientoTemporada.Text = " Desplazamiento De Temporada";
            this.btn_DesplazamientoTemporada.UseVisualStyleBackColor = true;
            this.btn_DesplazamientoTemporada.Click += new System.EventHandler(this.btn_DesplazamientoTemporada_Click);
            // 
            // btn_ArticulosMenosVendidos
            // 
            this.btn_ArticulosMenosVendidos.Location = new System.Drawing.Point(6, 19);
            this.btn_ArticulosMenosVendidos.Name = "btn_ArticulosMenosVendidos";
            this.btn_ArticulosMenosVendidos.Size = new System.Drawing.Size(164, 23);
            this.btn_ArticulosMenosVendidos.TabIndex = 18;
            this.btn_ArticulosMenosVendidos.Text = "Articulos Menos Vendidos";
            this.btn_ArticulosMenosVendidos.UseVisualStyleBackColor = true;
            this.btn_ArticulosMenosVendidos.Click += new System.EventHandler(this.btn_ArticulosMenosVendidos_Click);
            // 
            // btn_CancelacionesDevoluciones
            // 
            this.btn_CancelacionesDevoluciones.Location = new System.Drawing.Point(6, 16);
            this.btn_CancelacionesDevoluciones.Name = "btn_CancelacionesDevoluciones";
            this.btn_CancelacionesDevoluciones.Size = new System.Drawing.Size(164, 23);
            this.btn_CancelacionesDevoluciones.TabIndex = 19;
            this.btn_CancelacionesDevoluciones.Text = "Cancelaciones y Devoluciones";
            this.btn_CancelacionesDevoluciones.UseVisualStyleBackColor = true;
            this.btn_CancelacionesDevoluciones.Click += new System.EventHandler(this.btn_CancelacionesDevoluciones_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_ReporteApertura);
            this.groupBox2.Controls.Add(this.btn_ReporteVentasDescuento);
            this.groupBox2.Controls.Add(this.btn_ReporteIncidencias);
            this.groupBox2.Location = new System.Drawing.Point(12, 83);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(179, 104);
            this.groupBox2.TabIndex = 20;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "01:01";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btn_Top80);
            this.groupBox3.Location = new System.Drawing.Point(12, 189);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(179, 50);
            this.groupBox3.TabIndex = 21;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "01:30";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btn_VentaArticulos30Dias);
            this.groupBox4.Location = new System.Drawing.Point(12, 245);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(179, 48);
            this.groupBox4.TabIndex = 22;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "05:00";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btn_ComparativoSemanaSemana);
            this.groupBox5.Location = new System.Drawing.Point(18, 299);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(173, 48);
            this.groupBox5.TabIndex = 23;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "05:40";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.btn_ExistenciasDiasVentas);
            this.groupBox6.Location = new System.Drawing.Point(18, 353);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(173, 48);
            this.groupBox6.TabIndex = 24;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "06:00";
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.btn_AcumuladoVentasMensual);
            this.groupBox7.Location = new System.Drawing.Point(18, 407);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(173, 49);
            this.groupBox7.TabIndex = 25;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "06:15";
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.btn_SenalizacionPromociones);
            this.groupBox8.Location = new System.Drawing.Point(18, 462);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(173, 48);
            this.groupBox8.TabIndex = 26;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "06:30";
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.btn_DesplazamientoTemporada);
            this.groupBox9.Location = new System.Drawing.Point(197, 83);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(184, 47);
            this.groupBox9.TabIndex = 27;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "06:40";
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.btn_ArticulosMenosVendidos);
            this.groupBox10.Location = new System.Drawing.Point(197, 138);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(184, 52);
            this.groupBox10.TabIndex = 28;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "06:45";
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.btn_CancelacionesDevoluciones);
            this.groupBox11.Location = new System.Drawing.Point(197, 196);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(184, 49);
            this.groupBox11.TabIndex = 29;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "07:00";
            // 
            // groupBox12
            // 
            this.groupBox12.Controls.Add(this.btn_ComparativoPresupuesto);
            this.groupBox12.Location = new System.Drawing.Point(197, 304);
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.Size = new System.Drawing.Size(181, 51);
            this.groupBox12.TabIndex = 30;
            this.groupBox12.TabStop = false;
            this.groupBox12.Text = "12:00  -  15:00  -  18:00  -  22:35";
            // 
            // groupBox13
            // 
            this.groupBox13.Controls.Add(this.btn_CierreCedis);
            this.groupBox13.Location = new System.Drawing.Point(197, 251);
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.Size = new System.Drawing.Size(181, 47);
            this.groupBox13.TabIndex = 31;
            this.groupBox13.TabStop = false;
            this.groupBox13.Text = "17:00  -  22:00";
            // 
            // groupBox14
            // 
            this.groupBox14.Controls.Add(this.btn_IndicadorPresupuesto);
            this.groupBox14.Location = new System.Drawing.Point(197, 361);
            this.groupBox14.Name = "groupBox14";
            this.groupBox14.Size = new System.Drawing.Size(181, 49);
            this.groupBox14.TabIndex = 32;
            this.groupBox14.TabStop = false;
            this.groupBox14.Text = "22:45";
            // 
            // dtFechaReporte
            // 
            this.dtFechaReporte.CustomFormat = "MMMMdd, yyyy  |  hh:mm";
            this.dtFechaReporte.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtFechaReporte.Location = new System.Drawing.Point(12, 19);
            this.dtFechaReporte.Name = "dtFechaReporte";
            this.dtFechaReporte.Size = new System.Drawing.Size(176, 20);
            this.dtFechaReporte.TabIndex = 33;
            // 
            // FrmPrimcipal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1074, 551);
            this.Controls.Add(this.groupBox14);
            this.Controls.Add(this.groupBox13);
            this.Controls.Add(this.groupBox12);
            this.Controls.Add(this.groupBox11);
            this.Controls.Add(this.groupBox10);
            this.Controls.Add(this.groupBox9);
            this.Controls.Add(this.groupBox8);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.lblError);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblEstado);
            this.Controls.Add(this.DG1);
            this.Controls.Add(this.groupBox2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmPrimcipal";
            this.Text = "Envía Reportes";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DG1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox8.ResumeLayout(false);
            this.groupBox9.ResumeLayout(false);
            this.groupBox10.ResumeLayout(false);
            this.groupBox11.ResumeLayout(false);
            this.groupBox12.ResumeLayout(false);
            this.groupBox13.ResumeLayout(false);
            this.groupBox14.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView DG1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label lblEstado;
        private System.Windows.Forms.Button btn_ReporteApertura;
        private System.Windows.Forms.Button btn_ReporteVentasDescuento;
        private System.Windows.Forms.Button btn_ReporteIncidencias;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_Top80;
        private System.Windows.Forms.Button btn_ComparativoPresupuesto;
        private System.Windows.Forms.Button btn_CierreCedis;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.Button btn_IndicadorPresupuesto;
        private System.Windows.Forms.Button btn_ExistenciasDiasVentas;
        private System.Windows.Forms.Button btn_VentaArticulos30Dias;
        private System.Windows.Forms.Button btn_ComparativoSemanaSemana;
        private System.Windows.Forms.Button btn_AcumuladoVentasMensual;
        private System.Windows.Forms.Button btn_SenalizacionPromociones;
        private System.Windows.Forms.Button btn_DesplazamientoTemporada;
        private System.Windows.Forms.Button btn_ArticulosMenosVendidos;
        private System.Windows.Forms.Button btn_CancelacionesDevoluciones;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.GroupBox groupBox10;
        private System.Windows.Forms.GroupBox groupBox11;
        private System.Windows.Forms.GroupBox groupBox12;
        private System.Windows.Forms.GroupBox groupBox13;
        private System.Windows.Forms.GroupBox groupBox14;
        private System.Windows.Forms.DateTimePicker dtFechaReporte;
    }
}


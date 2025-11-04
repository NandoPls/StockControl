using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Windows.Forms;
using System.IO;
using System.Text.Json;

namespace Inventario
{
    public partial class InventarioForm : Form
    {
        private Dictionary<string, ProductoExcel> productosParaInventariar; // Key: ItemCode
        private Dictionary<string, int> conteoActual; // Key: ItemCode
        private Dictionary<string, List<string>> eanToItemCodes; // Lookup: EAN -> List of ItemCodes
        private string almacenSeleccionado;
        private List<string> clasificacionesSeleccionadas;

        private ComboBox cboAlmacen;
        private RadioButton rbTipo;
        private RadioButton rbSubtipo;
        private CheckedListBox clbClasificaciones;
        private CheckedListBox clbSubtipos;
        private TextBox txtEscaneo;
        private bool usarSubtipo = false;
        private Label lblInstruccion;
        private Label lblProgreso;
        private DataGridView dgvInventario;
        private Button btnIniciar;
        private Button btnFinalizar;
        private Button btnGenerarReporte;
        private Button btnCancelar;
        private Panel panelSeleccion;
        private Panel panelEscaneo;

        // Nuevos controles para mejoras v1.1.0
        private TextBox txtBusqueda;
        private CheckBox chkSoloDiferencias;
        private Panel panelEstadisticas;
        private Label lblTotalProductos;
        private Label lblProductosCorrectos;
        private Label lblProductosSobrantes;
        private Label lblProductosFaltantes;
        private Label lblPorcentajeAvance;
        private System.Windows.Forms.Timer timerAutoguardado;
        private string archivoSesion = "";
        private string archivoExcelRespaldo = ""; // Ruta del último Excel generado

        public InventarioForm()
        {
            InitializeComponent();
            InicializarComponentes();
            productosParaInventariar = new Dictionary<string, ProductoExcel>();
            conteoActual = new Dictionary<string, int>();
            eanToItemCodes = new Dictionary<string, List<string>>();
            clasificacionesSeleccionadas = new List<string>();
            ConfigurarIcono();
            ConfigurarAutoguardado();
            VerificarSesionPendiente();
        }

        private void ConfigurarIcono()
        {
            try
            {
                int iconSize = 64;
                using (Bitmap bitmap = new Bitmap(iconSize, iconSize))
                {
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                        Color blueColor = Color.FromArgb(41, 128, 185);
                        using (SolidBrush brush = new SolidBrush(blueColor))
                        {
                            graphics.FillEllipse(brush, 0, 0, iconSize, iconSize);
                        }

                        using (Pen whitePen = new Pen(Color.White, 4))
                        {
                            graphics.DrawRectangle(whitePen, 10, 12, 44, 40);
                        }

                        using (Pen whitePen2 = new Pen(Color.White, 2))
                        {
                            graphics.DrawLine(whitePen2, 10, 25, 54, 25);
                            graphics.DrawLine(whitePen2, 10, 38, 54, 38);
                        }

                        Color greenColor = Color.FromArgb(46, 204, 113);
                        using (Pen checkPen = new Pen(greenColor, 6))
                        {
                            checkPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                            checkPen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                            graphics.DrawLine(checkPen, 35, 42, 42, 50);
                            graphics.DrawLine(checkPen, 42, 50, 55, 32);
                        }
                    }

                    this.Icon = Icon.FromHandle(bitmap.GetHicon());
                }
            }
            catch
            {
                // Si hay error, continuar sin icono
            }
        }

        private void InicializarComponentes()
        {
            this.Text = "StockControl - Conteo de Inventario";
            this.Size = new Size(1000, 735);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);

            // Panel de selección (visible al inicio)
            panelSeleccion = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(940, 300),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            Label lblTitulo = new Label
            {
                Text = "📦 StockControl - Configuración del Inventario",
                Location = new Point(20, 20),
                Size = new Size(600, 30),
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204)
            };
            panelSeleccion.Controls.Add(lblTitulo);

            // Selección de almacén
            Label lblAlmacen = new Label
            {
                Text = "Seleccione el Almacén:",
                Location = new Point(30, 70),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };
            panelSeleccion.Controls.Add(lblAlmacen);

            cboAlmacen = new ComboBox
            {
                Location = new Point(240, 68),
                Size = new Size(300, 30),
                Font = new Font("Segoe UI", 11),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboAlmacen.SelectedIndexChanged += CboAlmacen_SelectedIndexChanged;
            panelSeleccion.Controls.Add(cboAlmacen);

            // RadioButtons para elegir entre Tipo o Subtipo
            Label lblModoClasificacion = new Label
            {
                Text = "Clasificar por:",
                Location = new Point(30, 110),
                Size = new Size(120, 25),
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };
            panelSeleccion.Controls.Add(lblModoClasificacion);

            rbTipo = new RadioButton
            {
                Text = "📦 Tipo (Marca)",
                Location = new Point(160, 110),
                Size = new Size(150, 25),
                Font = new Font("Segoe UI", 10),
                Checked = true,
                Enabled = false
            };
            rbTipo.CheckedChanged += RbTipo_CheckedChanged;
            panelSeleccion.Controls.Add(rbTipo);

            rbSubtipo = new RadioButton
            {
                Text = "📋 Subtipo (Detalle)",
                Location = new Point(320, 110),
                Size = new Size(180, 25),
                Font = new Font("Segoe UI", 10),
                Enabled = false
            };
            rbSubtipo.CheckedChanged += RbSubtipo_CheckedChanged;
            panelSeleccion.Controls.Add(rbSubtipo);

            // Selección de clasificaciones por Tipo (múltiple)
            Label lblClasificacion = new Label
            {
                Text = "Seleccione Clasificaciones (puede marcar varias):",
                Location = new Point(30, 145),
                Size = new Size(400, 25),
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };
            panelSeleccion.Controls.Add(lblClasificacion);

            clbClasificaciones = new CheckedListBox
            {
                Location = new Point(30, 175),
                Size = new Size(600, 100),
                Font = new Font("Segoe UI", 10),
                CheckOnClick = true,
                Enabled = false,
                BackColor = Color.FromArgb(250, 250, 250),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = true
            };
            clbClasificaciones.ItemCheck += ClbClasificaciones_ItemCheck;
            panelSeleccion.Controls.Add(clbClasificaciones);

            // CheckedListBox para Subtipos (inicialmente oculto)
            clbSubtipos = new CheckedListBox
            {
                Location = new Point(30, 175),
                Size = new Size(600, 100),
                Font = new Font("Segoe UI", 10),
                CheckOnClick = true,
                Enabled = false,
                BackColor = Color.FromArgb(250, 250, 250),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };
            clbSubtipos.ItemCheck += ClbSubtipos_ItemCheck;
            panelSeleccion.Controls.Add(clbSubtipos);

            // Botón Iniciar
            btnIniciar = new Button
            {
                Text = "INICIAR INVENTARIO",
                Location = new Point(650, 175),
                Size = new Size(250, 50),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false,
                Cursor = Cursors.Hand
            };
            btnIniciar.FlatAppearance.BorderSize = 0;
            btnIniciar.Click += BtnIniciar_Click;
            panelSeleccion.Controls.Add(btnIniciar);

            this.Controls.Add(panelSeleccion);

            // Panel de escaneo (oculto al inicio)
            panelEscaneo = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(940, 630),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };

            Label lblTituloEscaneo = new Label
            {
                Text = "Escaneo de Productos",
                Location = new Point(20, 20),
                Size = new Size(400, 30),
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204)
            };
            panelEscaneo.Controls.Add(lblTituloEscaneo);

            lblInstruccion = new Label
            {
                Text = "Escanee el código de barras del producto:",
                Location = new Point(30, 70),
                Size = new Size(500, 25),
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };
            panelEscaneo.Controls.Add(lblInstruccion);

            txtEscaneo = new TextBox
            {
                Location = new Point(30, 100),
                Size = new Size(500, 35),
                Font = new Font("Segoe UI", 14),
                BackColor = Color.FromArgb(255, 255, 200)
            };
            txtEscaneo.KeyPress += TxtEscaneo_KeyPress;
            panelEscaneo.Controls.Add(txtEscaneo);

            lblProgreso = new Label
            {
                Text = "Productos inventariados: 0 de 0",
                Location = new Point(550, 70),
                Size = new Size(350, 60),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204),
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(230, 240, 255)
            };
            panelEscaneo.Controls.Add(lblProgreso);

            // DataGridView para mostrar inventario
            dgvInventario = new DataGridView
            {
                Location = new Point(30, 150),
                Size = new Size(880, 400),
                Font = new Font("Segoe UI", 10),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.Fixed3D
            };

            dgvInventario.Columns.Add("Marca", "Marca");
            dgvInventario.Columns.Add("Clasificacion", "Clasificación");
            dgvInventario.Columns.Add("Detalle", "Detalle");
            dgvInventario.Columns.Add("Codigo", "Código");
            dgvInventario.Columns.Add("EAN", "EAN");
            dgvInventario.Columns.Add("StockSistema", "Stock Sistema");
            dgvInventario.Columns.Add("StockContado", "Stock Contado");
            dgvInventario.Columns.Add("Diferencia", "Diferencia");

            // Hacer solo editable la columna Stock Contado
            foreach (DataGridViewColumn column in dgvInventario.Columns)
            {
                column.ReadOnly = true;
            }
            dgvInventario.Columns["StockContado"].ReadOnly = false;

            dgvInventario.Columns["StockSistema"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvInventario.Columns["StockContado"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvInventario.Columns["StockContado"].DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 220);
            dgvInventario.Columns["Diferencia"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvInventario.CellEndEdit += DgvInventario_CellEndEdit;

            panelEscaneo.Controls.Add(dgvInventario);

            // ============================================
            // PANEL INFERIOR REORGANIZADO
            // ============================================

            // FILA 1: Búsqueda y Filtros (Y = 555)
            Label lblBusqueda = new Label
            {
                Text = "🔍 Buscar:",
                Location = new Point(30, 558),
                Size = new Size(70, 25),
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            panelEscaneo.Controls.Add(lblBusqueda);

            txtBusqueda = new TextBox
            {
                Location = new Point(105, 556),
                Size = new Size(250, 25),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Código, EAN o Detalle..."
            };
            txtBusqueda.TextChanged += TxtBusqueda_TextChanged;
            panelEscaneo.Controls.Add(txtBusqueda);

            chkSoloDiferencias = new CheckBox
            {
                Text = "📊 Solo Diferencias",
                Location = new Point(380, 558),
                Size = new Size(180, 25),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69)
            };
            chkSoloDiferencias.CheckedChanged += ChkSoloDiferencias_CheckedChanged;
            panelEscaneo.Controls.Add(chkSoloDiferencias);

            // FILA 2: Botones de Acción (Y = 590 - con más separación)
            btnFinalizar = new Button
            {
                Text = "💾 FINALIZAR Y GUARDAR",
                Location = new Point(30, 590),
                Size = new Size(250, 40),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnFinalizar.FlatAppearance.BorderSize = 0;
            btnFinalizar.Click += BtnFinalizar_Click;
            panelEscaneo.Controls.Add(btnFinalizar);

            btnGenerarReporte = new Button
            {
                Text = "📧 GENERAR REPORTE",
                Location = new Point(300, 590),
                Size = new Size(250, 40),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnGenerarReporte.FlatAppearance.BorderSize = 0;
            btnGenerarReporte.Click += BtnGenerarReporte_Click;
            panelEscaneo.Controls.Add(btnGenerarReporte);

            btnCancelar = new Button
            {
                Text = "❌ CANCELAR",
                Location = new Point(570, 590),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnCancelar.FlatAppearance.BorderSize = 0;
            btnCancelar.Click += BtnCancelar_Click;
            panelEscaneo.Controls.Add(btnCancelar);

            this.Controls.Add(panelEscaneo);

            // **NUEVO: Panel de Estadísticas en Tiempo Real**
            panelEstadisticas = new Panel
            {
                Location = new Point(20, 340),
                Size = new Size(940, 200),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };

            Label lblTituloEstadisticas = new Label
            {
                Text = "📊 ESTADÍSTICAS EN TIEMPO REAL",
                Location = new Point(20, 10),
                Size = new Size(400, 30),
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 122, 204)
            };
            panelEstadisticas.Controls.Add(lblTituloEstadisticas);

            lblTotalProductos = new Label
            {
                Text = "Total de Productos\n0",
                Location = new Point(50, 60),
                Size = new Size(180, 80),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(230, 240, 255),
                BorderStyle = BorderStyle.FixedSingle,
                ForeColor = Color.FromArgb(0, 122, 204)
            };
            panelEstadisticas.Controls.Add(lblTotalProductos);

            lblProductosCorrectos = new Label
            {
                Text = "✅ Correctos\n0",
                Location = new Point(250, 60),
                Size = new Size(150, 80),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(212, 237, 218),
                BorderStyle = BorderStyle.FixedSingle,
                ForeColor = Color.FromArgb(21, 87, 36)
            };
            panelEstadisticas.Controls.Add(lblProductosCorrectos);

            lblProductosSobrantes = new Label
            {
                Text = "⬆️ Sobrantes\n0",
                Location = new Point(420, 60),
                Size = new Size(150, 80),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(255, 243, 205),
                BorderStyle = BorderStyle.FixedSingle,
                ForeColor = Color.FromArgb(133, 100, 4)
            };
            panelEstadisticas.Controls.Add(lblProductosSobrantes);

            lblProductosFaltantes = new Label
            {
                Text = "⬇️ Faltantes\n0",
                Location = new Point(590, 60),
                Size = new Size(150, 80),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(248, 215, 218),
                BorderStyle = BorderStyle.FixedSingle,
                ForeColor = Color.FromArgb(114, 28, 36)
            };
            panelEstadisticas.Controls.Add(lblProductosFaltantes);

            lblPorcentajeAvance = new Label
            {
                Text = "AVANCE: 0%",
                Location = new Point(760, 60),
                Size = new Size(150, 80),
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(230, 240, 255),
                BorderStyle = BorderStyle.FixedSingle,
                ForeColor = Color.FromArgb(0, 122, 204)
            };
            panelEstadisticas.Controls.Add(lblPorcentajeAvance);

            Label lblInfoAutoguardado = new Label
            {
                Text = "💾 Autoguardado cada 2 minutos",
                Location = new Point(320, 155),
                Size = new Size(300, 25),
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panelEstadisticas.Controls.Add(lblInfoAutoguardado);

            this.Controls.Add(panelEstadisticas);

            // Footer
            Panel panelFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 35,
                BackColor = Color.FromArgb(52, 73, 94)
            };

            Label lblFooter = new Label
            {
                Text = "StockControl v1.2.2 | Desarrollado por Fernando Carrasco",
                Location = new Point(20, 10),
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.White
            };
            panelFooter.Controls.Add(lblFooter);

            this.Controls.Add(panelFooter);

            CargarAlmacenes();
        }

        private void CargarAlmacenes()
        {
            if (ExcelDataManager.ProductosExcel == null || !ExcelDataManager.ProductosExcel.Any())
            {
                MessageBox.Show("No hay datos cargados. Por favor, cargue un archivo Excel primero.",
                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var almacenes = ExcelDataManager.ProductosExcel
                .Select(p => p.WhsCode)
                .Distinct()
                .OrderBy(a => a)
                .ToList();

            cboAlmacen.Items.Clear();
            foreach (var almacen in almacenes)
            {
                cboAlmacen.Items.Add(almacen);
            }
        }

        private void CboAlmacen_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboAlmacen.SelectedItem == null) return;

            almacenSeleccionado = cboAlmacen.SelectedItem.ToString();

            // Cargar clasificaciones por Tipo (ItmsGrpNam)
            var clasificaciones = ExcelDataManager.ProductosExcel
                .Where(p => p.WhsCode == almacenSeleccionado)
                .Select(p => p.ItmsGrpNam)
                .Distinct()
                .OrderBy(c => c)
                .ToList();

            clbClasificaciones.Items.Clear();
            foreach (var clasificacion in clasificaciones)
            {
                clbClasificaciones.Items.Add(clasificacion);
            }

            // Cargar clasificaciones por Subtipo (U_Comercial3)
            var subtipos = ExcelDataManager.ProductosExcel
                .Where(p => p.WhsCode == almacenSeleccionado && !string.IsNullOrWhiteSpace(p.U_Comercial3))
                .Select(p => p.U_Comercial3)
                .Distinct()
                .OrderBy(c => c)
                .ToList();

            clbSubtipos.Items.Clear();
            foreach (var subtipo in subtipos)
            {
                clbSubtipos.Items.Add(subtipo);
            }

            clasificacionesSeleccionadas.Clear();

            // Habilitar radiobuttons y el CheckedListBox visible
            rbTipo.Enabled = true;
            rbSubtipo.Enabled = true;

            if (rbTipo.Checked)
            {
                clbClasificaciones.Enabled = true;
            }
            else
            {
                clbSubtipos.Enabled = true;
            }

            btnIniciar.Enabled = false;
        }

        private void ClbClasificaciones_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke(new Action(() =>
            {
                clasificacionesSeleccionadas.Clear();
                foreach (var item in clbClasificaciones.CheckedItems)
                {
                    clasificacionesSeleccionadas.Add(item.ToString());
                }
                btnIniciar.Enabled = clasificacionesSeleccionadas.Count > 0;
            }));
        }

        private void ClbSubtipos_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke(new Action(() =>
            {
                clasificacionesSeleccionadas.Clear();
                foreach (var item in clbSubtipos.CheckedItems)
                {
                    clasificacionesSeleccionadas.Add(item.ToString());
                }
                btnIniciar.Enabled = clasificacionesSeleccionadas.Count > 0;
            }));
        }

        private void RbTipo_CheckedChanged(object sender, EventArgs e)
        {
            if (rbTipo.Checked)
            {
                usarSubtipo = false;

                // Mostrar CheckedListBox de Tipos, ocultar de Subtipos
                clbClasificaciones.Visible = true;
                clbClasificaciones.Enabled = true;
                clbSubtipos.Visible = false;
                clbSubtipos.Enabled = false;

                // Limpiar selección y actualizar botón
                clasificacionesSeleccionadas.Clear();
                clbSubtipos.ClearSelected();
                for (int i = 0; i < clbSubtipos.Items.Count; i++)
                {
                    clbSubtipos.SetItemChecked(i, false);
                }

                btnIniciar.Enabled = clbClasificaciones.CheckedItems.Count > 0;
            }
        }

        private void RbSubtipo_CheckedChanged(object sender, EventArgs e)
        {
            if (rbSubtipo.Checked)
            {
                usarSubtipo = true;

                // Ocultar CheckedListBox de Tipos, mostrar de Subtipos
                clbClasificaciones.Visible = false;
                clbClasificaciones.Enabled = false;
                clbSubtipos.Visible = true;
                clbSubtipos.Enabled = true;

                // Limpiar selección y actualizar botón
                clasificacionesSeleccionadas.Clear();
                clbClasificaciones.ClearSelected();
                for (int i = 0; i < clbClasificaciones.Items.Count; i++)
                {
                    clbClasificaciones.SetItemChecked(i, false);
                }

                btnIniciar.Enabled = clbSubtipos.CheckedItems.Count > 0;
            }
        }

        private void BtnIniciar_Click(object sender, EventArgs e)
        {
            if (clasificacionesSeleccionadas.Count == 0)
            {
                MessageBox.Show("Por favor, seleccione al menos una clasificación.",
                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Cargar productos a inventariar
            productosParaInventariar.Clear();
            conteoActual.Clear();
            eanToItemCodes.Clear();
            dgvInventario.Rows.Clear();

            // Cargar productos según el tipo de clasificación seleccionado
            List<ProductoExcel> productos;
            if (usarSubtipo)
            {
                // Filtrar por Subtipo (U_Comercial3)
                productos = ExcelDataManager.ProductosExcel
                    .Where(p => p.WhsCode == almacenSeleccionado && clasificacionesSeleccionadas.Contains(p.U_Comercial3))
                    .ToList();
            }
            else
            {
                // Filtrar por Tipo (ItmsGrpNam)
                productos = ExcelDataManager.ProductosExcel
                    .Where(p => p.WhsCode == almacenSeleccionado && clasificacionesSeleccionadas.Contains(p.ItmsGrpNam))
                    .ToList();
            }

            foreach (var producto in productos)
            {
                // Usar ItemCode como clave única (permite duplicados de EAN)
                string key = producto.ItemCode;

                if (!productosParaInventariar.ContainsKey(key))
                {
                    productosParaInventariar[key] = producto;
                    conteoActual[key] = 0;

                    // Mantener lookup de EAN -> ItemCode para escaneo
                    if (!eanToItemCodes.ContainsKey(producto.CodeBars))
                    {
                        eanToItemCodes[producto.CodeBars] = new List<string>();
                    }
                    eanToItemCodes[producto.CodeBars].Add(key);

                    int rowIndex = dgvInventario.Rows.Add();
                    DataGridViewRow row = dgvInventario.Rows[rowIndex];
                    row.Cells["Marca"].Value = producto.ItmsGrpNam;
                    row.Cells["Clasificacion"].Value = producto.U_Comercial1;
                    row.Cells["Detalle"].Value = producto.U_Comercial3;
                    row.Cells["Codigo"].Value = producto.ItemCode;
                    row.Cells["EAN"].Value = producto.CodeBars;
                    row.Cells["StockSistema"].Value = producto.StockTienda;
                    row.Cells["StockContado"].Value = 0;
                    row.Cells["Diferencia"].Value = -producto.StockTienda;
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 240, 240);
                }
            }

            ActualizarProgreso();
            ActualizarEstadisticas();

            // Cambiar a modo escaneo
            panelSeleccion.Visible = false;
            panelEscaneo.Visible = true;
            panelEstadisticas.Visible = true;

            // Iniciar autoguardado
            timerAutoguardado.Start();

            txtEscaneo.Focus();
        }

        private void TxtEscaneo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                ProcesarEscaneo();
            }
        }

        private void ProcesarEscaneo()
        {
            string codigoEscaneado = txtEscaneo.Text.Trim();
            txtEscaneo.Clear();

            if (string.IsNullOrEmpty(codigoEscaneado))
                return;

            // Buscar por EAN en el lookup
            if (eanToItemCodes.ContainsKey(codigoEscaneado))
            {
                List<string> itemCodes = eanToItemCodes[codigoEscaneado];

                // Incrementar el conteo para TODOS los productos con este EAN
                foreach (string itemCode in itemCodes)
                {
                    if (conteoActual.ContainsKey(itemCode))
                    {
                        conteoActual[itemCode]++;
                    }

                    // Actualizar DataGridView
                    foreach (DataGridViewRow row in dgvInventario.Rows)
                    {
                        if (row.Cells["Codigo"].Value.ToString() == itemCode)
                        {
                            int stockSistema = Convert.ToInt32(row.Cells["StockSistema"].Value);
                            int stockContado = conteoActual[itemCode];
                            int diferencia = stockContado - stockSistema;

                            row.Cells["StockContado"].Value = stockContado;
                            row.Cells["Diferencia"].Value = diferencia;

                            // Colorear según diferencia
                            if (diferencia == 0)
                            {
                                row.DefaultCellStyle.BackColor = Color.FromArgb(220, 255, 220); // Verde claro
                            }
                            else if (diferencia > 0)
                            {
                                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 200); // Amarillo
                            }
                            else
                            {
                                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 240, 240); // Rojo claro
                            }

                            dgvInventario.FirstDisplayedScrollingRowIndex = row.Index;
                            row.Selected = true;

                            break;
                        }
                    }
                }

                SystemSounds.Beep.Play();
                ActualizarProgreso();
                ActualizarEstadisticas();
            }
            else
            {
                SystemSounds.Exclamation.Play();
                MessageBox.Show($"El código '{codigoEscaneado}' no pertenece a este almacén o clasificación.",
                    "Código no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            txtEscaneo.Focus();
        }

        private void DgvInventario_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvInventario.Columns["StockContado"].Index)
            {
                var row = dgvInventario.Rows[e.RowIndex];
                string itemCode = row.Cells["Codigo"].Value.ToString();

                if (int.TryParse(row.Cells["StockContado"].Value?.ToString(), out int nuevoValor))
                {
                    if (nuevoValor < 0)
                    {
                        MessageBox.Show("El stock contado no puede ser negativo.", "Valor inválido",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        row.Cells["StockContado"].Value = conteoActual[itemCode];
                        return;
                    }

                    // Actualizar el conteo actual
                    conteoActual[itemCode] = nuevoValor;

                    // Recalcular diferencia
                    int stockSistema = Convert.ToInt32(row.Cells["StockSistema"].Value);
                    int diferencia = nuevoValor - stockSistema;
                    row.Cells["Diferencia"].Value = diferencia;

                    // Actualizar color según diferencia
                    if (diferencia == 0)
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(220, 255, 220);
                    }
                    else if (diferencia > 0)
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 200);
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 240, 240);
                    }

                    ActualizarProgreso();
                    ActualizarEstadisticas();
                }
                else
                {
                    MessageBox.Show("Por favor, ingrese un número válido.", "Valor inválido",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    row.Cells["StockContado"].Value = conteoActual[itemCode];
                }
            }
        }

        private void ActualizarProgreso()
        {
            int totalProductos = productosParaInventariar.Count;
            int productosInventariados = conteoActual.Count(c => c.Value > 0);

            lblProgreso.Text = $"Productos inventariados: {productosInventariados} de {totalProductos}\n" +
                             $"Total escaneado: {conteoActual.Sum(c => c.Value)} unidades";
        }

        private void BtnFinalizar_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show(
                "¿Finalizar inventario y guardar respaldo?\n\nSe guardará un archivo Excel con todos los datos.",
                "Finalizar y Guardar",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (resultado == DialogResult.Yes)
            {
                // Detener autoguardado
                timerAutoguardado.Stop();

                GuardarRespaldoInventario();

                // Limpiar sesión guardada
                if (!string.IsNullOrEmpty(archivoSesion) && File.Exists(archivoSesion))
                {
                    try
                    {
                        File.Delete(archivoSesion);
                    }
                    catch { }
                }

                MessageBox.Show(
                    "✅ Inventario finalizado y guardado exitosamente.\n\nAhora puede generar el reporte por correo si lo desea.",
                    "Guardado Exitoso",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Habilitar botón de generar reporte y deshabilitar finalizar
                btnGenerarReporte.Enabled = true;
                btnFinalizar.Enabled = false;
                txtEscaneo.Enabled = false;
            }
        }

        private void BtnGenerarReporte_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show(
                "¿Generar reporte y abrir en Outlook?\n\nSe creará un correo con el archivo Excel adjunto.",
                "Generar Reporte",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (resultado == DialogResult.Yes)
            {
                GenerarReporteInventario();

                MessageBox.Show(
                    "✅ Reporte generado.\n\nSe ha abierto Outlook con el correo del reporte.",
                    "Reporte Generado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                this.Close();
            }
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show(
                "¿Está seguro de cancelar el inventario?\n\nSe perderán todos los datos escaneados.",
                "Cancelar Inventario",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void GenerarReporteInventario()
        {
            try
            {
                // Calcular estadísticas
                int productosSobrantes = 0;
                int productosFaltantes = 0;
                int productosCorrectos = 0;
                int totalArticulos = 0;

                foreach (DataGridViewRow row in dgvInventario.Rows)
                {
                    int diferencia = Convert.ToInt32(row.Cells["Diferencia"].Value);
                    totalArticulos++;

                    if (diferencia > 0) productosSobrantes++;
                    else if (diferencia < 0) productosFaltantes++;
                    else productosCorrectos++;
                }

                // Cuerpo del email simple y profesional
                string htmlBody = $@"
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #333; line-height: 1.6; }}
        .container {{ max-width: 550px; margin: 0 auto; padding: 20px; }}
        .header {{ background: #2980b9; color: white; padding: 20px; text-align: center; border-radius: 5px 5px 0 0; }}
        .header h2 {{ margin: 0; font-size: 20px; }}
        .content {{ background: #fff; padding: 25px; border: 1px solid #ddd; border-top: none; }}
        .info {{ margin: 20px 0; padding: 15px; background: #f8f9fa; border-left: 4px solid #2980b9; }}
        .stats {{ display: flex; justify-content: space-around; margin: 20px 0; }}
        .stat {{ text-align: center; }}
        .stat-num {{ font-size: 32px; font-weight: bold; }}
        .correct {{ color: #28a745; }}
        .excess {{ color: #ffc107; }}
        .shortage {{ color: #dc3545; }}
        .footer {{ text-align: center; padding: 15px; color: #999; font-size: 11px; }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <h2>📦 Inventario Selectivo</h2>
        </div>

        <div class='content'>
            <p><strong>Almacén:</strong> {almacenSeleccionado}<br>
            <strong>Clasificaciones:</strong> {string.Join(", ", clasificacionesSeleccionadas)}<br>
            <strong>Fecha:</strong> {DateTime.Now:dd/MM/yyyy HH:mm}</p>

            <div class='stats'>
                <div class='stat'>
                    <div class='stat-num correct'>{productosCorrectos}</div>
                    <div>Correctos</div>
                </div>
                <div class='stat'>
                    <div class='stat-num excess'>{productosSobrantes}</div>
                    <div>Sobrantes</div>
                </div>
                <div class='stat'>
                    <div class='stat-num shortage'>{productosFaltantes}</div>
                    <div>Faltantes</div>
                </div>
            </div>

            <div class='info'>
                📎 Ver detalle completo en el archivo Excel adjunto.
            </div>

            <p>Saludos,<br><strong>StockControl v1.2.2</strong></p>
        </div>

        <div class='footer'>
            Desarrollado por Fernando Carrasco
        </div>
    </div>
</body>
</html>";

                // Crear correo en Outlook
                dynamic outlook = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Outlook.Application"));
                dynamic mail = outlook.CreateItem(0); // 0 = MailItem

                mail.Subject = $"📦 Inventario Selectivo - {almacenSeleccionado} - {DateTime.Now:dd/MM/yyyy}";
                mail.HTMLBody = htmlBody;

                // Adjuntar el archivo Excel si existe
                if (!string.IsNullOrEmpty(archivoExcelRespaldo) && File.Exists(archivoExcelRespaldo))
                {
                    mail.Attachments.Add(archivoExcelRespaldo);
                }
                else
                {
                    MessageBox.Show("Advertencia: No se pudo encontrar el archivo Excel de respaldo.\n\nEl correo se generará sin adjunto.",
                        "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                mail.Display();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar el reporte en Outlook:\n\n{ex.Message}\n\nAsegúrese de que Outlook esté instalado correctamente.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GuardarRespaldoInventario()
        {
            try
            {
                // Crear carpeta de historial si no existe
                string carpetaHistorial = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Historial_Inventarios");
                if (!Directory.Exists(carpetaHistorial))
                {
                    Directory.CreateDirectory(carpetaHistorial);
                }

                // Generar nombre de archivo con fecha
                string fechaHoy = DateTime.Now.ToString("yyyy-MM-dd");
                string clasificacionesTexto = string.Join("-", clasificacionesSeleccionadas.Select(c => c.Replace(" ", "")));
                string nombreBase = $"Inventario_{almacenSeleccionado}_{clasificacionesTexto}_{fechaHoy}";
                string nombreArchivo = nombreBase + ".xlsx";
                string rutaCompleta = Path.Combine(carpetaHistorial, nombreArchivo);

                // Si ya existe, agregar sufijo numérico
                int contador = 2;
                while (File.Exists(rutaCompleta))
                {
                    nombreArchivo = $"{nombreBase}-{contador}.xlsx";
                    rutaCompleta = Path.Combine(carpetaHistorial, nombreArchivo);
                    contador++;
                }

                // Crear Excel con ClosedXML
                using (var workbook = new ClosedXML.Excel.XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Inventario");

                    // ENCABEZADO PRINCIPAL
                    worksheet.Cell("A1").Value = "REPORTE DE INVENTARIO SELECTIVO";
                    worksheet.Range("A1:H1").Merge();
                    worksheet.Cell("A1").Style.Font.Bold = true;
                    worksheet.Cell("A1").Style.Font.FontSize = 16;
                    worksheet.Cell("A1").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(41, 128, 185);
                    worksheet.Cell("A1").Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                    worksheet.Cell("A1").Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                    // INFORMACIÓN DEL INVENTARIO
                    int row = 3;
                    worksheet.Cell($"A{row}").Value = "Fecha:";
                    worksheet.Cell($"B{row}").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                    worksheet.Cell($"A{row}").Style.Font.Bold = true;

                    row++;
                    worksheet.Cell($"A{row}").Value = "Almacén:";
                    worksheet.Cell($"B{row}").Value = almacenSeleccionado;
                    worksheet.Cell($"A{row}").Style.Font.Bold = true;

                    row++;
                    worksheet.Cell($"A{row}").Value = "Clasificaciones:";
                    worksheet.Cell($"B{row}").Value = string.Join(", ", clasificacionesSeleccionadas);
                    worksheet.Cell($"A{row}").Style.Font.Bold = true;

                    // RESUMEN
                    row += 2;
                    int productosCorrectos = dgvInventario.Rows.Cast<DataGridViewRow>()
                        .Count(r => Convert.ToInt32(r.Cells["Diferencia"].Value) == 0);
                    int productosSobrantes = dgvInventario.Rows.Cast<DataGridViewRow>()
                        .Count(r => Convert.ToInt32(r.Cells["Diferencia"].Value) > 0);
                    int productosFaltantes = dgvInventario.Rows.Cast<DataGridViewRow>()
                        .Count(r => Convert.ToInt32(r.Cells["Diferencia"].Value) < 0);

                    worksheet.Cell($"A{row}").Value = "RESUMEN";
                    worksheet.Range($"A{row}:D{row}").Merge();
                    worksheet.Cell($"A{row}").Style.Font.Bold = true;
                    worksheet.Cell($"A{row}").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(52, 73, 94);
                    worksheet.Cell($"A{row}").Style.Font.FontColor = ClosedXML.Excel.XLColor.White;

                    row++;
                    worksheet.Cell($"A{row}").Value = "Productos Correctos:";
                    worksheet.Cell($"B{row}").Value = productosCorrectos;
                    worksheet.Cell($"B{row}").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(212, 237, 218);

                    worksheet.Cell($"C{row}").Value = "Productos Sobrantes:";
                    worksheet.Cell($"D{row}").Value = productosSobrantes;
                    worksheet.Cell($"D{row}").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(255, 243, 205);

                    row++;
                    worksheet.Cell($"A{row}").Value = "Productos Faltantes:";
                    worksheet.Cell($"B{row}").Value = productosFaltantes;
                    worksheet.Cell($"B{row}").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(248, 215, 218);

                    // TABLA DE DATOS
                    row += 2;
                    int headerRow = row;

                    // Encabezados
                    worksheet.Cell($"A{row}").Value = "Marca";
                    worksheet.Cell($"B{row}").Value = "Clasificación";
                    worksheet.Cell($"C{row}").Value = "Detalle";
                    worksheet.Cell($"D{row}").Value = "Código";
                    worksheet.Cell($"E{row}").Value = "EAN";
                    worksheet.Cell($"F{row}").Value = "Stock Sistema";
                    worksheet.Cell($"G{row}").Value = "Stock Contado";
                    worksheet.Cell($"H{row}").Value = "Diferencia";

                    var headerRange = worksheet.Range($"A{row}:H{row}");
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(52, 73, 94);
                    headerRange.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                    headerRange.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                    // Datos
                    foreach (DataGridViewRow dgvRow in dgvInventario.Rows)
                    {
                        row++;
                        worksheet.Cell($"A{row}").Value = dgvRow.Cells["Marca"].Value?.ToString();
                        worksheet.Cell($"B{row}").Value = dgvRow.Cells["Clasificacion"].Value?.ToString();
                        worksheet.Cell($"C{row}").Value = dgvRow.Cells["Detalle"].Value?.ToString();
                        worksheet.Cell($"D{row}").Value = dgvRow.Cells["Codigo"].Value?.ToString();
                        worksheet.Cell($"E{row}").Value = dgvRow.Cells["EAN"].Value?.ToString();
                        worksheet.Cell($"F{row}").Value = Convert.ToInt32(dgvRow.Cells["StockSistema"].Value);
                        worksheet.Cell($"G{row}").Value = Convert.ToInt32(dgvRow.Cells["StockContado"].Value);
                        worksheet.Cell($"H{row}").Value = Convert.ToInt32(dgvRow.Cells["Diferencia"].Value);

                        int diferencia = Convert.ToInt32(dgvRow.Cells["Diferencia"].Value);
                        if (diferencia == 0)
                        {
                            worksheet.Cell($"H{row}").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(212, 237, 218);
                            worksheet.Cell($"H{row}").Style.Font.FontColor = ClosedXML.Excel.XLColor.FromArgb(21, 87, 36);
                        }
                        else if (diferencia > 0)
                        {
                            worksheet.Cell($"H{row}").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(255, 243, 205);
                            worksheet.Cell($"H{row}").Style.Font.FontColor = ClosedXML.Excel.XLColor.FromArgb(133, 100, 4);
                        }
                        else
                        {
                            worksheet.Cell($"H{row}").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(248, 215, 218);
                            worksheet.Cell($"H{row}").Style.Font.FontColor = ClosedXML.Excel.XLColor.FromArgb(114, 28, 36);
                        }

                        worksheet.Cell($"H{row}").Style.Font.Bold = true;
                        worksheet.Cell($"H{row}").Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                    }

                    // Ajustar columnas
                    worksheet.Columns().AdjustToContents();

                    // Agregar bordes a la tabla
                    var dataRange = worksheet.Range($"A{headerRow}:H{row}");
                    dataRange.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Medium;
                    dataRange.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;

                    workbook.SaveAs(rutaCompleta);

                    // Guardar la ruta para adjuntar en el email
                    archivoExcelRespaldo = rutaCompleta;
                }
            }
            catch (Exception ex)
            {
                // No interrumpir el flujo si falla el respaldo
                System.Diagnostics.Debug.WriteLine($"Error al guardar respaldo: {ex.Message}");
            }
        }

        // ============================================
        // NUEVAS FUNCIONALIDADES v1.1.0
        // ============================================

        #region Persistencia de Sesión y Autoguardado

        private void ConfigurarAutoguardado()
        {
            timerAutoguardado = new System.Windows.Forms.Timer();
            timerAutoguardado.Interval = 120000; // 2 minutos
            timerAutoguardado.Tick += TimerAutoguardado_Tick;
        }

        private void TimerAutoguardado_Tick(object sender, EventArgs e)
        {
            if (productosParaInventariar.Count > 0 && panelEscaneo.Visible)
            {
                GuardarSesion();
            }
        }

        private void GuardarSesion()
        {
            try
            {
                string carpetaSesiones = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sesiones");
                if (!Directory.Exists(carpetaSesiones))
                {
                    Directory.CreateDirectory(carpetaSesiones);
                }

                if (string.IsNullOrEmpty(archivoSesion))
                {
                    archivoSesion = Path.Combine(carpetaSesiones, $"Sesion_{DateTime.Now:yyyyMMdd_HHmmss}.json");
                }

                var sesion = new
                {
                    FechaGuardado = DateTime.Now,
                    Almacen = almacenSeleccionado,
                    Clasificaciones = clasificacionesSeleccionadas,
                    UsarSubtipo = usarSubtipo,
                    ConteosActuales = conteoActual
                };

                string json = JsonSerializer.Serialize(sesion, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(archivoSesion, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al guardar sesión: {ex.Message}");
            }
        }

        private void VerificarSesionPendiente()
        {
            try
            {
                string carpetaSesiones = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sesiones");
                if (!Directory.Exists(carpetaSesiones)) return;

                var archivos = Directory.GetFiles(carpetaSesiones, "*.json");
                if (archivos.Length == 0) return;

                var resultado = MessageBox.Show(
                    "Se detectó una sesión anterior no finalizada.\n\n¿Desea recuperar el inventario en progreso?",
                    "Sesión Pendiente",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (resultado == DialogResult.Yes)
                {
                    string archivoMasReciente = archivos.OrderByDescending(f => File.GetLastWriteTime(f)).First();
                    CargarSesion(archivoMasReciente);
                }
                else
                {
                    // Limpiar sesiones antiguas
                    foreach (var archivo in archivos)
                    {
                        File.Delete(archivo);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al verificar sesiones: {ex.Message}");
            }
        }

        private void CargarSesion(string archivo)
        {
            try
            {
                string json = File.ReadAllText(archivo);
                var sesion = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);

                if (sesion != null)
                {
                    archivoSesion = archivo;

                    // Restaurar datos básicos
                    almacenSeleccionado = sesion["Almacen"].GetString() ?? "";
                    usarSubtipo = sesion["UsarSubtipo"].GetBoolean();

                    // Restaurar clasificaciones
                    clasificacionesSeleccionadas.Clear();
                    foreach (var item in sesion["Clasificaciones"].EnumerateArray())
                    {
                        clasificacionesSeleccionadas.Add(item.GetString() ?? "");
                    }

                    // Restaurar conteos
                    conteoActual.Clear();
                    foreach (var prop in sesion["ConteosActuales"].EnumerateObject())
                    {
                        conteoActual[prop.Name] = prop.Value.GetInt32();
                    }

                    // Cargar productos y aplicar conteos
                    var productos = ExcelDataManager.ProductosExcel
                        .Where(p => p.WhsCode == almacenSeleccionado &&
                                  (usarSubtipo ? clasificacionesSeleccionadas.Contains(p.U_Comercial3)
                                               : clasificacionesSeleccionadas.Contains(p.ItmsGrpNam)))
                        .ToList();

                    productosParaInventariar.Clear();
                    eanToItemCodes.Clear();
                    dgvInventario.Rows.Clear();

                    foreach (var producto in productos)
                    {
                        string key = producto.ItemCode;

                        if (!productosParaInventariar.ContainsKey(key))
                        {
                            productosParaInventariar[key] = producto;

                            // Mantener lookup de EAN -> ItemCode
                            if (!eanToItemCodes.ContainsKey(producto.CodeBars))
                            {
                                eanToItemCodes[producto.CodeBars] = new List<string>();
                            }
                            eanToItemCodes[producto.CodeBars].Add(key);

                            int stockContado = conteoActual.ContainsKey(key) ? conteoActual[key] : 0;

                            int rowIndex = dgvInventario.Rows.Add();
                            DataGridViewRow row = dgvInventario.Rows[rowIndex];
                            row.Cells["Marca"].Value = producto.ItmsGrpNam;
                            row.Cells["Clasificacion"].Value = producto.U_Comercial1;
                            row.Cells["Detalle"].Value = producto.U_Comercial3;
                            row.Cells["Codigo"].Value = producto.ItemCode;
                            row.Cells["EAN"].Value = producto.CodeBars;
                            row.Cells["StockSistema"].Value = producto.StockTienda;
                            row.Cells["StockContado"].Value = stockContado;
                            row.Cells["Diferencia"].Value = stockContado - producto.StockTienda;

                            int diferencia = stockContado - producto.StockTienda;
                            if (diferencia == 0)
                                row.DefaultCellStyle.BackColor = Color.FromArgb(220, 255, 220);
                            else if (diferencia > 0)
                                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 200);
                            else
                                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 240, 240);
                        }
                    }

                    // Cambiar a panel de escaneo
                    panelSeleccion.Visible = false;
                    panelEscaneo.Visible = true;
                    panelEstadisticas.Visible = true;
                    timerAutoguardado.Start();

                    ActualizarProgreso();
                    ActualizarEstadisticas();

                    MessageBox.Show("✅ Sesión recuperada exitosamente.", "Sesión Restaurada",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar sesión: {ex.Message}", "Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Búsqueda Rápida

        private void TxtBusqueda_TextChanged(object sender, EventArgs e)
        {
            string busqueda = txtBusqueda.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(busqueda))
            {
                // Mostrar todas las filas (respetando filtro de diferencias si está activo)
                AplicarFiltros();
                return;
            }

            foreach (DataGridViewRow row in dgvInventario.Rows)
            {
                string codigo = row.Cells["Codigo"].Value?.ToString()?.ToLower() ?? "";
                string ean = row.Cells["EAN"].Value?.ToString()?.ToLower() ?? "";
                string detalle = row.Cells["Detalle"].Value?.ToString()?.ToLower() ?? "";

                bool coincide = codigo.Contains(busqueda) || ean.Contains(busqueda) || detalle.Contains(busqueda);

                // Si está activo el filtro de diferencias, también verificar eso
                if (chkSoloDiferencias.Checked)
                {
                    int diferencia = Convert.ToInt32(row.Cells["Diferencia"].Value);
                    row.Visible = coincide && diferencia != 0;
                }
                else
                {
                    row.Visible = coincide;
                }
            }
        }

        #endregion

        #region Filtro Solo Diferencias

        private void ChkSoloDiferencias_CheckedChanged(object sender, EventArgs e)
        {
            AplicarFiltros();
        }

        private void AplicarFiltros()
        {
            string busqueda = txtBusqueda.Text.Trim().ToLower();
            bool soloDiferencias = chkSoloDiferencias.Checked;

            foreach (DataGridViewRow row in dgvInventario.Rows)
            {
                bool cumpleFiltros = true;

                // Filtro de búsqueda
                if (!string.IsNullOrEmpty(busqueda))
                {
                    string codigo = row.Cells["Codigo"].Value?.ToString()?.ToLower() ?? "";
                    string ean = row.Cells["EAN"].Value?.ToString()?.ToLower() ?? "";
                    string detalle = row.Cells["Detalle"].Value?.ToString()?.ToLower() ?? "";
                    cumpleFiltros = codigo.Contains(busqueda) || ean.Contains(busqueda) || detalle.Contains(busqueda);
                }

                // Filtro de diferencias
                if (cumpleFiltros && soloDiferencias)
                {
                    int diferencia = Convert.ToInt32(row.Cells["Diferencia"].Value);
                    cumpleFiltros = diferencia != 0;
                }

                row.Visible = cumpleFiltros;
            }
        }

        #endregion

        #region Estadísticas en Tiempo Real

        private void ActualizarEstadisticas()
        {
            int total = dgvInventario.Rows.Count;
            int correctos = 0;
            int sobrantes = 0;
            int faltantes = 0;
            int inventariados = 0;

            foreach (DataGridViewRow row in dgvInventario.Rows)
            {
                int stockContado = Convert.ToInt32(row.Cells["StockContado"].Value);
                int diferencia = Convert.ToInt32(row.Cells["Diferencia"].Value);

                if (stockContado > 0) inventariados++;

                if (diferencia == 0) correctos++;
                else if (diferencia > 0) sobrantes++;
                else faltantes++;
            }

            int porcentaje = total > 0 ? (inventariados * 100 / total) : 0;

            lblTotalProductos.Text = $"Total de Productos\n{total}";
            lblProductosCorrectos.Text = $"✅ Correctos\n{correctos}";
            lblProductosSobrantes.Text = $"⬆️ Sobrantes\n{sobrantes}";
            lblProductosFaltantes.Text = $"⬇️ Faltantes\n{faltantes}";
            lblPorcentajeAvance.Text = $"AVANCE\n{porcentaje}%";
        }

        #endregion

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ClientSize = new Size(1000, 700);
            this.Name = "InventarioForm";
            this.ResumeLayout(false);
        }
    }

}

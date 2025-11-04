using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Windows.Forms;

namespace Inventario
{
    public partial class InventarioForm : Form
    {
        private Dictionary<string, ProductoExcel> productosParaInventariar;
        private Dictionary<string, int> conteoActual;
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

        public InventarioForm()
        {
            InitializeComponent();
            InicializarComponentes();
            productosParaInventariar = new Dictionary<string, ProductoExcel>();
            conteoActual = new Dictionary<string, int>();
            clasificacionesSeleccionadas = new List<string>();
            ConfigurarIcono();
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

            // Botones finales
            btnFinalizar = new Button
            {
                Text = "💾 FINALIZAR Y GUARDAR",
                Location = new Point(350, 560),
                Size = new Size(220, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 150, 0),
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
                Location = new Point(580, 560),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(41, 128, 185),
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
                Location = new Point(790, 560),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(200, 50, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnCancelar.FlatAppearance.BorderSize = 0;
            btnCancelar.Click += BtnCancelar_Click;
            panelEscaneo.Controls.Add(btnCancelar);

            this.Controls.Add(panelEscaneo);

            // Footer
            Panel panelFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 35,
                BackColor = Color.FromArgb(52, 73, 94)
            };

            Label lblFooter = new Label
            {
                Text = "StockControl v1.0.1 | Desarrollado por Fernando Carrasco",
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
                if (!productosParaInventariar.ContainsKey(producto.CodeBars))
                {
                    productosParaInventariar[producto.CodeBars] = producto;
                    conteoActual[producto.CodeBars] = 0;

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

            // Cambiar a modo escaneo
            panelSeleccion.Visible = false;
            panelEscaneo.Visible = true;
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

            if (productosParaInventariar.ContainsKey(codigoEscaneado))
            {
                conteoActual[codigoEscaneado]++;

                // Actualizar DataGridView
                foreach (DataGridViewRow row in dgvInventario.Rows)
                {
                    if (row.Cells["EAN"].Value.ToString() == codigoEscaneado)
                    {
                        int stockSistema = Convert.ToInt32(row.Cells["StockSistema"].Value);
                        int stockContado = conteoActual[codigoEscaneado];
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

                        SystemSounds.Beep.Play();
                        break;
                    }
                }

                ActualizarProgreso();
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
                string ean = row.Cells["EAN"].Value.ToString();

                if (int.TryParse(row.Cells["StockContado"].Value?.ToString(), out int nuevoValor))
                {
                    if (nuevoValor < 0)
                    {
                        MessageBox.Show("El stock contado no puede ser negativo.", "Valor inválido",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        row.Cells["StockContado"].Value = conteoActual[ean];
                        return;
                    }

                    // Actualizar el conteo actual
                    conteoActual[ean] = nuevoValor;

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
                }
                else
                {
                    MessageBox.Show("Por favor, ingrese un número válido.", "Valor inválido",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    row.Cells["StockContado"].Value = conteoActual[ean];
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
                GuardarRespaldoInventario();

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
                "¿Generar reporte y abrir en Outlook?\n\nSe creará un correo con el reporte HTML.",
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
                int productosSobrantes = 0;
                int productosFaltantes = 0;
                int productosCorrectos = 0;
                int totalArticulos = 0;

                var productosDiferentes = new System.Collections.Generic.List<string>();

                foreach (DataGridViewRow row in dgvInventario.Rows)
                {
                    int diferencia = Convert.ToInt32(row.Cells["Diferencia"].Value);
                    totalArticulos++;

                    if (diferencia > 0) productosSobrantes++;
                    else if (diferencia < 0) productosFaltantes++;
                    else productosCorrectos++;

                    if (diferencia != 0)
                    {
                        string colorFondo = diferencia > 0 ? "#fff3cd" : "#f8d7da";
                        string colorTexto = diferencia > 0 ? "#856404" : "#721c24";

                        productosDiferentes.Add($@"
                            <tr style='background-color: {colorFondo};'>
                                <td style='padding: 12px; border: 1px solid #dee2e6;'>{row.Cells["Codigo"].Value}</td>
                                <td style='padding: 12px; border: 1px solid #dee2e6;'>{row.Cells["EAN"].Value}</td>
                                <td style='padding: 12px; border: 1px solid #dee2e6;'>{row.Cells["Detalle"].Value}</td>
                                <td style='padding: 12px; border: 1px solid #dee2e6; text-align: center;'>{row.Cells["StockSistema"].Value}</td>
                                <td style='padding: 12px; border: 1px solid #dee2e6; text-align: center;'>{row.Cells["StockContado"].Value}</td>
                                <td style='padding: 12px; border: 1px solid #dee2e6; text-align: center; font-weight: bold; color: {colorTexto};'>{diferencia:+#;-#;0}</td>
                            </tr>");
                    }
                }

                string htmlBody = $@"
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; }}
        .header {{ background-color: #2980b9; color: white; padding: 20px; text-align: center; }}
        .content {{ padding: 20px; }}
        .summary {{ background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin: 20px 0; }}
        .summary-item {{ display: inline-block; margin: 10px 20px; }}
        .summary-label {{ font-weight: bold; color: #2c3e50; }}
        .summary-value {{ font-size: 24px; font-weight: bold; color: #2980b9; }}
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th {{ background-color: #34495e; color: white; padding: 12px; text-align: left; border: 1px solid #2c3e50; }}
        td {{ padding: 12px; border: 1px solid #dee2e6; }}
        .footer {{ text-align: center; padding: 20px; color: #7f8c8d; font-size: 12px; }}
        .badge {{ display: inline-block; padding: 5px 10px; border-radius: 3px; font-weight: bold; }}
        .badge-success {{ background-color: #d4edda; color: #155724; }}
        .badge-warning {{ background-color: #fff3cd; color: #856404; }}
        .badge-danger {{ background-color: #f8d7da; color: #721c24; }}
    </style>
</head>
<body>
    <div class='header'>
        <h1>REPORTE DE INVENTARIO SELECTIVO</h1>
        <p>Fecha: {DateTime.Now:dd/MM/yyyy HH:mm}</p>
    </div>

    <div class='content'>
        <h2>Información del Inventario</h2>
        <p><strong>Almacén:</strong> {almacenSeleccionado}</p>
        <p><strong>Clasificaciones:</strong> {string.Join(", ", clasificacionesSeleccionadas)}</p>
        <p><strong>Total de Artículos:</strong> {totalArticulos}</p>

        <div class='summary'>
            <h3 style='margin-top: 0; color: #2c3e50;'>Resumen de Resultados</h3>
            <div class='summary-item'>
                <div class='summary-label'>Productos Correctos</div>
                <div class='summary-value' style='color: #27ae60;'>{productosCorrectos}</div>
                <span class='badge badge-success'>✓ Sin diferencias</span>
            </div>
            <div class='summary-item'>
                <div class='summary-label'>Productos con Sobrante</div>
                <div class='summary-value' style='color: #f39c12;'>{productosSobrantes}</div>
                <span class='badge badge-warning'>↑ Excedente</span>
            </div>
            <div class='summary-item'>
                <div class='summary-label'>Productos con Faltante</div>
                <div class='summary-value' style='color: #e74c3c;'>{productosFaltantes}</div>
                <span class='badge badge-danger'>↓ Faltante</span>
            </div>
        </div>

        <h3>Detalle de Diferencias</h3>
        {(productosDiferentes.Count > 0 ?
            $@"<table>
                <thead>
                    <tr>
                        <th>Código</th>
                        <th>EAN</th>
                        <th>Descripción</th>
                        <th style='text-align: center;'>Stock Sistema</th>
                        <th style='text-align: center;'>Stock Contado</th>
                        <th style='text-align: center;'>Diferencia</th>
                    </tr>
                </thead>
                <tbody>
                    {string.Join("", productosDiferentes)}
                </tbody>
            </table>" :
            "<p style='text-align: center; padding: 20px; background-color: #d4edda; color: #155724; border-radius: 5px;'><strong>¡Excelente!</strong> No se encontraron diferencias en el inventario.</p>"
        )}
    </div>

    <div class='footer'>
        <p>StockControl v1.0.1 | Desarrollado por Fernando Carrasco</p>
        <p>Este reporte fue generado automáticamente el {DateTime.Now:dd/MM/yyyy} a las {DateTime.Now:HH:mm}</p>
    </div>
</body>
</html>";

                // Crear correo en Outlook
                dynamic outlook = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Outlook.Application"));
                dynamic mail = outlook.CreateItem(0); // 0 = MailItem

                mail.Subject = $"Inventario Selectivo {DateTime.Now:dd/MM/yyyy} - Almacén {almacenSeleccionado}";
                mail.HTMLBody = htmlBody;
                mail.Display();

                MessageBox.Show("Se ha generado el reporte en Outlook.\n\nPor favor, revise el correo y envíelo a los destinatarios correspondientes.",
                    "Reporte Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                }
            }
            catch (Exception ex)
            {
                // No interrumpir el flujo si falla el respaldo
                System.Diagnostics.Debug.WriteLine($"Error al guardar respaldo: {ex.Message}");
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ClientSize = new Size(1000, 700);
            this.Name = "InventarioForm";
            this.ResumeLayout(false);
        }
    }

}

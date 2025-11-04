using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Inventario
{
    public partial class InventarioForm : Form
    {
        private Dictionary<string, ProductoExcel> productosParaInventariar;
        private Dictionary<string, int> conteoActual;
        private string almacenSeleccionado;
        private string clasificacionSeleccionada;

        private ComboBox cboAlmacen;
        private ComboBox cboClasificacion;
        private TextBox txtEscaneo;
        private Label lblInstruccion;
        private Label lblProgreso;
        private DataGridView dgvInventario;
        private Button btnIniciar;
        private Button btnFinalizar;
        private Button btnCancelar;
        private Panel panelSeleccion;
        private Panel panelEscaneo;

        public InventarioForm()
        {
            InitializeComponent();
            InicializarComponentes();
            productosParaInventariar = new Dictionary<string, ProductoExcel>();
            conteoActual = new Dictionary<string, int>();
        }

        private void InicializarComponentes()
        {
            this.Text = "Inventario de Almacén";
            this.Size = new Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);

            // Panel de selección (visible al inicio)
            panelSeleccion = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(940, 200),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            Label lblTitulo = new Label
            {
                Text = "Configuración del Inventario",
                Location = new Point(20, 20),
                Size = new Size(400, 30),
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

            // Selección de clasificación
            Label lblClasificacion = new Label
            {
                Text = "Seleccione la Clasificación:",
                Location = new Point(30, 120),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };
            panelSeleccion.Controls.Add(lblClasificacion);

            cboClasificacion = new ComboBox
            {
                Location = new Point(240, 118),
                Size = new Size(300, 30),
                Font = new Font("Segoe UI", 11),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false
            };
            panelSeleccion.Controls.Add(cboClasificacion);

            // Botón Iniciar
            btnIniciar = new Button
            {
                Text = "INICIAR INVENTARIO",
                Location = new Point(240, 160),
                Size = new Size(300, 35),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false,
                Cursor = Cursors.Hand
            };
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
                ReadOnly = true,
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

            dgvInventario.Columns["StockSistema"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvInventario.Columns["StockContado"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvInventario.Columns["Diferencia"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            panelEscaneo.Controls.Add(dgvInventario);

            // Botones de finalizar y cancelar
            btnFinalizar = new Button
            {
                Text = "FINALIZAR INVENTARIO",
                Location = new Point(550, 560),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 150, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnFinalizar.Click += BtnFinalizar_Click;
            panelEscaneo.Controls.Add(btnFinalizar);

            btnCancelar = new Button
            {
                Text = "CANCELAR",
                Location = new Point(760, 560),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(200, 50, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnCancelar.Click += BtnCancelar_Click;
            panelEscaneo.Controls.Add(btnCancelar);

            this.Controls.Add(panelEscaneo);

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

            var clasificaciones = ExcelDataManager.ProductosExcel
                .Where(p => p.WhsCode == almacenSeleccionado)
                .Select(p => p.ItmsGrpNam)
                .Distinct()
                .OrderBy(c => c)
                .ToList();

            cboClasificacion.Items.Clear();
            foreach (var clasificacion in clasificaciones)
            {
                cboClasificacion.Items.Add(clasificacion);
            }

            cboClasificacion.Enabled = true;
            cboClasificacion.SelectedIndex = -1;
            btnIniciar.Enabled = false;
        }

        private void CboClasificacion_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnIniciar.Enabled = cboClasificacion.SelectedItem != null;
        }

        private void BtnIniciar_Click(object sender, EventArgs e)
        {
            clasificacionSeleccionada = cboClasificacion.SelectedItem.ToString();

            // Cargar productos a inventariar
            productosParaInventariar.Clear();
            conteoActual.Clear();

            var productos = ExcelDataManager.ProductosExcel
                .Where(p => p.WhsCode == almacenSeleccionado && p.ItmsGrpNam == clasificacionSeleccionada)
                .ToList();

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
                "¿Está seguro de finalizar el inventario?\n\nSe generará un reporte con las diferencias encontradas.",
                "Finalizar Inventario",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (resultado == DialogResult.Yes)
            {
                GenerarReporteInventario();
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
                var reporte = new System.Text.StringBuilder();
                reporte.AppendLine("REPORTE DE INVENTARIO");
                reporte.AppendLine($"Fecha: {DateTime.Now:dd/MM/yyyy HH:mm}");
                reporte.AppendLine($"Almacén: {almacenSeleccionado}");
                reporte.AppendLine($"Clasificación: {clasificacionSeleccionada}");
                reporte.AppendLine(new string('=', 100));
                reporte.AppendLine();

                int productosSobrantes = 0;
                int productosFaltantes = 0;
                int productosCorrectos = 0;

                foreach (DataGridViewRow row in dgvInventario.Rows)
                {
                    int diferencia = Convert.ToInt32(row.Cells["Diferencia"].Value);

                    if (diferencia > 0) productosSobrantes++;
                    else if (diferencia < 0) productosFaltantes++;
                    else productosCorrectos++;

                    if (diferencia != 0)
                    {
                        reporte.AppendLine($"Código: {row.Cells["Codigo"].Value}");
                        reporte.AppendLine($"EAN: {row.Cells["EAN"].Value}");
                        reporte.AppendLine($"Detalle: {row.Cells["Detalle"].Value}");
                        reporte.AppendLine($"Stock Sistema: {row.Cells["StockSistema"].Value}");
                        reporte.AppendLine($"Stock Contado: {row.Cells["StockContado"].Value}");
                        reporte.AppendLine($"Diferencia: {diferencia}");
                        reporte.AppendLine(new string('-', 100));
                    }
                }

                reporte.AppendLine();
                reporte.AppendLine("RESUMEN:");
                reporte.AppendLine($"Productos correctos: {productosCorrectos}");
                reporte.AppendLine($"Productos con sobrante: {productosSobrantes}");
                reporte.AppendLine($"Productos con faltante: {productosFaltantes}");

                string rutaReporte = $"Inventario_{almacenSeleccionado}_{clasificacionSeleccionada}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                System.IO.File.WriteAllText(rutaReporte, reporte.ToString());

                MessageBox.Show($"Reporte generado exitosamente:\n{rutaReporte}",
                    "Inventario Finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar el reporte: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

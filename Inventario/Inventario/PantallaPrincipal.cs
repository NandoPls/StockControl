namespace Inventario
{
    public partial class PantallaPrincipal : Form
    {
        private string? archivoExcelSeleccionado;

        public PantallaPrincipal()
        {
            InitializeComponent();
            ConfigurarEstiloInicial();
            ConfigurarIcono();
        }

        private void ConfigurarEstiloInicial()
        {
            // Configuración inicial del formulario
            this.BackColor = Color.FromArgb(240, 244, 248);
        }

        private void ConfigurarIcono()
        {
            try
            {
                // Crear icono en memoria
                int iconSize = 64;
                using (Bitmap bitmap = new Bitmap(iconSize, iconSize))
                {
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                        // Fondo azul
                        Color blueColor = Color.FromArgb(41, 128, 185);
                        using (SolidBrush brush = new SolidBrush(blueColor))
                        {
                            graphics.FillEllipse(brush, 0, 0, iconSize, iconSize);
                        }

                        // Dibujar caja blanca
                        using (Pen whitePen = new Pen(Color.White, 4))
                        {
                            graphics.DrawRectangle(whitePen, 10, 12, 44, 40);
                        }

                        // Líneas internas
                        using (Pen whitePen2 = new Pen(Color.White, 2))
                        {
                            graphics.DrawLine(whitePen2, 10, 25, 54, 25);
                            graphics.DrawLine(whitePen2, 10, 38, 54, 38);
                        }

                        // Checkmark verde
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
                // Si hay error al crear el icono, continuar sin él
            }
        }

        private void btnCargarArchivo_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Seleccionar archivo de Excel";
                openFileDialog.Filter = "Archivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Todos los archivos (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    archivoExcelSeleccionado = openFileDialog.FileName;
                    lblArchivoSeleccionado.Text = $"Archivo: {Path.GetFileName(archivoExcelSeleccionado)}";
                    lblArchivoSeleccionado.Visible = true;
                    btnContinuar.Enabled = true;
                    btnContinuar.Visible = true;
                }
            }
        }

        private void btnContinuar_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(archivoExcelSeleccionado))
            {
                try
                {
                    // Cargar datos del Excel
                    if (ExcelDataManager.CargarExcel(archivoExcelSeleccionado))
                    {
                        MessageBox.Show($"Archivo cargado exitosamente.\n\nProductos cargados: {ExcelDataManager.ProductosExcel.Count}",
                            "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Abrir ventana de inventario
                        InventarioForm formInventario = new InventarioForm();
                        formInventario.ShowDialog();
                    }
                    else
                    {
                        MessageBox.Show("No se encontraron datos en el archivo Excel.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (IOException ioEx)
                {
                    MessageBox.Show(ioEx.Message, "Archivo en Uso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    string nombreArchivo = Path.GetFileName(archivoExcelSeleccionado);
                    MessageBox.Show($"Error al cargar el archivo \"{nombreArchivo}\":\n\n{ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

    }
}

using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Inventario
{
    public partial class PantallaPrincipal : Form
    {
        private string? archivoExcelSeleccionado;
        private bool usarSap = false;

        public PantallaPrincipal()
        {
            InitializeComponent();
            ConfigurarEstiloInicial();
            ConfigurarIcono();
            VerificarConfiguracionSap();
            this.Load += PantallaPrincipal_Load;
        }

        private async void PantallaPrincipal_Load(object? sender, EventArgs e)
        {
            // Verificar actualizaciones en segundo plano
            await AutoUpdater.CheckForUpdates(this);
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
            if (usarSap)
            {
                // Cargar desde SAP
                CargarDesdeSap();
            }
            else if (!string.IsNullOrEmpty(archivoExcelSeleccionado))
            {
                // Cargar desde Excel
                CargarDesdeExcel();
            }
        }

        private void CargarDesdeExcel()
        {
            try
            {
                // Cargar datos del Excel
                if (ExcelDataManager.CargarExcel(archivoExcelSeleccionado!))
                {
                    MessageBox.Show($"✅ Archivo Excel cargado exitosamente.\n\nProductos cargados: {ExcelDataManager.ProductosExcel.Count}",
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
                string nombreArchivo = Path.GetFileName(archivoExcelSeleccionado!);
                MessageBox.Show($"Error al cargar el archivo \"{nombreArchivo}\":\n\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CargarDesdeSap()
        {
            try
            {
                // Intentar conectar a SAP
                var config = ConfigManager.CargarConfiguracion();

                if (!SapConnector.ConfigurarConexion(config.SapConnection))
                {
                    MessageBox.Show("No se pudo conectar a SAP Business One.\n\nVerifique la configuración en appsettings.json",
                        "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Cargar productos desde SAP (por ahora sin filtro de almacén, se hace en InventarioForm)
                MessageBox.Show("✅ Conectado a SAP Business One exitosamente.\n\nLos datos se cargarán al seleccionar almacén y clasificación.",
                    "Conexión Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Marcar que estamos usando SAP
                ExcelDataManager.ConfigurarOrigenDatos("SAP");

                // Abrir ventana de inventario
                InventarioForm formInventario = new InventarioForm();
                formInventario.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al conectar con SAP:\n\n{ex.Message}\n\nVerifique appsettings.json y que el servidor SQL de SAP esté accesible.",
                    "Error SAP", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void VerificarConfiguracionSap()
        {
            // Verificar si SAP está habilitado en configuración
            var config = ConfigManager.CargarConfiguracion();

            if (config.SapConnection.Enabled)
            {
                // Mostrar opción SAP en la interfaz
                // Esto se puede expandir más adelante con un RadioButton o botón adicional
                usarSap = true;
            }
        }

        private void btnConectarSap_Click(object sender, EventArgs e)
        {
            usarSap = true;
            CargarDesdeSap();
        }

    }
}

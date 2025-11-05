using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Inventario
{
    public static class AutoUpdater
    {
        private const string VERSION_URL = "https://raw.githubusercontent.com/NandoPls/StockControl/master/version.txt";
        private const string RELEASE_URL = "https://github.com/NandoPls/StockControl/releases/download/v{0}/StockControl_v{0}.zip";
        private const string CURRENT_VERSION = "1.2.4";

        public static async Task<bool> CheckForUpdates(Form parentForm)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(10);

                    // Descargar versión más reciente desde GitHub
                    string latestVersion = await client.GetStringAsync(VERSION_URL);
                    latestVersion = latestVersion.Trim();

                    // Comparar versiones
                    if (IsNewerVersion(latestVersion, CURRENT_VERSION))
                    {
                        var result = MessageBox.Show(
                            $"¡Nueva versión disponible!\n\n" +
                            $"Versión actual: v{CURRENT_VERSION}\n" +
                            $"Nueva versión: v{latestVersion}\n\n" +
                            $"¿Desea descargar e instalar la actualización?\n" +
                            $"El programa se reiniciará automáticamente.",
                            "Actualización Disponible",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information);

                        if (result == DialogResult.Yes)
                        {
                            return await DownloadAndInstallUpdate(latestVersion, parentForm);
                        }
                    }
                }
            }
            catch (HttpRequestException)
            {
                // Sin conexión a internet o GitHub no disponible - ignorar silenciosamente
            }
            catch (TaskCanceledException)
            {
                // Timeout - ignorar silenciosamente
            }
            catch (Exception ex)
            {
                // Error inesperado - mostrar pero no bloquear
                System.Diagnostics.Debug.WriteLine($"Error al verificar actualizaciones: {ex.Message}");
            }

            return false;
        }

        private static bool IsNewerVersion(string latestVersion, string currentVersion)
        {
            try
            {
                var latest = new Version(latestVersion);
                var current = new Version(currentVersion);
                return latest > current;
            }
            catch
            {
                return false;
            }
        }

        private static async Task<bool> DownloadAndInstallUpdate(string version, Form parentForm)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "StockControl_Update");
            string zipPath = Path.Combine(tempPath, $"StockControl_v{version}.zip");
            string extractPath = Path.Combine(tempPath, "extracted");

            try
            {
                // Crear carpetas temporales
                if (Directory.Exists(tempPath))
                    Directory.Delete(tempPath, true);
                Directory.CreateDirectory(tempPath);
                Directory.CreateDirectory(extractPath);

                // Descargar actualización
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(5);
                    string downloadUrl = string.Format(RELEASE_URL, version);

                    var progressForm = new Form
                    {
                        Text = "Descargando actualización...",
                        Size = new System.Drawing.Size(400, 120),
                        FormBorderStyle = FormBorderStyle.FixedDialog,
                        StartPosition = FormStartPosition.CenterParent,
                        MaximizeBox = false,
                        MinimizeBox = false
                    };

                    var label = new Label
                    {
                        Text = $"Descargando StockControl v{version}...",
                        Location = new System.Drawing.Point(20, 20),
                        AutoSize = true
                    };

                    var progressBar = new ProgressBar
                    {
                        Location = new System.Drawing.Point(20, 50),
                        Size = new System.Drawing.Size(340, 23),
                        Style = ProgressBarStyle.Marquee
                    };

                    progressForm.Controls.Add(label);
                    progressForm.Controls.Add(progressBar);
                    progressForm.Show();

                    var response = await client.GetAsync(downloadUrl);
                    response.EnsureSuccessStatusCode();

                    using (var fs = new FileStream(zipPath, FileMode.Create))
                    {
                        await response.Content.CopyToAsync(fs);
                    }

                    progressForm.Close();
                }

                // Extraer archivos
                ZipFile.ExtractToDirectory(zipPath, extractPath);

                // Crear script de actualización
                string batchPath = Path.Combine(tempPath, "update.bat");
                string currentDir = AppDomain.CurrentDomain.BaseDirectory;
                string exePath = Process.GetCurrentProcess().MainModule?.FileName ?? "";

                string batchContent = $@"@echo off
timeout /t 2 /nobreak > nul
echo Aplicando actualizacion...
xcopy ""{extractPath}\*.*"" ""{currentDir}"" /E /Y /I /Q
echo Iniciando StockControl...
start """" ""{exePath}""
del ""%~f0""
";

                File.WriteAllText(batchPath, batchContent);

                // Ejecutar script y cerrar aplicación
                Process.Start(new ProcessStartInfo
                {
                    FileName = batchPath,
                    CreateNoWindow = true,
                    UseShellExecute = true
                });

                Application.Exit();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error al descargar la actualización:\n\n{ex.Message}\n\n" +
                    $"Por favor, descargue manualmente desde:\n" +
                    $"https://github.com/NandoPls/StockControl/releases",
                    "Error de Actualización",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                // Limpiar archivos temporales
                try
                {
                    if (Directory.Exists(tempPath))
                        Directory.Delete(tempPath, true);
                }
                catch { }

                return false;
            }
        }
    }
}

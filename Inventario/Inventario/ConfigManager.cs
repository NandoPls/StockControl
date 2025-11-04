using System;
using System.IO;
using System.Text.Json;

namespace Inventario
{
    /// <summary>
    /// Administrador de configuración de la aplicación
    /// Lee y escribe appsettings.json
    /// </summary>
    public static class ConfigManager
    {
        private static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        private static AppConfig? configCache = null;

        /// <summary>
        /// Carga la configuración desde appsettings.json
        /// </summary>
        public static AppConfig CargarConfiguracion()
        {
            if (configCache != null)
                return configCache;

            try
            {
                if (!File.Exists(ConfigPath))
                {
                    // Crear configuración por defecto si no existe
                    var defaultConfig = new AppConfig
                    {
                        SapConnection = new SapConfig
                        {
                            Enabled = false,
                            Server = "SERVIDOR\\SQLEXPRESS",
                            Database = "SBO_EMPRESA",
                            Username = "sa",
                            Password = "",
                            UseWindowsAuth = false,
                            ConnectionTimeout = 30
                        },
                        General = new GeneralConfig
                        {
                            DefaultDataSource = "Excel",
                            AutoBackupEnabled = true,
                            AutoBackupIntervalMinutes = 2
                        }
                    };

                    GuardarConfiguracion(defaultConfig);
                    return defaultConfig;
                }

                string json = File.ReadAllText(ConfigPath);
                var config = JsonSerializer.Deserialize<AppConfig>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                configCache = config ?? new AppConfig();
                return configCache;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al cargar configuración: {ex.Message}");
                return new AppConfig(); // Retornar config por defecto
            }
        }

        /// <summary>
        /// Guarda la configuración en appsettings.json
        /// </summary>
        public static bool GuardarConfiguracion(AppConfig config)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                string json = JsonSerializer.Serialize(config, options);
                File.WriteAllText(ConfigPath, json);

                configCache = config;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al guardar configuración: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Limpia el caché de configuración
        /// </summary>
        public static void LimpiarCache()
        {
            configCache = null;
        }

        /// <summary>
        /// Verifica si SAP está configurado y habilitado
        /// </summary>
        public static bool SapEstaHabilitado()
        {
            var config = CargarConfiguracion();
            return config.SapConnection?.Enabled ?? false;
        }

        /// <summary>
        /// Obtiene la fuente de datos por defecto (Excel o SAP)
        /// </summary>
        public static string ObtenerFuenteDatos()
        {
            var config = CargarConfiguracion();
            return config.General?.DefaultDataSource ?? "Excel";
        }
    }

    /// <summary>
    /// Configuración completa de la aplicación
    /// </summary>
    public class AppConfig
    {
        public SapConfig SapConnection { get; set; } = new SapConfig();
        public GeneralConfig General { get; set; } = new GeneralConfig();
    }

    /// <summary>
    /// Configuración general de la aplicación
    /// </summary>
    public class GeneralConfig
    {
        public string DefaultDataSource { get; set; } = "Excel";
        public bool AutoBackupEnabled { get; set; } = true;
        public int AutoBackupIntervalMinutes { get; set; } = 2;
    }
}

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace Inventario
{
    /// <summary>
    /// Conector para integración con SAP Business One
    /// Actualmente configurado para conexión SQL directa (solo lectura)
    /// Futuro: Implementar Service Layer para escritura
    /// </summary>
    public static class SapConnector
    {
        private static string connectionString = "";
        private static bool isConfigured = false;

        /// <summary>
        /// Configura la conexión a SAP B1
        /// </summary>
        public static bool ConfigurarConexion(SapConfig config)
        {
            try
            {
                if (!config.Enabled)
                {
                    isConfigured = false;
                    return false;
                }

                // Construir connection string
                if (config.UseWindowsAuth)
                {
                    connectionString = $"Server={config.Server};Database={config.Database};Integrated Security=true;Connection Timeout={config.ConnectionTimeout};";
                }
                else
                {
                    connectionString = $"Server={config.Server};Database={config.Database};User Id={config.Username};Password={config.Password};Connection Timeout={config.ConnectionTimeout};";
                }

                // Probar conexión
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    isConfigured = true;
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al configurar SAP: {ex.Message}");
                isConfigured = false;
                return false;
            }
        }

        /// <summary>
        /// Verifica si hay conexión activa con SAP
        /// </summary>
        public static bool EstaConectado()
        {
            if (!isConfigured) return false;

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Obtiene lista de almacenes disponibles en SAP
        /// </summary>
        public static List<string> ObtenerAlmacenes()
        {
            var almacenes = new List<string>();

            if (!isConfigured) return almacenes;

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                        SELECT WhsCode, WhsName
                        FROM OWHS
                        WHERE Locked = 'N'
                        ORDER BY WhsCode
                    ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string codigo = reader["WhsCode"].ToString() ?? "";
                            string nombre = reader["WhsName"].ToString() ?? "";
                            almacenes.Add($"{codigo} - {nombre}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al obtener almacenes: {ex.Message}");
            }

            return almacenes;
        }

        /// <summary>
        /// Carga productos desde SAP B1 para un almacén específico
        /// IMPORTANTE: Solo lectura, no modifica datos en SAP
        /// </summary>
        public static List<ProductoExcel> CargarProductosDesdeSap(string codigoAlmacen)
        {
            var productos = new List<ProductoExcel>();

            if (!isConfigured)
            {
                throw new InvalidOperationException("SAP no está configurado. Verifique appsettings.json");
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // Query adaptada a la estructura de SAP B1
                    // OITM: Items Master Data
                    // OITW: Item Warehouse Info
                    // OITB: Item Groups
                    string query = @"
                        SELECT
                            T0.ItemCode,
                            T0.ItemName,
                            T0.CodeBars,
                            ISNULL(T1.OnHand, 0) AS OnHand,
                            T1.WhsCode,
                            T2.ItmsGrpNam,
                            ISNULL(T0.U_Comercial1, '') AS U_Comercial1,
                            ISNULL(T0.U_Comercial3, '') AS U_Comercial3
                        FROM OITM T0
                        LEFT JOIN OITW T1 ON T0.ItemCode = T1.ItemCode
                        LEFT JOIN OITB T2 ON T0.ItmsGrpCod = T2.ItmsGrpCod
                        WHERE T1.WhsCode = @WhsCode
                        AND T0.InvntItem = 'Y'
                        AND T0.validFor = 'Y'
                        ORDER BY T0.ItemCode
                    ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@WhsCode", codigoAlmacen);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var producto = new ProductoExcel
                                {
                                    ItemCode = reader["ItemCode"]?.ToString() ?? "",
                                    CodeBars = reader["CodeBars"]?.ToString() ?? "",
                                    StockTienda = Convert.ToInt32(reader["OnHand"]),
                                    WhsCode = reader["WhsCode"]?.ToString() ?? "",
                                    ItmsGrpNam = reader["ItmsGrpNam"]?.ToString() ?? "",
                                    U_Comercial1 = reader["U_Comercial1"]?.ToString() ?? "",
                                    U_Comercial3 = reader["U_Comercial3"]?.ToString() ?? ""
                                };

                                // Si no tiene código de barras, usar ItemCode
                                if (string.IsNullOrWhiteSpace(producto.CodeBars))
                                {
                                    producto.CodeBars = producto.ItemCode;
                                }

                                productos.Add(producto);
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                throw new Exception($"Error de SQL al conectar con SAP: {ex.Message}\n\nVerifique:\n- Servidor y base de datos correctos\n- Credenciales válidas\n- SQL Server accesible", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar productos desde SAP: {ex.Message}", ex);
            }

            return productos;
        }

        /// <summary>
        /// Obtiene información detallada de un producto por código de barras
        /// </summary>
        public static ProductoExcel? ObtenerProductoPorEan(string ean, string almacen)
        {
            if (!isConfigured) return null;

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"
                        SELECT TOP 1
                            T0.ItemCode,
                            T0.ItemName,
                            T0.CodeBars,
                            ISNULL(T1.OnHand, 0) AS OnHand,
                            T1.WhsCode,
                            T2.ItmsGrpNam,
                            ISNULL(T0.U_Comercial1, '') AS U_Comercial1,
                            ISNULL(T0.U_Comercial3, '') AS U_Comercial3
                        FROM OITM T0
                        LEFT JOIN OITW T1 ON T0.ItemCode = T1.ItemCode
                        LEFT JOIN OITB T2 ON T0.ItmsGrpCod = T2.ItmsGrpCod
                        WHERE (T0.CodeBars = @EAN OR T0.ItemCode = @EAN)
                        AND T1.WhsCode = @WhsCode
                        AND T0.InvntItem = 'Y'
                    ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@EAN", ean);
                        cmd.Parameters.AddWithValue("@WhsCode", almacen);

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new ProductoExcel
                                {
                                    ItemCode = reader["ItemCode"]?.ToString() ?? "",
                                    CodeBars = reader["CodeBars"]?.ToString() ?? ean,
                                    StockTienda = Convert.ToInt32(reader["OnHand"]),
                                    WhsCode = reader["WhsCode"]?.ToString() ?? "",
                                    ItmsGrpNam = reader["ItmsGrpNam"]?.ToString() ?? "",
                                    U_Comercial1 = reader["U_Comercial1"]?.ToString() ?? "",
                                    U_Comercial3 = reader["U_Comercial3"]?.ToString() ?? ""
                                };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al buscar producto: {ex.Message}");
            }

            return null;
        }

        // ==========================================
        // MÉTODOS FUTUROS PARA SERVICE LAYER
        // ==========================================
        // TODO: Implementar cuando tengas Service Layer configurado

        /// <summary>
        /// [FUTURO] Crear ajuste de inventario en SAP usando Service Layer
        /// Requiere: SAP B1 9.3+ y Service Layer configurado
        /// </summary>
        public static bool CrearAjusteInventario(Dictionary<string, int> diferencias, string almacen, string comentarios)
        {
            // TODO: Implementar cuando Service Layer esté disponible
            /*
            usando REST API a:
            POST https://servidor:50000/b1s/v1/InventoryGenEntries
            {
                "Comments": comentarios,
                "DocumentLines": [
                    {
                        "ItemCode": "ITEM001",
                        "WarehouseCode": "01",
                        "Quantity": diferencia
                    }
                ]
            }
            */

            throw new NotImplementedException("Service Layer no configurado. Esta función estará disponible en futuras versiones.");
        }

        /// <summary>
        /// [FUTURO] Actualizar stock mediante DI API
        /// Requiere: SAP DI API SDK instalado
        /// </summary>
        public static bool ActualizarStockConDiApi(string itemCode, string almacen, int nuevaCantidad)
        {
            // TODO: Implementar cuando DI API esté instalado
            /*
            Requiere referencia a:
            - SAPbobsCOM.dll

            Company oCompany = new Company();
            oCompany.Connect();
            InventoryGenEntry oInv = oCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
            // ... configurar y Add()
            */

            throw new NotImplementedException("DI API no instalado. Esta función requiere SAP DI API SDK.");
        }
    }

    /// <summary>
    /// Configuración de conexión SAP
    /// </summary>
    public class SapConfig
    {
        public bool Enabled { get; set; }
        public string Server { get; set; } = "";
        public string Database { get; set; } = "";
        public string Username { get; set; } = "";
        public string Password { get; set; } = "";
        public bool UseWindowsAuth { get; set; }
        public int ConnectionTimeout { get; set; } = 30;
    }
}

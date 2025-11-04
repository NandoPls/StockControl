using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace Inventario
{
    public static class ExcelDataManager
    {
        public static List<ProductoExcel> ProductosExcel { get; private set; }
        private static string origenDatos = "Excel"; // "Excel" o "SAP"

        /// <summary>
        /// Configura el origen de datos (Excel o SAP)
        /// </summary>
        public static void ConfigurarOrigenDatos(string origen)
        {
            origenDatos = origen;
        }

        /// <summary>
        /// Obtiene el origen de datos actual
        /// </summary>
        public static string ObtenerOrigenDatos()
        {
            return origenDatos;
        }

        /// <summary>
        /// Carga productos desde SAP para un almacén específico
        /// </summary>
        public static bool CargarDesdeSap(string almacen)
        {
            try
            {
                ProductosExcel = SapConnector.CargarProductosDesdeSap(almacen);
                origenDatos = "SAP";
                return ProductosExcel.Count > 0;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar desde SAP: {ex.Message}", ex);
            }
        }

        public static bool CargarExcel(string rutaArchivo)
        {
            try
            {
                ProductosExcel = new List<ProductoExcel>();

                // Verificar que el archivo no esté en uso
                try
                {
                    using (var fileStream = new System.IO.FileStream(rutaArchivo,
                        System.IO.FileMode.Open,
                        System.IO.FileAccess.Read,
                        System.IO.FileShare.None))
                    {
                        // Si llegamos aquí, el archivo no está en uso
                    }
                }
                catch (System.IO.IOException)
                {
                    throw new System.IO.IOException($"El archivo está abierto en otra aplicación. Por favor, cierre el archivo Excel para continuar.");
                }

                using (var workbook = new XLWorkbook(rutaArchivo))
                {
                    var worksheet = workbook.Worksheet(1);

                    // Buscar la fila de encabezados (fila 3)
                    var filaInicio = 4; // Los datos comienzan en la fila 4

                    var filasConDatos = worksheet.RowsUsed().Skip(3); // Skip filas 1, 2, 3

                    foreach (var fila in filasConDatos)
                    {
                        try
                        {
                            var producto = new ProductoExcel
                            {
                                // Columna A: MARCA (ItmsGrpNam)
                                ItmsGrpNam = fila.Cell(1).GetValue<string>()?.Trim() ?? "",

                                // Columna B: Clasificación (U_Comercial1)
                                U_Comercial1 = fila.Cell(2).GetValue<string>()?.Trim() ?? "",

                                // Columna C: Clasificación detallada (U_Comercial3)
                                U_Comercial3 = fila.Cell(3).GetValue<string>()?.Trim() ?? "",

                                // Columna D: CODIGO MADRE (ItemCode)
                                ItemCode = fila.Cell(4).GetValue<string>()?.Trim() ?? "",

                                // Columna E: EAN (código de barras) (CodeBars)
                                CodeBars = fila.Cell(5).GetValue<string>()?.Trim() ?? "",

                                // Columna F: STOCK (Stock Tienda)
                                StockTienda = ConvertirAEntero(fila.Cell(6).GetValue<string>()),

                                // Columna G: ALMACEN (WhsCode)
                                WhsCode = fila.Cell(7).GetValue<string>()?.Trim() ?? ""
                            };

                            // Solo agregar si tiene datos válidos
                            if (!string.IsNullOrEmpty(producto.ItemCode) && !string.IsNullOrEmpty(producto.CodeBars))
                            {
                                ProductosExcel.Add(producto);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log o continuar con la siguiente fila
                            Console.WriteLine($"Error en fila {fila.RowNumber()}: {ex.Message}");
                            continue;
                        }
                    }
                }

                return ProductosExcel.Any();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar el archivo Excel: {ex.Message}", ex);
            }
        }

        private static int ConvertirAEntero(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return 0;

            // Intentar convertir directamente
            if (int.TryParse(valor.Trim(), out int resultado))
                return resultado;

            // Intentar convertir si es decimal y tomar la parte entera
            if (double.TryParse(valor.Trim(), out double resultadoDouble))
                return (int)resultadoDouble;

            return 0;
        }

        public static List<string> ObtenerAlmacenes()
        {
            if (ProductosExcel == null || !ProductosExcel.Any())
                return new List<string>();

            return ProductosExcel
                .Select(p => p.WhsCode)
                .Distinct()
                .Where(w => !string.IsNullOrEmpty(w))
                .OrderBy(w => w)
                .ToList();
        }

        public static List<string> ObtenerClasificacionesPorAlmacen(string almacen)
        {
            if (ProductosExcel == null || !ProductosExcel.Any())
                return new List<string>();

            return ProductosExcel
                .Where(p => p.WhsCode == almacen)
                .Select(p => p.U_Comercial1)
                .Distinct()
                .Where(c => !string.IsNullOrEmpty(c))
                .OrderBy(c => c)
                .ToList();
        }

        public static List<ProductoExcel> ObtenerProductosPorAlmacenYClasificacion(string almacen, string clasificacion)
        {
            if (ProductosExcel == null || !ProductosExcel.Any())
                return new List<ProductoExcel>();

            return ProductosExcel
                .Where(p => p.WhsCode == almacen && p.U_Comercial1 == clasificacion)
                .ToList();
        }

        public static ProductoExcel BuscarPorCodigoBarras(string codigoBarras)
        {
            if (ProductosExcel == null || !ProductosExcel.Any())
                return null;

            return ProductosExcel.FirstOrDefault(p => p.CodeBars == codigoBarras);
        }
    }

    public class ProductoExcel
    {
        public string ItmsGrpNam { get; set; }          // MARCA (Columna A)
        public string U_Comercial1 { get; set; }        // Clasificación (Columna B)
        public string U_Comercial3 { get; set; }        // Clasificación detallada (Columna C)
        public string ItemCode { get; set; }            // CODIGO MADRE (Columna D)
        public string CodeBars { get; set; }            // EAN - Código de barras (Columna E)
        public int StockTienda { get; set; }            // STOCK (Columna F)
        public string WhsCode { get; set; }             // ALMACEN (Columna G)
    }
}

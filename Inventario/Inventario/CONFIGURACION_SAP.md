# 🔌 Configuración de Integración con SAP Business One

StockControl v1.2.2 incluye soporte para integración directa con SAP Business One. Esta guía te ayudará a configurarlo.

## 📋 Requisitos Previos

- SAP Business One instalado y funcionando
- SQL Server accesible (mismo servidor o remoto)
- Credenciales de base de datos SAP
- StockControl v1.2.2 o superior

## ⚙️ Configuración Paso a Paso

### 1. Ubicar el archivo de configuración

El archivo `appsettings.json` se encuentra en la misma carpeta que `Inventario.exe`:

```
C:\...\Programa Inventario\appsettings.json
```

### 2. Editar appsettings.json

Abre el archivo con **Notepad** o cualquier editor de texto y modifica los valores:

```json
{
  "SapConnection": {
    "Enabled": true,                    ← Cambiar a true para activar SAP
    "Server": "SERVIDOR\\SQLEXPRESS",   ← Tu servidor SQL
    "Database": "SBO_EMPRESA",          ← Nombre de tu base de datos SAP
    "Username": "sa",                   ← Usuario SQL
    "Password": "tu_contraseña",        ← Contraseña SQL
    "UseWindowsAuth": false,            ← true si usas autenticación Windows
    "ConnectionTimeout": 30
  },
  "General": {
    "DefaultDataSource": "SAP",         ← Cambiar a "SAP" para usar por defecto
    "AutoBackupEnabled": true,
    "AutoBackupIntervalMinutes": 2
  }
}
```

### 3. Valores a completar

#### **Server** (Servidor SQL)
- Formato local: `NOMBREPC\\SQLEXPRESS`
- Formato red: `192.168.1.100\\SQLEXPRESS`
- Formato nombre: `SERVIDOR-SAP\\SQLEXPRESS`
- Solo servidor: `SERVIDOR-SAP` (si usa instancia por defecto)

#### **Database** (Base de Datos)
- Generalmente tiene formato: `SBO_EMPRESA`
- Ejemplos: `SBO_DEMO`, `SBO_PRODUCCION`, `SBO_MIEMPRESA`
- Lo puedes verificar en SQL Server Management Studio

#### **Username y Password**
- Usuario SQL (comúnmente `sa`)
- Contraseña del usuario SQL
- **IMPORTANTE**: Guarda este archivo de forma segura, contiene credenciales

#### **UseWindowsAuth**
- `true`: Usa tu usuario de Windows actual (no necesita Username/Password)
- `false`: Usa credenciales SQL (Username/Password requeridos)

### 4. Guardar y probar

1. Guarda el archivo `appsettings.json`
2. Ejecuta `Inventario.exe`
3. El programa intentará conectarse automáticamente a SAP

## 🔍 Verificar Conexión

### Conexión Exitosa ✅
Si todo está correcto, verás:
```
✅ Conectado a SAP Business One exitosamente.
Los datos se cargarán al seleccionar almacén y clasificación.
```

### Error de Conexión ❌
Si hay un problema, verás un mensaje indicando:
- **"No se pudo conectar"**: Verifica Server y Database
- **"Error de login"**: Verifica Username y Password
- **"Timeout"**: El servidor no es accesible (firewall/red)

## 🎯 Modo de Uso

### Con SAP Habilitado

1. Ejecuta `Inventario.exe`
2. Se conectará automáticamente a SAP
3. **No necesitas cargar Excel** - Los datos vienen de SAP
4. Selecciona Almacén y Clasificaciones
5. Todo lo demás funciona igual

### Volver a Excel

Si quieres volver a usar Excel:
```json
{
  "SapConnection": {
    "Enabled": false,    ← Cambiar a false
    ...
  }
}
```

## 📊 Estructura de Datos SAP

El programa lee las siguientes tablas de SAP B1:

| Tabla | Descripción | Campos Usados |
|-------|-------------|---------------|
| **OITM** | Items Master Data | ItemCode, ItemName, CodeBars, U_Comercial1, U_Comercial3 |
| **OITW** | Item Warehouse Info | WhsCode, OnHand |
| **OITB** | Item Groups | ItmsGrpNam |
| **OWHS** | Warehouses | WhsCode, WhsName |

## ⚠️ Notas Importantes

### Permisos SQL
- El usuario SQL debe tener permisos de **LECTURA** en las tablas de SAP
- No se requieren permisos de escritura (por ahora solo lectura)
- No se modifican datos en SAP en esta versión

### Seguridad
- **NO** compartas tu `appsettings.json` - contiene credenciales
- Considera usar autenticación Windows (`UseWindowsAuth: true`) para mayor seguridad
- El archivo se copia junto al ejecutable en cada actualización

### Rendimiento
- La carga inicial puede tomar más tiempo que Excel
- Depende de la cantidad de productos y velocidad de red
- Usa el filtro de almacén para reducir datos

## 🚀 Funcionalidades Futuras

### v1.3.0 (Próximamente)
- ✨ Escritura de ajustes de inventario directamente en SAP
- ✨ Integración con Service Layer (REST API)
- ✨ Soporte para DI API oficial de SAP
- ✨ Creación automática de documentos de entrada/salida

### v1.4.0 (Planificado)
- 📊 Reportes directos en SAP Crystal Reports
- 🔄 Sincronización bidireccional
- 📱 Acceso remoto vía web

## 🆘 Solución de Problemas

### "Server not found"
- Verifica que el nombre del servidor es correcto
- Prueba con la IP en lugar del nombre
- Verifica que SQL Server Browser está ejecutándose

### "Login failed"
- Usuario o contraseña incorrectos
- El usuario no tiene permisos en la base de datos
- Prueba con `UseWindowsAuth: true`

### "Database not found"
- El nombre de la base de datos está mal escrito
- La base de datos no existe
- Verifica con SQL Management Studio

### Firewall/Red
- Puerto SQL Server (1433) debe estar abierto
- Firewall de Windows permite SQL Server
- Red permite conexión al servidor

## 📞 Soporte

Para problemas de configuración:
1. Verifica los logs en la consola de Windows
2. Contacta al administrador de SAP de tu empresa
3. Revisa la documentación de SAP Business One

---

**Desarrollado por Fernando Carrasco**
**StockControl v1.2.2**

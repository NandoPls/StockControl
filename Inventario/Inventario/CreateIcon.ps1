# Script para crear un icono simple para InventoryPro
Add-Type -AssemblyName System.Drawing

$iconSize = 256
$bitmap = New-Object System.Drawing.Bitmap($iconSize, $iconSize)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# Activar antialiasing
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias

# Fondo azul
$blueColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
$brush = New-Object System.Drawing.SolidBrush($blueColor)
$graphics.FillRectangle($brush, 0, 0, $iconSize, $iconSize)

# Dibujar caja de inventario (rectángulo blanco)
$whitePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 12)
$boxRect = New-Object System.Drawing.Rectangle(40, 50, 176, 156)
$graphics.DrawRectangle($whitePen, $boxRect)

# Dibujar líneas horizontales dentro (estantes)
$whitePen2 = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 8)
$graphics.DrawLine($whitePen2, 40, 110, 216, 110)
$graphics.DrawLine($whitePen2, 40, 160, 216, 160)

# Dibujar checkmark grande
$greenColor = [System.Drawing.Color]::FromArgb(46, 204, 113)
$greenBrush = New-Object System.Drawing.SolidBrush($greenColor)
$checkPen = New-Object System.Drawing.Pen($greenBrush, 20)
$checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round

$graphics.DrawLine($checkPen, 145, 165, 175, 195)
$graphics.DrawLine($checkPen, 175, 195, 225, 135)

# Guardar como PNG primero
$pngPath = "C:\Users\Nandy\source\repos\NewRepo\inventario\Inventario\icon_temp.png"
$bitmap.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

Write-Host "Icono PNG creado en: $pngPath"
Write-Host "Para convertir a ICO, usa una herramienta online o ImageMagick"

$graphics.Dispose()
$bitmap.Dispose()

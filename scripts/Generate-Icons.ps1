[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Run this asset generator with powershell.exe -NoProfile -STA -File scripts\Generate-Icons.ps1.'
}
Add-Type -AssemblyName PresentationCore, WindowsBase
$assets = Join-Path (Split-Path $PSScriptRoot -Parent) 'assets'
[xml]$document = Get-Content (Join-Path $assets 'orglens.svg') -Raw
$svg = $document.DocumentElement
if ($svg.viewBox -ne '0 0 24 24' -or $svg.fill -ne 'none') {
    throw 'The icon generator expects the 24px outline SVG grid.'
}
$brush = [Windows.Media.BrushConverter]::new().ConvertFromInvariantString($svg.stroke)
$pen = [Windows.Media.Pen]::new($brush, [double]::Parse($svg.'stroke-width', [Globalization.CultureInfo]::InvariantCulture))
$pen.StartLineCap = $pen.EndLineCap = [Windows.Media.PenLineCap]::Round
$pen.LineJoin = [Windows.Media.PenLineJoin]::Round
$frames = @()
foreach ($size in @(16, 20, 24, 32, 40, 48, 64, 128, 256)) {
    $visual = [Windows.Media.DrawingVisual]::new()
    $context = $visual.RenderOpen()
    try {
        $context.PushTransform([Windows.Media.ScaleTransform]::new($size / 24.0, $size / 24.0))
        foreach ($element in $svg.ChildNodes) {
            switch ($element.LocalName) {
                'rect' {
                    $rectangle = [Windows.Rect]::new([double]$element.x, [double]$element.y, [double]$element.width, [double]$element.height)
                    $geometry = [Windows.Media.RectangleGeometry]::new($rectangle, [double]$element.rx, [double]$element.rx)
                }
                'path' { $geometry = [Windows.Media.Geometry]::Parse($element.d) }
                default { throw "Unsupported SVG element: $($element.LocalName)" }
            }
            $context.DrawGeometry($null, $pen, $geometry)
        }
        $context.Pop()
    } finally { $context.Close() }
    $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new($size, $size, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.MemoryStream]::new()
    try {
        $encoder.Save($stream)
        $data = $stream.ToArray()
        $frames += [pscustomobject]@{ Size = $size; Data = $data }
        if ($size -in @(32, 64)) { [IO.File]::WriteAllBytes((Join-Path $assets "orglens-$size.png"), $data) }
    } finally { $stream.Dispose() }
}
$writer = [IO.BinaryWriter]::new([IO.File]::Create((Join-Path $assets 'orglens.ico')))
try {
    $writer.Write([uint16]0)
    $writer.Write([uint16]1)
    $writer.Write([uint16]$frames.Count)
    $offset = 6 + 16 * $frames.Count
    foreach ($frame in $frames) {
        $dimension = if ($frame.Size -eq 256) { 0 } else { $frame.Size }
        $writer.Write([byte]$dimension)
        $writer.Write([byte]$dimension)
        $writer.Write([byte]0)
        $writer.Write([byte]0)
        $writer.Write([uint16]1)
        $writer.Write([uint16]32)
        $writer.Write([uint32]$frame.Data.Length)
        $writer.Write([uint32]$offset)
        $offset += $frame.Data.Length
    }
    foreach ($frame in $frames) { $writer.Write([byte[]]$frame.Data) }
} finally { $writer.Dispose() }
Write-Host 'Generated transparent 32px/64px PNGs and a nine-size Windows ICO from orglens.svg.'

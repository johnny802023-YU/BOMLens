param(
  [Parameter(Mandatory = $true)]
  [string]$ExecutablePath,
  [string]$IconPath = "public\BOMLens.ico"
)

$ErrorActionPreference = "Stop"
$exe = (Resolve-Path $ExecutablePath).Path
$icon = (Resolve-Path $IconPath).Path
Add-Type -AssemblyName System.Drawing

$embeddedIcon = $null
$sourceIcon = $null
$embeddedBitmap = $null
$sourceBitmap = $null
$normalizedEmbedded = $null
try {
  $embeddedIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($exe)
  if ($null -eq $embeddedIcon) { throw "執行檔沒有可讀取的 Windows icon：$exe" }

  $sourceIcon = [System.Drawing.Icon]::new($icon, 32, 32)
  $sourceBitmap = $sourceIcon.ToBitmap()
  $embeddedBitmap = $embeddedIcon.ToBitmap()
  $normalizedEmbedded = [System.Drawing.Bitmap]::new($embeddedBitmap, 32, 32)

  $differentPixels = 0
  $totalPixels = 32 * 32
  for ($x = 0; $x -lt 32; $x++) {
    for ($y = 0; $y -lt 32; $y++) {
      $expected = $sourceBitmap.GetPixel($x, $y)
      $actual = $normalizedEmbedded.GetPixel($x, $y)
      $distance = [Math]::Abs([int]$expected.A - [int]$actual.A) +
        [Math]::Abs([int]$expected.R - [int]$actual.R) +
        [Math]::Abs([int]$expected.G - [int]$actual.G) +
        [Math]::Abs([int]$expected.B - [int]$actual.B)
      if ($distance -gt 24) { $differentPixels++ }
    }
  }

  $differenceRatio = $differentPixels / $totalPixels
  if ($differenceRatio -gt 0.05) {
    throw "執行檔 icon 與 BOMLens.ico 不一致（差異 $([Math]::Round($differenceRatio * 100, 1))%）：$exe"
  }
  Write-Host "Windows icon 驗證通過：$exe"
} finally {
  if ($null -ne $normalizedEmbedded) { $normalizedEmbedded.Dispose() }
  if ($null -ne $embeddedBitmap) { $embeddedBitmap.Dispose() }
  if ($null -ne $sourceBitmap) { $sourceBitmap.Dispose() }
  if ($null -ne $sourceIcon) { $sourceIcon.Dispose() }
  if ($null -ne $embeddedIcon) { $embeddedIcon.Dispose() }
}

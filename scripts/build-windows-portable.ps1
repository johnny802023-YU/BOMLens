param(
  [string]$OutputDirectory = "release",
  [string]$CertificateThumbprint = "",
  [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$project = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $project $OutputDirectory
$packageRoot = Join-Path $outputRoot "BOMLens-Windows-x64-Portable"

Set-Location $project
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw "建置電腦需要先安裝 Node.js 22 或更新版本。" }
if (-not (Test-Path "node_modules")) { throw "找不到 node_modules，請先執行 npm ci。" }

if (-not $SkipBuild) {
  & npm run build
  if ($LASTEXITCODE -ne 0) { throw "BOMLens 建置失敗。" }
} elseif (-not (Test-Path "dist")) {
  throw "使用 -SkipBuild 時必須已存在 dist，請先執行 npm run build。"
}

Remove-Item $packageRoot -Recurse -Force -ErrorAction SilentlyContinue
New-Item $packageRoot -ItemType Directory -Force | Out-Null
New-Item (Join-Path $packageRoot "runtime") -ItemType Directory -Force | Out-Null

Copy-Item "dist" (Join-Path $packageRoot "dist") -Recurse
Copy-Item "node_modules" (Join-Path $packageRoot "node_modules") -Recurse
if (Test-Path "public") { Copy-Item "public" (Join-Path $packageRoot "public") -Recurse }
Copy-Item "package.json" (Join-Path $packageRoot "package.json")
Copy-Item (Get-Command node).Source (Join-Path $packageRoot "runtime\node.exe")
Copy-Item "LOCAL_OFFLINE_GUIDE.md" (Join-Path $packageRoot "使用說明.md")

$launcherSource = @'
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows.Forms;

public static class BOMLensLauncher {
  private static Process server;
  [STAThread]
  public static void Main() {
    string root = AppDomain.CurrentDomain.BaseDirectory;
    string node = Path.Combine(root, "runtime", "node.exe");
    string cli = Path.Combine(root, "node_modules", "vinext", "dist", "cli.js");
    if (!File.Exists(node) || !File.Exists(cli)) {
      MessageBox.Show("BOMLens 可攜式檔案不完整，請重新解壓縮。", "BOMLens", MessageBoxButtons.OK, MessageBoxIcon.Error);
      return;
    }
    var info = new ProcessStartInfo(node, "\"" + cli + "\" start --hostname 127.0.0.1 --port 3784") {
      WorkingDirectory = root,
      UseShellExecute = false,
      CreateNoWindow = true,
      WindowStyle = ProcessWindowStyle.Hidden
    };
    info.EnvironmentVariables["WRANGLER_LOG_PATH"] = Path.Combine(root, ".wrangler", "wrangler.log");
    server = Process.Start(info);
    Application.EnableVisualStyles();
    var menu = new ContextMenuStrip();
    menu.Items.Add("開啟 BOMLens", null, (s, e) => OpenBrowser());
    menu.Items.Add("結束", null, (s, e) => Application.Exit());
    var tray = new NotifyIcon { Icon = SystemIcons.Shield, Text = "BOMLens 離線模式", Visible = true, ContextMenuStrip = menu };
    tray.DoubleClick += (s, e) => OpenBrowser();
    new Thread(() => {
      for (int i = 0; i < 40; i++) {
        try { using (var client = new WebClient()) { client.DownloadString("http://127.0.0.1:3784/"); } OpenBrowser(); return; }
        catch { Thread.Sleep(500); }
      }
      MessageBox.Show("BOMLens 本機服務啟動逾時，請重新啟動程式。", "BOMLens", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }) { IsBackground = true }.Start();
    Application.ApplicationExit += (s, e) => { tray.Visible = false; try { if (server != null && !server.HasExited) server.Kill(); } catch {} };
    Application.Run();
  }
  private static void OpenBrowser() {
    try { Process.Start(new ProcessStartInfo("http://127.0.0.1:3784/") { UseShellExecute = true }); } catch {}
  }
}
'@

$launcher = Join-Path $packageRoot "BOMLens.exe"
$launcherSourceFile = Join-Path $outputRoot "BOMLensLauncher.cs"
$compiler = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $compiler)) { throw "找不到 Windows 64 位元 C# 編譯器：$compiler" }
$launcherSource | Set-Content $launcherSourceFile -Encoding UTF8
& $compiler /nologo /target:winexe /platform:x64 "/out:$launcher" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll $launcherSourceFile
if ($LASTEXITCODE -ne 0) { throw "BOMLens.exe 編譯失敗。" }
Remove-Item $launcherSourceFile -Force

if ($CertificateThumbprint) {
  $certificate = Get-ChildItem "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction Stop
  Set-AuthenticodeSignature -FilePath $launcher -Certificate $certificate -TimestampServer "http://timestamp.digicert.com" | Out-Null
}

$zip = Join-Path $outputRoot "BOMLens-Windows-x64-Portable.zip"
Remove-Item $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path "$packageRoot\*" -DestinationPath $zip -CompressionLevel Optimal
Write-Host "完成：$zip"

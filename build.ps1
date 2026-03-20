param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

Write-Host "== Corvus Macro Automator Build ==" -ForegroundColor Cyan
Write-Host "Configuration: $Configuration"

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    throw "uv command not found. Please install uv first."
}

$appName = "CorvusMacroAutomator"
$distDir = "dist"
$buildDir = "build"

if (Test-Path $distDir) {
    Remove-Item -Recurse -Force $distDir
}
if (Test-Path $buildDir) {
    Remove-Item -Recurse -Force $buildDir
}

# Release uses folder mode (--onedir), not single-file mode (--onefile)
$pyInstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--name", $appName,
    "--windowed",
    "--onedir",
    "main.py"
)

if ($Configuration -eq "Debug") {
    # Debug keeps console enabled for troubleshooting
    $pyInstallerArgs = @(
        "--noconfirm",
        "--clean",
        "--name", $appName,
        "--onedir",
        "main.py"
    )
}

Write-Host "Running PyInstaller..."
uv run --with pyinstaller pyinstaller @pyInstallerArgs

Write-Host ""
Write-Host "Build completed." -ForegroundColor Green
Write-Host "Output folder: $distDir\$appName"

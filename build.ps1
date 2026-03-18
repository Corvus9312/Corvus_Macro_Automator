Param(
  # 一般打包 (onedir)；預設不顯示 console
  [switch]$Console
)

$ErrorActionPreference = "Stop"

function Write-Step($msg) {
  Write-Host ""
  Write-Host "==> $msg"
}

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $RepoRoot

$Entry = Join-Path $RepoRoot "main.py"
if (-not (Test-Path $Entry)) {
  throw "找不到入口檔 main.py：$Entry"
}

Write-Step "確認 Python"
$python = (Get-Command python -ErrorAction SilentlyContinue)
if (-not $python) { throw "找不到 python，請先安裝 Python 3.12+" }

Write-Step "更新 pip / build 工具"
python -m pip install --upgrade pip wheel setuptools

if (Test-Path (Join-Path $RepoRoot "pyproject.toml")) {
  Write-Step "安裝專案依賴（從 pyproject.toml）"
  # 以 editable 安裝，方便本機開發/打包一致
  python -m pip install -e .
} elseif (Test-Path (Join-Path $RepoRoot "requirements.txt")) {
  Write-Step "安裝專案依賴（requirements.txt）"
  python -m pip install -r requirements.txt
} else {
  Write-Step "未找到 pyproject.toml / requirements.txt，跳過依賴安裝"
}

Write-Step "安裝/更新 PyInstaller"
python -m pip install --upgrade pyinstaller

Write-Step "清理目標資料夾 (build/dist) 與 spec"
if (Test-Path (Join-Path $RepoRoot "build")) { Remove-Item (Join-Path $RepoRoot "build") -Recurse -Force }
if (Test-Path (Join-Path $RepoRoot "dist")) { Remove-Item (Join-Path $RepoRoot "dist") -Recurse -Force }
Get-ChildItem -Path $RepoRoot -Filter "*.spec" -File -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Step "開始打包"

$args = @(
  "--noconfirm",
  "--clean",
  "--name", "Corvus_Macro_Automator",
  "--distpath", (Join-Path $RepoRoot "dist"),
  "--workpath", (Join-Path $RepoRoot "build")
)

if (-not $Console) {
  $args += "--noconsole"
}

# 常見 hidden imports（避免某些環境 hook 不完整）
$args += @(
  "--hidden-import", "PyQt6.QtNetwork",
  "--hidden-import", "PyQt6.QtGui",
  "--hidden-import", "PyQt6.QtWidgets",
  "--hidden-import", "keyboard",
  "--hidden-import", "pyautogui",
  "--hidden-import", "cv2",
  "--hidden-import", "PIL"
)

$args += $Entry

python -m PyInstaller @args

Write-Step "完成"
Write-Host "輸出位置：" (Join-Path $RepoRoot "dist")
Write-Host "一般用法：" ".\build.ps1"
Write-Host "若要顯示 console：" ".\build.ps1 -Console"

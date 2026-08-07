# ============================================================
#  領収書読取ツール 配布用ランタイム(_runtime)作成  ※メンテ担当者のみ
#
#  Python本体(Embeddable版)＋必要ライブラリを _runtime\python\ に丸ごと入れる。
#  これで配布先PCは Python インストール不要になる（PRPツールと同方式）。
#
#  使い方: PowerShell で 98_dist\配布ランタイム作成.ps1 を実行
#          （ネット接続・5分程度）。作成後 _runtime を含めて配布。
# ============================================================
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$PyVersion = '3.12.10'
$PyTag     = 'python312'

$DistDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root    = Split-Path -Parent $DistDir
$Runtime = Join-Path $Root '_runtime'
$PyDir   = Join-Path $Runtime 'python'
$Work    = Join-Path $Runtime '_tmp'

Write-Host ''
Write-Host '=== 配布用ランタイムを作成します ===' -ForegroundColor Cyan
Write-Host "  出力先: $PyDir"

if (Test-Path $Runtime) { Remove-Item -Recurse -Force $Runtime }
New-Item -ItemType Directory -Force -Path $PyDir | Out-Null
New-Item -ItemType Directory -Force -Path $Work  | Out-Null

# 1) Embeddable Python
$EmbedZip = Join-Path $Work 'python-embed.zip'
Write-Host "[1/4] Python $PyVersion (Embeddable) を取得..."
Invoke-WebRequest -Uri "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-embed-amd64.zip" `
                  -OutFile $EmbedZip -TimeoutSec 300
Expand-Archive -Path $EmbedZip -DestinationPath $PyDir -Force

# 2) ._pth を書き換え（site を有効化＋ src をパスに追加）
#    ._pth の相対パスは _runtime\python\ 基準。src はツール直下＝ ..\..\src
Write-Host '[2/4] sys.path を設定 (site-packages / src)...'
$PthFile = Join-Path $PyDir "$PyTag._pth"
@"
$PyTag.zip
.
Lib\site-packages
..\..\src
import site
"@ | Set-Content -Path $PthFile -Encoding ascii

# 3) pip
Write-Host '[3/4] pip を導入...'
$GetPip = Join-Path $Work 'get-pip.py'
Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $GetPip -TimeoutSec 300
$PyExe = Join-Path $PyDir 'python.exe'
& $PyExe $GetPip --no-warn-script-location
if ($LASTEXITCODE -ne 0) { throw 'pip の導入に失敗' }

# 4) 必要ライブラリ
Write-Host '[4/4] ライブラリを導入 (pymupdf/opencv/pytesseract/Pillow/openpyxl/numpy)...'
& $PyExe -m pip install --no-warn-script-location `
    pymupdf opencv-python-headless pytesseract Pillow openpyxl numpy
if ($LASTEXITCODE -ne 0) { throw 'ライブラリ導入に失敗' }

Remove-Item -Recurse -Force $Work

# 検証: 実際に import できるか
Write-Host ''
Write-Host '=== 検証中 ===' -ForegroundColor Cyan
& $PyExe -c "import fitz,cv2,pytesseract,PIL,openpyxl,numpy; print('  import OK (cv2', cv2.__version__, ')')"
if ($LASTEXITCODE -ne 0) { throw '検証失敗（import不可）' }

$size = [math]::Round(((Get-ChildItem -Recurse $Runtime | Measure-Object -Property Length -Sum).Sum / 1MB), 1)
Write-Host ''
Write-Host "完了。_runtime サイズ: ${size}MB" -ForegroundColor Green
Write-Host '配布は 98_dist\配布ZIP作成.ps1 で行ってください。'


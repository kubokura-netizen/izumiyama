# ============================================================
#  配布用ランタイム(_runtime)作成スクリプト  ※メンテ担当者のみ使用
#
#  何をするか:
#    Python 本体(Embeddable版)と必要ライブラリを、ツールフォルダ内の
#    _runtime\python\ に丸ごと入れます。これにより配布先のPCでは
#    Python のインストールが一切不要になります。
#
#  使い方:
#    PowerShell で 98_dashboard\配布ランタイム作成.ps1 を実行
#    （ネット接続が必要・5分程度・作成後は _runtime を含めてフォルダごと配布）
#
#  なぜ venv を配らないのか:
#    venv は Python 本体を含まず、作成時の絶対パス(pyvenv.cfg の home=)を
#    参照する。ユーザー名が違う他人のPCでは必ず起動に失敗する。
# ============================================================
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# app.py が要求する Python。変更する場合は動作確認のうえで。
$PyVersion = '3.12.10'
$PyTag     = 'python312'

$DashDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root    = Split-Path -Parent $DashDir
$Runtime = Join-Path $Root '_runtime'
$PyDir   = Join-Path $Runtime 'python'
$Work    = Join-Path $Runtime '_tmp'

Write-Host ''
Write-Host '=== 配布用ランタイムを作成します ===' -ForegroundColor Cyan
Write-Host "  出力先: $PyDir"
Write-Host ''

if (Test-Path $Runtime) {
    Write-Host '既存の _runtime を削除しています...'
    Remove-Item -Recurse -Force $Runtime
}
New-Item -ItemType Directory -Force -Path $PyDir | Out-Null
New-Item -ItemType Directory -Force -Path $Work  | Out-Null

# --- 1. Embeddable Python を取得・展開 ---------------------------------
$EmbedZip = Join-Path $Work 'python-embed.zip'
Write-Host "[1/4] Python $PyVersion (Embeddable) をダウンロード中..."
Invoke-WebRequest -Uri "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-embed-amd64.zip" `
                  -OutFile $EmbedZip -TimeoutSec 300
Expand-Archive -Path $EmbedZip -DestinationPath $PyDir -Force

# --- 2. ._pth を書き換え（同梱方式の肝） -------------------------------
# Embeddable版は ._pth が sys.path を固定するため、次の2点の対応が要る。
#   (1) 既定で `import site` が無効 ＝ pip で入れたライブラリを読み込めない
#   (2) 実行するスクリプトのフォルダが sys.path に入らない
#       → app.py の `import index_html` や、transcribe.py の
#         `from transcribe_docs import run_docs` が ModuleNotFoundError になる
# そこで 98_dashboard と 99_data\src を明示的に登録する。
# ._pth 内の相対パスは「この ._pth があるフォルダ」基準なので、
#   _runtime\python\ から見て ..\.. がツールのルート。
# ※ フォルダ構成を変えた場合はここも直すこと。
Write-Host '[2/4] sys.path を設定しています (site-packages / 98_dashboard / 99_data\src)...'
$PthFile = Join-Path $PyDir "$PyTag._pth"
if (-not (Test-Path $PthFile)) { throw "._pth が見つかりません: $PthFile" }
@"
$PyTag.zip
.
Lib\site-packages
..\..\98_dashboard
..\..\99_data\src
import site
"@ | Set-Content -Path $PthFile -Encoding ascii

# --- 3. pip を導入 -----------------------------------------------------
Write-Host '[3/4] pip を導入しています...'
$GetPip = Join-Path $Work 'get-pip.py'
Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $GetPip -TimeoutSec 300
$PyExe = Join-Path $PyDir 'python.exe'
& $PyExe $GetPip --no-warn-script-location
if ($LASTEXITCODE -ne 0) { throw 'pip の導入に失敗しました' }

# --- 4. 必要ライブラリを導入 -------------------------------------------
Write-Host '[4/4] 必要ライブラリ (flask / openpyxl / python-docx) を導入しています...'
$Req = Join-Path $DashDir 'requirements.txt'
& $PyExe -m pip install --no-warn-script-location -r $Req
if ($LASTEXITCODE -ne 0) { throw 'ライブラリの導入に失敗しました' }

Remove-Item -Recurse -Force $Work

# --- 検証: 配布先で本当に動くかを確認 -----------------------------------
# 外部ライブラリだけでなく、ツール自身のモジュール(index_html /
# transcribe_docs)まで読めることを必ず確認する。._pth の設定漏れは
# ここでしか気付けない。
Write-Host ''
Write-Host '=== 検証中 ===' -ForegroundColor Cyan
& $PyExe -c @"
import importlib
for m in ['flask', 'openpyxl', 'docx', 'index_html', 'transcribe', 'transcribe_docs']:
    importlib.import_module(m)
    print('  import OK :', m)
"@
if ($LASTEXITCODE -ne 0) { throw '検証に失敗しました（モジュールを読み込めません）' }

$size = [math]::Round(((Get-ChildItem -Recurse $Runtime | Measure-Object -Property Length -Sum).Sum / 1MB), 1)
Write-Host ''
Write-Host "完了しました。_runtime のサイズ: ${size}MB" -ForegroundColor Green
Write-Host ''
Write-Host '配布方法: ツールのフォルダごと(_runtime を含めて)ZIPにして渡してください。'
Write-Host '受け取った人は「ダッシュボード起動.bat」をダブルクリックするだけです。'
Write-Host ''

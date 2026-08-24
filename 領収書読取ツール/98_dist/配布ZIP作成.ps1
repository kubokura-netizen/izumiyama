# ============================================================
#  領収書読取ツール 配布ZIP作成  ※メンテ担当者のみ
#
#  そのまま人に渡せるZIPをデスクトップに作る。
#   - _runtime(Python同梱) と _ocr(Tesseract同梱) を含める → 配布先はPython/管理者不要
#   - 入出力データ(01_input/02_output/04_work の中身)は除外
#   - .venv / .git / __pycache__ / 98_dist(メンテ用) は除外
#   - Ollama本体とモデルは含めない（サイズ大。配布先で「Ollama準備.bat」で導入）
#
#  先に 98_dist\配布ランタイム作成.ps1 を実行して _runtime を作っておくこと。
# ============================================================
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$DistDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root    = Split-Path -Parent $DistDir
$ToolName = Split-Path -Leaf $Root

if (-not (Test-Path (Join-Path $Root '_runtime\python\python.exe'))) {
    Write-Host 'エラー: _runtime がありません。先に 配布ランタイム作成.ps1 を実行してください。' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path (Join-Path $Root '_ocr\tesseract\tesseract.exe'))) {
    Write-Host 'エラー: _ocr(Tesseract) がありません。' -ForegroundColor Red
    exit 1
}

$EmptyDirs    = @('01_input', '02_output', '04_work')      # 中身は入れずフォルダだけ作る
$ExcludeNames = @('.venv', '.git', '.claude', '__pycache__', '98_dist')

$Stage = Join-Path $env:SystemDrive ('\_rcptdist_' + [System.Guid]::NewGuid().ToString('N').Substring(0, 6))
$StageTool = Join-Path $Stage $ToolName
New-Item -ItemType Directory -Force -Path $StageTool | Out-Null

Write-Host ''
Write-Host '=== 配布ZIPを作成します ===' -ForegroundColor Cyan
Write-Host '[1/3] ファイルをコピー...'
Get-ChildItem -LiteralPath $Root -Force | Where-Object { $ExcludeNames -notcontains $_.Name } | ForEach-Object {
    $isEmpty = $false
    foreach ($e in $EmptyDirs) { if ($_.Name -eq $e) { $isEmpty = $true } }
    if ($isEmpty) {
        New-Item -ItemType Directory -Force -Path (Join-Path $StageTool $_.Name) | Out-Null
    }
    else {
        Copy-Item -LiteralPath $_.FullName -Destination $StageTool -Recurse -Force
    }
}

Write-Host '[2/3] 不要ファイルを除去...'
Get-ChildItem -LiteralPath $StageTool -Recurse -Force -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -eq '__pycache__' } |
    ForEach-Object { Remove-Item -Recurse -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
Get-ChildItem -LiteralPath $StageTool -Recurse -Force -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like '~$*' } |
    ForEach-Object { Remove-Item -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue }

Write-Host '[3/3] ZIP化...'
$Stamp = Get-Date -Format 'yyyyMMdd'
$ZipPath = Join-Path ([Environment]::GetFolderPath('Desktop')) ("{0}_配布用_{1}.zip" -f $ToolName, $Stamp)
if (Test-Path $ZipPath) { Remove-Item -Force -LiteralPath $ZipPath }
Compress-Archive -Path $StageTool -DestinationPath $ZipPath -CompressionLevel Optimal
Remove-Item -Recurse -Force -LiteralPath $Stage

$size = [math]::Round((Get-Item $ZipPath).Length / 1MB, 1)
Write-Host ''
Write-Host "完了 (${size}MB)" -ForegroundColor Green
Write-Host "  $ZipPath"
Write-Host ''
Write-Host '渡すときに伝えること:'
Write-Host '  1. ZIPを右クリック→すべて展開（ZIPのまま実行しない）'
Write-Host '  2. 高精度にするなら「Ollama準備.bat」を一度実行（約3GB DL・任意）'
Write-Host '  3. PDFを 01_input に入れ、「読取実行.bat」をダブルクリック'


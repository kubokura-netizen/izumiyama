# ============================================================
#  配布ZIP作成スクリプト  ※メンテ担当者のみ使用
#
#  何をするか:
#    そのまま人に渡せるZIPをデスクトップに作ります。
#      - _runtime（Python同梱）を含める → 配布先は Python 不要
#      - 患者情報(01_input / 02_output / 03_logs の中身)は除外する
#      - 配布に使えない .venv_dashboard / .venv_dashboard.zip は除外する
#
#  先に 配布ランタイム作成.ps1 を実行して _runtime を作っておくこと。
# ============================================================
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$DashDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root    = Split-Path -Parent $DashDir
$ToolName = Split-Path -Leaf $Root

if (-not (Test-Path (Join-Path $Root '_runtime\python\python.exe'))) {
    Write-Host ''
    Write-Host 'エラー: _runtime が見つかりません。' -ForegroundColor Red
    Write-Host '先に 98_dashboard\配布ランタイム作成.ps1 を実行してください。'
    Write-Host '(_runtime が無いZIPを配ると、配布先で起動できません)'
    Write-Host ''
    exit 1
}

# 中身を配布しないフォルダ（フォルダ自体は空で残す＝ツールが参照するため）
$EmptyDirs = @('01_input', '02_output', '03_logs')
# まるごと除外
$ExcludeNames = @('.venv_dashboard', '.venv_dashboard.zip', '.git', '.claude', '__pycache__', '_web_profile')

# 作業フォルダはドライブ直下の短いパスにする。テンプレートに長いファイル名
# （論文PDF等）があり、%TEMP% 配下だと Windows の260文字制限に当たるため。
$Stage = Join-Path $env:SystemDrive ("\_prpdist_" + [System.Guid]::NewGuid().ToString('N').Substring(0,6))
$StageTool = Join-Path $Stage $ToolName
New-Item -ItemType Directory -Force -Path $StageTool | Out-Null

Write-Host ''
Write-Host '=== 配布ZIPを作成します ===' -ForegroundColor Cyan
Write-Host '[1/3] ファイルをコピーしています...'

Get-ChildItem -LiteralPath $Root | Where-Object { $ExcludeNames -notcontains $_.Name } | ForEach-Object {
    $isEmptyDir = $false
    foreach ($e in $EmptyDirs) { if ($_.Name.StartsWith($e)) { $isEmptyDir = $true } }
    if ($isEmptyDir) {
        # フォルダ名だけ作る（中身＝患者情報は入れない）
        New-Item -ItemType Directory -Force -Path (Join-Path $StageTool $_.Name) | Out-Null
    } else {
        Copy-Item -LiteralPath $_.FullName -Destination $StageTool -Recurse -Force
    }
}

# コピー後に紛れ込んだキャッシュ類を掃除
Write-Host '[2/3] 不要ファイルを除去しています...'
Get-ChildItem -LiteralPath $StageTool -Recurse -Force -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -eq '__pycache__' -or $_.Name -eq '_web_profile' } |
    ForEach-Object { Remove-Item -Recurse -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
# Excel の一時ファイル(~$...)も除外
Get-ChildItem -LiteralPath $StageTool -Recurse -Force -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like '~$*' } |
    ForEach-Object { Remove-Item -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue }

# --- 患者情報が残っていないか最終確認 ---
$leak = Get-ChildItem -LiteralPath $StageTool -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '\\(01_input|02_output|03_logs)' }
if ($leak) {
    Write-Host ''
    Write-Host 'エラー: 入出力フォルダにファイルが残っています（中止しました）:' -ForegroundColor Red
    $leak | ForEach-Object { Write-Host ('  ' + $_.FullName.Replace($StageTool, '')) }
    Remove-Item -Recurse -Force -LiteralPath $Stage
    exit 1
}

Write-Host '[3/3] ZIPにまとめています...'
$Stamp = Get-Date -Format 'yyyyMMdd'
$ZipPath = Join-Path ([Environment]::GetFolderPath('Desktop')) ("{0}_配布用_{1}.zip" -f $ToolName, $Stamp)
if (Test-Path $ZipPath) { Remove-Item -Force -LiteralPath $ZipPath }
Compress-Archive -Path $StageTool -DestinationPath $ZipPath -CompressionLevel Optimal
Remove-Item -Recurse -Force -LiteralPath $Stage

$size = [math]::Round((Get-Item $ZipPath).Length / 1MB, 1)
Write-Host ''
Write-Host "完了しました (${size}MB)" -ForegroundColor Green
Write-Host "  $ZipPath"
Write-Host ''
Write-Host '渡すときに伝えること:'
Write-Host '  1. ZIPを右クリック →「すべて展開」で展開する（ZIPのまま実行しない）'
Write-Host '  2. 中の「ダッシュボード起動.bat」をダブルクリックする'
Write-Host '  3. 初回に「WindowsによってPCが保護されました」と出たら'
Write-Host '     「詳細情報」→「実行」を押す'
Write-Host ''

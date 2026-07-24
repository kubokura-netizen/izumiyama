# ============================================================
#  PRP申請書類 自動転記ツール  ワンファイル導入／更新スクリプト
#  install.ps1
#
#  これ1本で、何も無いPCにツール一式を構築できます:
#    1) GitHub(main) から最新コードを取得
#    2) Python同梱ランタイム(_runtime)を作成（配布ランタイム作成.ps1 を利用）
#    3) Web転記/PDF化の部品(playwright/pywin32/PyMuPDF) と Chromium を導入
#  既に導入済みのフォルダがある場合は「更新」として動作し、
#  患者情報(01_input/02_output/03_logs)やログイン情報はそのまま保持します。
#
#  クライアントへ配布するのは bootstrap\PRPツール_セットアップ.bat だけ。
#  （この install.ps1 は .bat が GitHub から取得して実行します）
#
#  ※ 送信・申請などのフォーム操作は一切しません。導入・更新のみ。
# ============================================================
param(
    [string]$TargetDir = '',
    [switch]$SkipBrowser   # Chromium導入を省く（Web転記を使わない/検証用）
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$Owner   = 'kubokura-netizen'
$Repo    = 'izumiyama'
$Branch  = 'main'
$ZipUrl  = "https://codeload.github.com/$Owner/$Repo/zip/refs/heads/$Branch"
$ApiUrl  = "https://api.github.com/repos/$Owner/$Repo/commits/$Branch"
$ToolName = 'PRP申請書類 自動転記ツール'

function Info($m) { Write-Host $m -ForegroundColor Cyan }
function Ok($m)   { Write-Host $m -ForegroundColor Green }
function Warn($m) { Write-Host $m -ForegroundColor Yellow }

$work = $null
try {
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (-not $TargetDir) {
        $TargetDir = Join-Path ([Environment]::GetFolderPath('Desktop')) $ToolName
    }
    $isUpdate = Test-Path (Join-Path $TargetDir '98_dashboard\app.py')

    Write-Host ''
    Info '============================================================'
    if ($isUpdate) { Info '  PRP申請書類 自動転記ツール を最新版に更新します' }
    else           { Info '  PRP申請書類 自動転記ツール を導入します（初回）' }
    Info '============================================================'
    Write-Host "  場所 : $TargetDir"
    if (-not $isUpdate) {
        Write-Host ''
        Write-Host '  ※ 初回は Python・ブラウザ部品のダウンロードで 10〜15分ほど' -ForegroundColor DarkGray
        Write-Host '     かかります（ネット接続が必要）。そのままお待ちください。' -ForegroundColor DarkGray
    }
    Write-Host ''

    # 作業フォルダ（ドライブ直下の短いパス：日本語長名＋260字制限対策）
    $work = Join-Path $env:SystemDrive ('\_prpinst_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
    New-Item -ItemType Directory -Force -Path $work | Out-Null

    # --- 1. 最新コードを取得・展開 ---------------------------------------
    Info '[1/4] 最新のコードをダウンロードしています...'
    $zip = Join-Path $work 'code.zip'
    Invoke-WebRequest -Uri $ZipUrl -OutFile $zip -UseBasicParsing -TimeoutSec 300
    $ext = Join-Path $work 'x'; New-Item -ItemType Directory -Force -Path $ext | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $ext, [System.Text.Encoding]::UTF8)
    $marker = Get-ChildItem -LiteralPath $ext -Recurse -Filter 'app.py' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.DirectoryName -match '98_dashboard$' } | Select-Object -First 1
    if (-not $marker) { throw '取得したデータにツール本体が見つかりませんでした。' }
    $srcTool = Split-Path -Parent (Split-Path -Parent $marker.FullName)

    # --- 2. コードを配置（初回=丸ごと / 更新=データ保持で上書き）---------
    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    if ($isUpdate) {
        Info '[2/4] コードを最新化しています（患者情報・実行環境は保持）...'
        $keep = @(
            '01_input【ヒアリングシートをここへ】', '02_output【転記済みファイルがここに生成】', '03_logs',
            '_runtime', '.venv_dashboard', '_web_profile', '.git', '.claude', '__pycache__'
        )
        $ra = @($srcTool, $TargetDir, '/E', '/R:2', '/W:2', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
        foreach ($d in $keep) { $ra += '/XD'; $ra += (Join-Path $srcTool $d); $ra += (Join-Path $TargetDir $d) }
        & robocopy @ra | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "コピーに失敗しました (robocopy=$LASTEXITCODE)" }
        Get-ChildItem -LiteralPath $TargetDir -Recurse -Force -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq '__pycache__' } |
            ForEach-Object { Remove-Item -Recurse -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
    }
    else {
        Info '[2/4] ファイルを配置しています...'
        $ra = @($srcTool, $TargetDir, '/E', '/R:2', '/W:2', '/NFL', '/NDL', '/NJH', '/NJS', '/NP',
                '/XD', (Join-Path $srcTool '.git'), (Join-Path $srcTool '.claude'))
        & robocopy @ra | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "コピーに失敗しました (robocopy=$LASTEXITCODE)" }
    }

    # --- 3. Python同梱ランタイム（_runtime）を用意 -----------------------
    $pyExe = Join-Path $TargetDir '_runtime\python\python.exe'
    if (-not (Test-Path $pyExe)) {
        Info '[3/4] Python同梱ランタイムを作成しています（初回のみ・数分）...'
        $mk = Join-Path $TargetDir '98_dashboard\配布ランタイム作成.ps1'
        if (-not (Test-Path $mk)) { throw '配布ランタイム作成.ps1 が見つかりません。' }
        & $mk
    }
    else {
        Info '[3/4] Python同梱ランタイムは既にあります（そのまま使用）。'
    }
    if (-not (Test-Path $pyExe)) { throw 'ランタイムの作成に失敗しました（_runtime が作られませんでした）。' }

    # --- 4. Web転記/PDF化の部品 + ブラウザ ------------------------------
    Info '[4/4] Web転記・PDF化の部品を導入しています...'
    & $pyExe -m pip install --no-warn-script-location playwright pywin32 PyMuPDF 2>&1 | Out-Null
    if ($SkipBrowser) {
        Warn '  ブラウザ(Chromium)の導入はスキップしました（-SkipBrowser）。'
    }
    else {
        Write-Host '  ブラウザ(Chromium)を導入しています（初回は数分）...'
        & $pyExe -m playwright install chromium 2>&1 | Out-Null
    }

    # デスクトップに起動ショートカットを作成（見つけやすく）
    try {
        $lnkTarget = Join-Path $TargetDir 'ダッシュボード起動.bat'
        if (Test-Path $lnkTarget) {
            $desktop = [Environment]::GetFolderPath('Desktop')
            $lnk = Join-Path $desktop 'PRPツールを起動.lnk'
            $ws = New-Object -ComObject WScript.Shell
            $sc = $ws.CreateShortcut($lnk)
            $sc.TargetPath = $lnkTarget
            $sc.WorkingDirectory = $TargetDir
            $sc.Description = 'PRP申請書類 自動転記ツール を起動'
            $sc.Save()
        }
    }
    catch {}

    # 反映バージョン（確認用・失敗しても無視）
    $verNote = ''
    try {
        $c = Invoke-RestMethod -Uri $ApiUrl -Headers @{ 'User-Agent' = 'prp-installer' } -TimeoutSec 20
        $verNote = '  反映バージョン: ' + $c.sha.Substring(0, 7) +
                   ' (' + ([datetime]$c.commit.committer.date).ToLocalTime().ToString('yyyy-MM-dd HH:mm') + ')'
    }
    catch {}

    Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue
    $work = $null

    Write-Host ''
    if ($isUpdate) { Ok '✔ 最新版に更新しました。' } else { Ok '✔ 導入が完了しました。' }
    if ($verNote) { Write-Host $verNote }
    Write-Host ''
    Write-Host '  使い方: デスクトップの「PRPツールを起動」（または' -NoNewline
    Write-Host " $TargetDir の「ダッシュボード起動.bat」）"
    Write-Host '          をダブルクリックしてください。'
    Write-Host ''
    try { Start-Process explorer.exe $TargetDir } catch {}
}
catch {
    Write-Host ''
    Warn ('失敗しました: ' + $_.Exception.Message)
    Write-Host '  ・インターネット接続をご確認のうえ、もう一度お試しください。'
    Write-Host '  ・会社等でダウンロード制限がある場合は解除が必要なことがあります。'
    Write-Host '  ・解決しない場合は、この画面を開発担当（大野／窪倉）へお知らせください。'
    Write-Host ''
    if ($work) { Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue }
}
Read-Host '終了するには Enter キーを押してください'


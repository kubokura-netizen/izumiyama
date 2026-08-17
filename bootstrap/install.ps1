# ============================================================
#  ワンファイル導入／更新スクリプト  install.ps1
#
#  対応ツール（-Tool で選択）:
#    prp     … PRP申請書類 自動転記ツール（既定）
#    receipt … 領収書読取ツール（OCR＋任意のローカルLLM）
#
#  これ1本で、何も無いPCにツール一式を構築できます:
#    1) GitHub(main) から最新コードを取得
#    2) Python同梱ランタイム(_runtime)を作成（各ツールの配布ランタイム作成.ps1）
#    3) 部品を導入
#         prp     : playwright/pywin32/PyMuPDF ＋ Chromium
#         receipt : OCR本体(Tesseract, GitHub Releaseから) ＋ pip各種
#  既に導入済みのフォルダがある場合は「更新」として動作し、
#  入出力データ(01_input/02_output 等)・実行環境(_runtime/_ocr)は保持します。
#
#  クライアントへ配布するのは bootstrap\*_セットアップ.bat だけ。
#  （この install.ps1 は .bat が GitHub から取得して実行します）
#
#  ※ 送信・申請などのフォーム操作は一切しません。導入・更新のみ。
# ============================================================
param(
    [string]$TargetDir = '',
    [ValidateSet('prp', 'receipt')]
    [string]$Tool = 'prp',
    [switch]$SkipBrowser   # prp: Chromium導入を省く（検証用）
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$Owner   = 'kubokura-netizen'
$Repo    = 'izumiyama'
$Branch  = 'main'
$ZipUrl  = "https://codeload.github.com/$Owner/$Repo/zip/refs/heads/$Branch"
$ApiUrl  = "https://api.github.com/repos/$Owner/$Repo/commits/$Branch"
# 領収書ツールのOCR本体(Tesseract)。GitHub Release資産から取得（手動DL不要）。
$OcrUrl  = "https://github.com/$Owner/$Repo/releases/download/ocr-assets/receipt_ocr.zip"

# --- ツール別の設定 -------------------------------------------------
if ($Tool -eq 'receipt') {
    $cfg = @{
        Name         = '領収書読取ツール'
        MarkerFile   = 'receipt_ocr.py'
        MarkerDirRx  = 'src$'
        MarkerRel    = 'src\receipt_ocr.py'
        RuntimeMaker = '98_dist\配布ランタイム作成.ps1'
        Keep         = @('01_input', '02_output', '03_work',
                         '_runtime', '_ocr', '.venv', '.git', '.claude', '__pycache__')
        Launcher     = '読取実行.bat'
        Shortcut     = '領収書読取ツールを起動'
        # クライアントの目に触れないよう隠す開発者用（削除はしない）
        Hide         = @('初回準備.bat', '98_dist')
    }
}
else {
    $cfg = @{
        Name         = 'PRP申請書類 自動転記ツール'
        MarkerFile   = 'app.py'
        MarkerDirRx  = '98_dashboard$'
        MarkerRel    = '98_dashboard\app.py'
        RuntimeMaker = '98_dashboard\配布ランタイム作成.ps1'
        Keep         = @('01_input【ヒアリングシートをここへ】',
                         '02_output【転記済みファイルがここに生成】', '03_logs',
                         '99_data\テンプレート',
                         '_runtime', '.venv_dashboard', '_web_profile', '.git', '.claude', '__pycache__')
        Launcher     = 'ダッシュボード起動.bat'
        Shortcut     = 'PRPツールを起動'
        # 開発者用は隠す。ただし 99_data\テンプレート（クライアント編集）は残す。
        # 00_マニュアル/01_input/02_output/03_logs/各.bat も表示のまま。
        Hide         = @('98_dashboard', '99_data\src', '99_data\マッピング',
                         '99_data\パッチノート', '99_data\参考元【原本・設計書】',
                         '99_data\設計書', '99_data\_web_profile')
    }
}

function Info($m) { Write-Host $m -ForegroundColor Cyan }
function Ok($m)   { Write-Host $m -ForegroundColor Green }
function Warn($m) { Write-Host $m -ForegroundColor Yellow }

# セットアップの詳細ログ（失敗原因を後から確認できるように、全出力を残す）
$script:LogFile = $null
function Log($m) {
    if (-not $script:LogFile) { return }
    try { Add-Content -LiteralPath $script:LogFile -Value ([string]$m) -Encoding UTF8 } catch {}
}

# ネイティブ実行（pip等）の安全な呼び出し。
#   ・全出力（stdout/stderr）をログに残す
#   ・pip等がstderrに“注意書き”を出しただけで失敗扱いになるPS5.1の罠を回避
#     （$ErrorActionPreferenceを一時的にContinueにして 2>&1 を安全に扱う）
#   ・成否は必ず終了コード($LASTEXITCODE)で判定する
function Invoke-Step($exe, $argList, $what) {
    Write-Host ("  " + $what + " ...")
    Log ''
    Log ('### ' + $what)
    Log ('# ' + $exe + ' ' + ($argList -join ' '))
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $exe @argList 2>&1 | ForEach-Object { Log $_ }
    }
    finally {
        $ErrorActionPreference = $prev
    }
    if ($LASTEXITCODE -ne 0) {
        throw ($what + ' に失敗しました (終了コード=' + $LASTEXITCODE + ')')
    }
}

# 開発者専用ファイルをクライアントの目に触れないよう「隠し属性」にする（削除はしない）。
function Hide-Item($path) {
    try {
        $it = Get-Item -LiteralPath $path -Force -ErrorAction Stop
        $it.Attributes = $it.Attributes -bor [System.IO.FileAttributes]::Hidden
    }
    catch {}
}

$work = $null
try {
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # 既定は「今いるフォルダ（＝セットアップ.bat を置いたフォルダ）」に導入する。
    if (-not $TargetDir) { $TargetDir = (Get-Location).Path }
    $isUpdate = Test-Path (Join-Path $TargetDir $cfg.MarkerRel)

    # 詳細ログの出力先（TargetDir直下・保持対象。失敗時にこの場所を案内する）
    $script:LogFile = Join-Path $TargetDir '_setup_last.log'
    try {
        New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
        Set-Content -LiteralPath $script:LogFile -Value ('setup log [' + $Tool + ']  ' +
            (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    }
    catch {}

    Write-Host ''
    Info '============================================================'
    if ($isUpdate) { Info ('  ' + $cfg.Name + ' を最新版に更新します') }
    else           { Info ('  ' + $cfg.Name + ' を導入します（初回）') }
    Info '============================================================'
    Write-Host "  場所 : $TargetDir"
    if (-not $isUpdate) {
        Write-Host ''
        Write-Host '  ※ 初回は Python・部品のダウンロードで 10〜15分ほど' -ForegroundColor DarkGray
        Write-Host '     かかります（ネット接続が必要）。そのままお待ちください。' -ForegroundColor DarkGray
    }
    Write-Host ''

    # 作業フォルダ（ドライブ直下の短いパス：日本語長名＋260字制限対策）
    $work = Join-Path $env:SystemDrive ('\_inst_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
    New-Item -ItemType Directory -Force -Path $work | Out-Null

    # --- 1. 最新コードを取得・展開 ---------------------------------------
    Info '[1/4] 最新のコードをダウンロードしています...'
    $zip = Join-Path $work 'code.zip'
    Invoke-WebRequest -Uri $ZipUrl -OutFile $zip -UseBasicParsing -TimeoutSec 300
    $ext = Join-Path $work 'x'; New-Item -ItemType Directory -Force -Path $ext | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $ext, [System.Text.Encoding]::UTF8)
    $marker = Get-ChildItem -LiteralPath $ext -Recurse -Filter $cfg.MarkerFile -File -ErrorAction SilentlyContinue |
        Where-Object { $_.DirectoryName -match $cfg.MarkerDirRx } | Select-Object -First 1
    if (-not $marker) { throw '取得したデータにツール本体が見つかりませんでした。' }
    $srcTool = Split-Path -Parent (Split-Path -Parent $marker.FullName)

    # --- 2. コードを配置（初回=丸ごと / 更新=データ保持で上書き）---------
    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    if ($isUpdate) {
        Info '[2/4] コードを最新化しています（データ・実行環境は保持）...'
        $ra = @($srcTool, $TargetDir, '/E', '/R:2', '/W:2', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
        foreach ($d in $cfg.Keep) { $ra += '/XD'; $ra += (Join-Path $srcTool $d); $ra += (Join-Path $TargetDir $d) }
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
        $mk = Join-Path $TargetDir $cfg.RuntimeMaker
        if (-not (Test-Path $mk)) { throw '配布ランタイム作成.ps1 が見つかりません。' }
        & $mk
    }
    else {
        Info '[3/4] Python同梱ランタイムは既にあります（そのまま使用）。'
    }
    if (-not (Test-Path $pyExe)) { throw 'ランタイムの作成に失敗しました（_runtime が作られませんでした）。' }

    # --- 4. 部品（ツール別） --------------------------------------------
    if ($Tool -eq 'receipt') {
        Info '[4/4] OCR部品(Tesseract)と読取ライブラリを用意しています...'
        # 4-1) Tesseract本体を GitHub Release から取得・展開（無ければ）
        $tessExe = Join-Path $TargetDir '_ocr\tesseract\tesseract.exe'
        if (-not (Test-Path $tessExe)) {
            Write-Host '  OCR本体(Tesseract, 約90MB)をダウンロードしています...'
            $ocrZip = Join-Path $work 'ocr.zip'
            Invoke-WebRequest -Uri $OcrUrl -OutFile $ocrZip -UseBasicParsing -TimeoutSec 600
            $ocrTmp = Join-Path $work 'ocrx'
            [System.IO.Compression.ZipFile]::ExtractToDirectory($ocrZip, $ocrTmp, [System.Text.Encoding]::UTF8)
            $ra2 = @((Join-Path $ocrTmp 'tesseract'), (Join-Path $TargetDir '_ocr\tesseract'),
                     '/E', '/R:2', '/W:2', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
            & robocopy @ra2 | Out-Null
            if ($LASTEXITCODE -ge 8) { throw "OCR本体の展開に失敗しました (robocopy=$LASTEXITCODE)" }
        }
        else {
            Write-Host '  Tesseractは既にあります（そのまま使用）。'
        }
        if (-not (Test-Path $tessExe)) { throw 'Tesseractの用意に失敗しました。' }
        # 4-2) 読取ライブラリ（更新時など runtime を作り直さない場合の保険）
        Invoke-Step $pyExe @('-m', 'pip', 'install', '--no-warn-script-location',
            '--disable-pip-version-check', '--no-input', '--retries', '3', '--timeout', '60',
            'pymupdf', 'opencv-python-headless', 'pytesseract', 'Pillow', 'openpyxl', 'numpy') `
            'PyPIから読取ライブラリを取得'
    }
    else {
        Info '[4/4] Web転記・PDF化の部品を導入しています...'
        $pipArgs = @('-m', 'pip', 'install', '--no-warn-script-location',
                     '--disable-pip-version-check', '--no-input',
                     '--retries', '3', '--timeout', '60',
                     'playwright', 'pywin32', 'PyMuPDF')
        Invoke-Step $pyExe $pipArgs 'PyPIから部品(playwright/pywin32/PyMuPDF)を取得'
        if ($SkipBrowser) {
            Warn '  ブラウザ(Chromium)の導入はスキップしました（-SkipBrowser）。'
        }
        else {
            Invoke-Step $pyExe @('-m', 'playwright', 'install', 'chromium') 'ブラウザ(Chromium)を取得（初回は数分）'
        }
    }

    # クライアントの見た目をすっきりさせる：開発者専用ファイル/フォルダを隠し属性に
    # （削除はしない。$cfg.Hide で定義。テンプレート等の編集資材は対象外）。
    foreach ($h in $cfg.Hide) {
        Hide-Item (Join-Path $TargetDir $h)
    }

    # デスクトップに起動ショートカットを作成（見つけやすく）
    try {
        $lnkTarget = Join-Path $TargetDir $cfg.Launcher
        if (Test-Path $lnkTarget) {
            $desktop = [Environment]::GetFolderPath('Desktop')
            $lnk = Join-Path $desktop ($cfg.Shortcut + '.lnk')
            $ws = New-Object -ComObject WScript.Shell
            $sc = $ws.CreateShortcut($lnk)
            $sc.TargetPath = $lnkTarget
            $sc.WorkingDirectory = $TargetDir
            $sc.Description = $cfg.Name + ' を起動'
            $sc.Save()
        }
    }
    catch {}

    # 反映バージョン（確認用・失敗しても無視）
    $verNote = ''
    try {
        $c = Invoke-RestMethod -Uri $ApiUrl -Headers @{ 'User-Agent' = 'gu-installer' } -TimeoutSec 20
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
    Write-Host ('  使い方: デスクトップの「' + $cfg.Shortcut + '」（または') -NoNewline
    Write-Host (" $TargetDir の「" + $cfg.Launcher + '」）')
    Write-Host '          をダブルクリックしてください。'
    if ($Tool -eq 'receipt') {
        Write-Host '  ※ 金額・相手先の精度を上げるなら Ollama準備.bat を1回実行（任意）。' -ForegroundColor DarkGray
    }
    Write-Host ''
    try { Start-Process explorer.exe $TargetDir } catch {}
}
catch {
    Write-Host ''
    Warn ('失敗しました: ' + $_.Exception.Message)
    Log ('!!! FAILED: ' + $_.Exception.Message)
    Log ($_.ScriptStackTrace)
    # 失敗直前の実出力（本当のエラー）を画面にも出す
    if ($script:LogFile -and (Test-Path $script:LogFile)) {
        Write-Host ''
        Write-Host '  ── エラーの詳細（ログ末尾） ─────────────────' -ForegroundColor DarkGray
        try {
            Get-Content -LiteralPath $script:LogFile -Tail 15 -ErrorAction SilentlyContinue |
                ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor DarkGray }
        }
        catch {}
        Write-Host '  ─────────────────────────────────────────' -ForegroundColor DarkGray
        Write-Host ('  詳しいログ: ' + $script:LogFile) -ForegroundColor Yellow
        Write-Host '  ↑このファイルを開発担当（大野／窪倉）へ送っていただくと原因が特定できます。'
    }
    Write-Host ''
    Write-Host '  よくある原因:'
    Write-Host '  ・会社のネットワーク制限/プロキシで PyPI・GitHub・配布元がブロックされている'
    Write-Host '  ・保存場所が Google ドライブ(マイドライブ) 等の同期フォルダ → 導入中にファイルがロックされる'
    Write-Host '    （その場合は C:\ の通常フォルダに置いて再実行すると解決することがあります）'
    Write-Host '  ・解決しない場合は、この画面と上記ログを開発担当（大野／窪倉）へお知らせください。'
    Write-Host ''
    if ($work) { Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue }
}
Read-Host '終了するには Enter キーを押してください'

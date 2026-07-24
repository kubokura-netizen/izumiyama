# ============================================================
#  最新版に更新（パッチャー）ワーカー  _update.ps1
#
#  役割:
#    GitHub(main) の最新コードを取得して、このツールを最新化する。
#    クライアントは「最新版に更新.bat」をダブルクリックするだけ。
#
#  保持するもの（絶対に上書き・削除しない）:
#    01_input / 02_output / 03_logs … 患者情報・生成物
#    _runtime … 同梱Python（配布物・GitHubには無い）
#    .venv_dashboard … 開発用の仮想環境
#    99_data\_web_profile … ブラウザのログイン情報
#
#  更新するもの:
#    98_dashboard / 99_data\src / 99_data\マッピング / テンプレート類 /
#    各種 .bat / マニュアル など（＝コードと設定）
#
#  設計メモ:
#    - リポジトリが public のため、認証トークンは不要。
#      codeload の main.zip は常に「最新の push」を反映する
#      ＝どちらが push しても、次に実行した時点で最新になる。
#    - この _update.ps1 自身は更新対象に含む（PowerShellは起動時に
#      全文を読み込むため、実行中に上書きされても安全）。
#    - 実行中の「最新版に更新.bat」だけは上書きしない（cmd破損防止）。
#  ※ 送信・申請などのフォーム操作は一切しない。ファイル更新のみ。
# ============================================================
param(
  [Parameter(Mandatory = $true)][string]$ToolDir
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# --- 取得元リポジトリ（変更する場合はここだけ） ---
$Owner  = 'kubokura-netizen'
$Repo   = 'izumiyama'
$Branch = 'main'
$ZipUrl = "https://codeload.github.com/$Owner/$Repo/zip/refs/heads/$Branch"
$ApiUrl = "https://api.github.com/repos/$Owner/$Repo/commits/$Branch"

function Info($m) { Write-Host $m -ForegroundColor Cyan }
function Ok($m)   { Write-Host $m -ForegroundColor Green }
function Warn($m) { Write-Host $m -ForegroundColor Yellow }

$work = $null
try {
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
    Write-Host ''
    Info '=== ツールを最新版に更新します ==='
    Write-Host "  取得元 : $Owner/$Repo ($Branch)"
    Write-Host "  更新先 : $ToolDir"
    Write-Host ''

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # 作業フォルダはドライブ直下の短いパス（日本語の長いファイル名＋260字制限対策）
    $work = Join-Path $env:SystemDrive ('\_prppatch_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    $zip = Join-Path $work 'latest.zip'

    Info '[1/4] 最新ファイルをダウンロードしています...'
    Invoke-WebRequest -Uri $ZipUrl -OutFile $zip -UseBasicParsing -TimeoutSec 300

    Info '[2/4] 展開しています...'
    $ext = Join-Path $work 'x'
    New-Item -ItemType Directory -Force -Path $ext | Out-Null
    # GitHubのzipはUTF-8名だが Expand-Archive(PS5.1) は文字化けするため .NET でUTF-8指定展開
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $ext, [System.Text.Encoding]::UTF8)

    # 展開物から「98_dashboard\app.py」を探し、その2つ上＝ツール本体フォルダを特定
    # （フォルダ名が日本語でも、実体パスで扱うので名前ゆらぎに強い）
    $marker = Get-ChildItem -LiteralPath $ext -Recurse -Filter 'app.py' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.DirectoryName -match '98_dashboard$' } | Select-Object -First 1
    if (-not $marker) { throw '展開物にツール本体（98_dashboard\app.py）が見つかりませんでした。' }
    $srcTool = Split-Path -Parent (Split-Path -Parent $marker.FullName)

    Info '[3/4] コードを最新化しています（患者情報・実行環境は保持）...'
    # 上書き・削除してはいけないフォルダ
    $keep = @(
        '01_input【ヒアリングシートをここへ】', '02_output【転記済みファイルがここに生成】', '03_logs',
        '_runtime', '.venv_dashboard', '_web_profile', '.git', '.claude', '__pycache__'
    )
    # robocopy は splatting で渡す（日本語・空白・角括弧の引用符問題を回避）
    $roboArgs = @($srcTool, $ToolDir, '/E', '/R:2', '/W:2', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
    foreach ($d in $keep) {
        $roboArgs += '/XD'; $roboArgs += (Join-Path $srcTool $d); $roboArgs += (Join-Path $ToolDir $d)
    }
    # 実行中の「最新版に更新.bat」は上書きしない（cmdが破損するため）
    $roboArgs += @('/XF', (Join-Path $srcTool '最新版に更新.bat'))
    & robocopy @roboArgs | Out-Null
    $rc = $LASTEXITCODE
    if ($rc -ge 8) { throw "ファイルのコピーに失敗しました (robocopy=$rc)" }

    # 消えた古いモジュールの残骸（.pyc）を掃除
    Get-ChildItem -LiteralPath $ToolDir -Recurse -Force -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq '__pycache__' } |
        ForEach-Object { Remove-Item -Recurse -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue }

    Info '[4/4] 必要な部品を確認しています（不足があれば自動導入）...'
    # 実行に使う Python（配布は _runtime、開発は .venv_dashboard）
    $py = $null
    $c1 = Join-Path $ToolDir '_runtime\python\python.exe'
    $c2 = Join-Path $ToolDir '.venv_dashboard\Scripts\python.exe'
    if (Test-Path $c1) { $py = $c1 } elseif (Test-Path $c2) { $py = $c2 }
    if ($py) {
        try {
            $req = Join-Path $ToolDir '98_dashboard\requirements.txt'
            if (Test-Path $req) { & $py -m pip install --no-warn-script-location -r $req 2>&1 | Out-Null }
            & $py -m pip install --no-warn-script-location playwright pywin32 PyMuPDF 2>&1 | Out-Null
            & $py -m playwright install chromium 2>&1 | Out-Null
            Ok '  部品の確認が完了しました。'
        }
        catch {
            Warn '  部品の自動導入に一部失敗しました。ダッシュボードの「初回起動準備」からも導入できます。'
        }
    }
    else {
        Warn '  実行環境が見つからないため部品確認はスキップ（ダッシュボード起動時に案内されます）。'
    }

    # 反映したバージョン（確認用・失敗しても無視）
    $verNote = ''
    try {
        $c = Invoke-RestMethod -Uri $ApiUrl -Headers @{ 'User-Agent' = 'prp-patcher' } -TimeoutSec 20
        $sha = $c.sha.Substring(0, 7)
        $dt  = ([datetime]$c.commit.committer.date).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        $verNote = "  反映バージョン: $sha ($dt)"
    }
    catch {}

    Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue
    $work = $null

    Write-Host ''
    Ok '✔ 最新版に更新しました。'
    if ($verNote) { Write-Host $verNote }
    Write-Host '  「ダッシュボード起動.bat」をダブルクリックしてご利用ください。'
    Write-Host ''
}
catch {
    Write-Host ''
    Warn ('更新に失敗しました: ' + $_.Exception.Message)
    Write-Host '  ・インターネット接続をご確認のうえ、もう一度お試しください。'
    Write-Host '  ・解決しない場合は、この画面を開発担当（大野／窪倉）へお知らせください。'
    Write-Host ''
    if ($work) { Remove-Item -Recurse -Force -LiteralPath $work -ErrorAction SilentlyContinue }
}
Read-Host '終了するには Enter キーを押してください'


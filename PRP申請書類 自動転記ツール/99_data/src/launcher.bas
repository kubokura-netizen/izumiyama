Attribute VB_Name = "launcher"
' ============================================================
'  PRP申請書類 自動転記ツール  VBAランチャー（任意）
'  使い方:
'   1) マクロ有効ブック(.xlsm)を作り、本モジュールをインポート
'   2) ボタンにマクロ「転記実行」を割り当て
'   3) 押すと 99_data\transcribe.py を起動し、02_output に出力
'  前提: PCに Python（py または python）がインストール済み
' ============================================================
Option Explicit

' ツールのルート（このブックを置いた場所）を返す
Private Function RootDir() As String
    RootDir = ThisWorkbook.Path
End Function

' 先頭が prefix で始まるサブフォルダを返す（【説明】付きでも拾う）
Private Function ResolveDir(base As String, prefix As String) As String
    Dim f As String
    ResolveDir = base & "\" & prefix
    If Dir(ResolveDir, vbDirectory) <> "" Then Exit Function
    f = Dir(base & "\*", vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then
            If Left$(f, Len(prefix)) = prefix Then
                ResolveDir = base & "\" & f
                Exit Function
            End If
        End If
        f = Dir
    Loop
End Function

Private Function PyCmd() As String
    PyCmd = "py"   ' 無ければ "python" に変更
End Function

Public Sub 転記実行()
    Dim root As String, script As String, cmd As String
    Dim wsh As Object, rc As Long, outDir As String, logDir As String

    root = RootDir()
    script = root & "\99_data\src\transcribe.py"
    If Dir(script) = "" Then
        MsgBox "transcribe.py が見つかりません。" & vbCrLf & script, vbExclamation, "エラー"
        Exit Sub
    End If
    outDir = ResolveDir(root, "02_output")
    logDir = root & "\03_logs"

    cmd = "cmd /c " & PyCmd() & " """ & script & """"
    Set wsh = CreateObject("WScript.Shell")
    rc = wsh.Run(cmd, 1, True)   ' 完了まで待機・通常ウィンドウ

    If rc = 0 Then
        Dim latest As String
        latest = LatestFile(logDir, "*.xlsx")
        If MsgBox("転記が完了しました。" & vbCrLf & vbCrLf & _
                  "出力: " & outDir & vbCrLf & "ログ: " & latest & vbCrLf & vbCrLf & _
                  "ログを開きますか？", vbYesNo + vbInformation, "完了") = vbYes Then
            If latest <> "" Then Workbooks.Open latest
        End If
    Else
        MsgBox "転記処理でエラーが発生しました（終了コード " & rc & "）。", vbCritical, "エラー"
    End If
End Sub

Private Function LatestFile(folder As String, pattern As String) As String
    Dim f As String, newest As String, nt As Date, t As Date
    f = Dir(folder & "\" & pattern)
    Do While f <> ""
        t = FileDateTime(folder & "\" & f)
        If newest = "" Or t > nt Then newest = f: nt = t
        f = Dir
    Loop
    If newest <> "" Then LatestFile = folder & "\" & newest
End Function

Public Sub 出力フォルダを開く()
    Shell "explorer.exe """ & ResolveDir(RootDir(), "02_output") & """", vbNormalFocus
End Sub

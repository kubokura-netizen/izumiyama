# -*- coding: utf-8 -*-
"""
PRP申請書類 自動転記ツール ― ローカルWebダッシュボード

既存の 99_data/src/transcribe.py には一切手を加えず、その実行を
ブラウザ画面から操作できるようにする薄いラッパー（Flask）。

  - ヒアリングシート(.xlsx)をドラッグ&ドロップ → 01_input へ保存
  - 「転記実行」→ transcribe.py をサブプロセス実行し、進捗をライブ表示(SSE)
  - 実行後、03_logs の最新ログExcelを解析して 統計・要確認(黄色)項目 を表示
  - 02_output のフォルダをZIPでダウンロード
  - 過去の実行ログ履歴を一覧・再表示

起動:  python3 98_dashboard/app.py   （ブラウザで http://127.0.0.1:8765 ）
"""
import os
import io
import sys
import glob
import json
import zipfile
import datetime
import threading
import subprocess

from flask import (
    Flask, request, Response, jsonify,
    send_file, stream_with_context,
)
import openpyxl

# ---------------------------------------------------------------------------
# パス解決（transcribe.py と同じ流儀で、ツールのルートを基準にする）
#   このファイル: <ルート>/98_dashboard/app.py
# ---------------------------------------------------------------------------
DASH_DIR = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(DASH_DIR)                      # ツールのルート
DATA = os.path.join(BASE, "99_data")
TRANSCRIBE = os.path.join(DATA, "src", "transcribe.py")
WEB_FILL = os.path.join(DATA, "src", "web_fill.py")
WEB_MAPPING = os.path.join(DATA, "マッピング", "web_mapping.json")


def resolve_dir(prefix, create=False):
    """先頭が prefix に一致するルート直下フォルダを返す（【説明】付きでも拾う）。"""
    exact = os.path.join(BASE, prefix)
    if os.path.isdir(exact):
        return exact
    try:
        for name in sorted(os.listdir(BASE)):
            full = os.path.join(BASE, name)
            if os.path.isdir(full) and name.startswith(prefix):
                return full
    except OSError:
        pass
    if create:
        os.makedirs(exact, exist_ok=True)
    return exact


DIR_INPUT = resolve_dir("01_input", create=True)
DIR_OUTPUT = resolve_dir("02_output", create=True)
DIR_LOGS = resolve_dir("03_logs", create=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024   # 100MB まで


# ---------------------------------------------------------------------------
# ヘルパ
# ---------------------------------------------------------------------------
def safe_name(name):
    """パス区切りを除去した安全なファイル/フォルダ名にする。"""
    name = os.path.basename(name or "")
    return name.replace("\\", "").replace("/", "").strip()


def list_inputs():
    files = [f for f in glob.glob(os.path.join(DIR_INPUT, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")]
    files.sort(key=os.path.getmtime, reverse=True)
    return [{
        "name": os.path.basename(f),
        "mtime": datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M"),
        "size_kb": round(os.path.getsize(f) / 1024),
    } for f in files]


def list_outputs():
    """02_output 直下のフォルダを新しい順に。"""
    items = []
    try:
        entries = os.listdir(DIR_OUTPUT)
    except OSError:
        entries = []
    for name in entries:
        full = os.path.join(DIR_OUTPUT, name)
        if not os.path.isdir(full):
            continue
        nfiles = 0
        for _root, _dirs, fs in os.walk(full):
            nfiles += len([x for x in fs if not x.startswith("~$")])
        items.append({
            "name": name,
            "mtime_ts": os.path.getmtime(full),
            "mtime": datetime.datetime.fromtimestamp(os.path.getmtime(full)).strftime("%Y-%m-%d %H:%M"),
            "files": nfiles,
        })
    items.sort(key=lambda x: x["mtime_ts"], reverse=True)
    for it in items:
        del it["mtime_ts"]
    return items


def list_logs():
    files = glob.glob(os.path.join(DIR_LOGS, "転記ログ_*.xlsx"))
    files.sort(key=os.path.getmtime, reverse=True)
    return [{
        "name": os.path.basename(f),
        "mtime": datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M"),
    } for f in files]


def parse_log(log_path):
    """転記ログ_*.xlsx を解析して 統計・要確認行・実行情報 を返す。"""
    wb = openpyxl.load_workbook(log_path, read_only=True, data_only=True)
    result = {"stats": {}, "checks": [], "info": {}, "total": 0}

    if "転記ログ" in wb.sheetnames:
        ws = wb["転記ログ"]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if header:
            idx = {str(h): i for i, h in enumerate(header) if h is not None}

            def g(row, key):
                i = idx.get(key)
                return row[i] if (i is not None and i < len(row)) else None

            stats = {}
            for row in rows:
                if row is None or all(c is None for c in row):
                    continue
                result["total"] += 1
                status = str(g(row, "処理結果") or "")
                stats[status] = stats.get(status, 0) + 1
                if str(g(row, "確認") or "") == "要確認":
                    result["checks"].append({
                        "doc": str(g(row, "文書") or ""),
                        "sheet": str(g(row, "シート") or ""),
                        "cell": str(g(row, "セル") or ""),
                        "var": str(g(row, "変数") or ""),
                        "content": str(g(row, "転記内容") or ""),
                        "status": status,
                        "note": str(g(row, "備考") or ""),
                    })
            result["stats"] = stats

    if "実行情報" in wb.sheetnames:
        ws2 = wb["実行情報"]
        for row in ws2.iter_rows(values_only=True):
            if row and row[0]:
                result["info"][str(row[0])] = str(row[1]) if len(row) > 1 and row[1] is not None else ""
    wb.close()
    return result


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
@app.get("/api/state")
def api_state():
    return jsonify({
        "inputs": list_inputs(),
        "outputs": list_outputs(),
        "logs": list_logs(),
        "root": BASE,
    })


@app.post("/api/upload")
def api_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "ファイルがありません"}), 400
    name = safe_name(f.filename)
    if not name.lower().endswith(".xlsx"):
        return jsonify({"error": "拡張子が .xlsx のヒアリングシートを選んでください"}), 400
    dest = os.path.join(DIR_INPUT, name)
    f.save(dest)
    return jsonify({"name": name})


@app.get("/api/run")
def api_run():
    """指定ヒアリングシートで transcribe.py を実行し、進捗をSSEで流す。"""
    fname = safe_name(request.args.get("file", ""))
    hearing_path = os.path.join(DIR_INPUT, fname) if fname else ""

    @stream_with_context
    def generate():
        def sse(event, data):
            return "event: %s\ndata: %s\n\n" % (event, json.dumps(data, ensure_ascii=False))

        if not fname or not os.path.exists(hearing_path):
            yield sse("error", {"message": "ヒアリングシートが見つかりません（先にアップロードしてください）"})
            return

        yield sse("log", {"line": "▶ 実行開始: %s" % fname})

        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        env["PYTHONUNBUFFERED"] = "1"

        try:
            proc = subprocess.Popen(
                [sys.executable, TRANSCRIBE, hearing_path],
                cwd=BASE, env=env,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                bufsize=1, universal_newlines=True,
                encoding="utf-8", errors="replace",
            )
        except Exception as e:
            yield sse("error", {"message": "起動に失敗: %r" % e})
            return

        for line in iter(proc.stdout.readline, ""):
            line = line.rstrip("\n")
            if line != "":
                yield sse("log", {"line": line})
        proc.stdout.close()
        code = proc.wait()

        if code != 0:
            yield sse("error", {"message": "転記処理が異常終了しました（コード %s）" % code})
            return

        # 最新ログを解析して結果を返す
        logs = list_logs()
        result = {}
        log_name = ""
        if logs:
            log_name = logs[0]["name"]
            try:
                result = parse_log(os.path.join(DIR_LOGS, log_name))
            except Exception as e:
                yield sse("log", {"line": "（ログ解析でエラー: %r）" % e})

        yield sse("done", {
            "log": log_name,
            "result": result,
            "outputs": list_outputs(),
        })

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.get("/api/log")
def api_log():
    """過去ログの解析結果を返す。"""
    name = safe_name(request.args.get("name", ""))
    path = os.path.join(DIR_LOGS, name)
    if not name or not os.path.exists(path):
        return jsonify({"error": "ログが見つかりません"}), 404
    return jsonify(parse_log(path))


@app.get("/api/download-log")
def api_download_log():
    name = safe_name(request.args.get("name", ""))
    path = os.path.join(DIR_LOGS, name)
    if not name or not os.path.exists(path):
        return jsonify({"error": "ログが見つかりません"}), 404
    return send_file(path, as_attachment=True, download_name=name)


@app.get("/api/download")
def api_download():
    """02_output 配下のフォルダをZIPにして返す。"""
    folder = safe_name(request.args.get("folder", ""))
    full = os.path.join(DIR_OUTPUT, folder)
    if not folder or not os.path.isdir(full):
        return jsonify({"error": "フォルダが見つかりません"}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(full):
            for fn in files:
                if fn.startswith("~$"):
                    continue
                fp = os.path.join(root, fn)
                arc = os.path.join(folder, os.path.relpath(fp, full))
                zf.write(fp, arc)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="%s.zip" % folder,
                     mimetype="application/zip")


# ---------------------------------------------------------------------------
# WEB転記（e-再生医療フォームへ自動入力）
#   既存の 99_data/src/web_fill.py（対話型）を一切変更せず、サブプロセスとして
#   起動し、stdin にEnter/qを送り、stdout をSSEでライブ表示する薄いラッパー。
#   ・ブラウザは各自のPC上に開く（127.0.0.1 のダッシュボードから起動）
#   ・送信はしない（web_fill.py の仕様どおり、入力＝下書きまで）
# ---------------------------------------------------------------------------
WEB_LOCK = threading.Lock()
WEB = {"proc": None, "mode": None}       # 同時に1セッションのみ


def _web_mapping():
    try:
        with io.open(WEB_MAPPING, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _playwright_ready():
    """web_fill.py と同じ Python で playwright パッケージが入っているか。"""
    try:
        import importlib.util
        return importlib.util.find_spec("playwright") is not None
    except Exception:
        return False


def _web_output_folder():
    """web_mapping の output_folder_contains に一致する最新の 02_output フォルダ名。"""
    m = _web_mapping()
    want = m.get("output_folder_contains", "")
    best = None
    try:
        for name in os.listdir(DIR_OUTPUT):
            full = os.path.join(DIR_OUTPUT, name)
            if os.path.isdir(full) and (not want or want in name):
                ts = os.path.getmtime(full)
                if best is None or ts > best[0]:
                    best = (ts, name)
    except OSError:
        pass
    return best[1] if best else ""


@app.get("/api/web/status")
def api_web_status():
    proc = WEB["proc"]
    running = proc is not None and proc.poll() is None
    m = _web_mapping()
    return jsonify({
        "playwright": _playwright_ready(),
        "url": m.get("url", ""),
        "output_folder": _web_output_folder(),
        "running": running,
        "mode": WEB["mode"] if running else "",
    })


def _sse(event, data):
    return "event: %s\ndata: %s\n\n" % (event, json.dumps(data, ensure_ascii=False))


def _child_env():
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    env["PYTHONUNBUFFERED"] = "1"
    return env


@app.get("/api/web/start")
def api_web_start():
    """web_fill.py を起動し、ブラウザを開いて自動入力の対話を開始。進捗をSSEで流す。
       mode=fill: 通常の自動入力 / mode=dump: フォーム項目の抽出（メンテ用）。"""
    mode = request.args.get("mode", "fill")
    args = [sys.executable, WEB_FILL] + (["--dump"] if mode == "dump" else [])

    @stream_with_context
    def generate():
        err = None
        proc = None
        WEB_LOCK.acquire()
        try:
            if WEB["proc"] is not None and WEB["proc"].poll() is None:
                err = "すでにWEB転記セッションが実行中です。先に「終了」してください。"
            elif not os.path.exists(WEB_FILL):
                err = "web_fill.py が見つかりません。"
            else:
                try:
                    proc = subprocess.Popen(
                        args, cwd=BASE, env=_child_env(),
                        stdin=subprocess.PIPE,
                        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        bufsize=1, universal_newlines=True,
                        encoding="utf-8", errors="replace",
                    )
                    WEB["proc"] = proc
                    WEB["mode"] = mode
                except Exception as e:
                    err = "起動に失敗: %r" % e
        finally:
            WEB_LOCK.release()

        if err:
            yield _sse("error", {"message": err})
            return

        yield _sse("ready", {"mode": mode})
        try:
            for line in iter(proc.stdout.readline, ""):
                line = line.rstrip("\n")
                if line != "":
                    yield _sse("log", {"line": line})
        finally:
            try:
                proc.stdout.close()
            except Exception:
                pass
            code = proc.wait()
            with WEB_LOCK:
                if WEB["proc"] is proc:
                    WEB["proc"] = None
                    WEB["mode"] = None
            yield _sse("done", {"code": code})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.post("/api/web/send")
def api_web_send():
    """実行中セッションの stdin に指示を送る。cmd=fill→Enter（このページを入力/次へ） / cmd=quit→q（終了）。"""
    cmd = "fill"
    if request.is_json:
        cmd = (request.get_json(silent=True) or {}).get("cmd", "fill")
    else:
        cmd = request.form.get("cmd", "fill")

    proc = WEB["proc"]
    if proc is None or proc.poll() is not None:
        return jsonify({"error": "WEB転記セッションが実行されていません。"}), 400
    try:
        proc.stdin.write("q\n" if cmd == "quit" else "\n")
        proc.stdin.flush()
    except Exception as e:
        return jsonify({"error": "送信に失敗: %r" % e}), 500
    return jsonify({"ok": True})


@app.post("/api/web/stop")
def api_web_stop():
    """セッションを強制停止（ブラウザも閉じる）。まず q を送って猶予を与え、無理なら terminate。"""
    proc = WEB["proc"]
    if proc is None or proc.poll() is not None:
        return jsonify({"ok": True})
    try:
        proc.stdin.write("q\n")
        proc.stdin.flush()
    except Exception:
        pass
    try:
        proc.wait(timeout=3)
    except Exception:
        try:
            proc.terminate()
        except Exception:
            pass
    return jsonify({"ok": True})


@app.get("/api/web/setup")
def api_web_setup():
    """初回準備：playwright（pip）と Chromium ブラウザを導入。進捗をSSEで流す。"""
    @stream_with_context
    def generate():
        steps = [
            ("Playwright を導入しています…", [sys.executable, "-m", "pip", "install", "playwright"]),
            ("Chromium ブラウザを導入しています（数分かかることがあります）…",
             [sys.executable, "-m", "playwright", "install", "chromium"]),
        ]
        for title, cmd in steps:
            yield _sse("log", {"line": "▶ " + title})
            try:
                proc = subprocess.Popen(
                    cmd, cwd=BASE, env=_child_env(),
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    bufsize=1, universal_newlines=True,
                    encoding="utf-8", errors="replace",
                )
            except Exception as e:
                yield _sse("error", {"message": "導入コマンドの起動に失敗: %r" % e})
                return
            for line in iter(proc.stdout.readline, ""):
                line = line.rstrip("\n")
                if line != "":
                    yield _sse("log", {"line": line})
            proc.stdout.close()
            if proc.wait() != 0:
                yield _sse("error", {"message": "導入に失敗しました。ネットワーク環境をご確認ください。"})
                return
        yield _sse("done", {"playwright": _playwright_ready()})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.get("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html; charset=utf-8")


# INDEX_HTML は index_html.py から読み込む（見通しのため分離）
from index_html import INDEX_HTML  # noqa: E402


if __name__ == "__main__":
    import webbrowser
    import threading

    port = int(os.environ.get("PRP_DASH_PORT", "8765"))
    url = "http://127.0.0.1:%d" % port
    print("=" * 60)
    print(" PRP 自動転記ツール ダッシュボード")
    print(" ブラウザで開く: %s" % url)
    print(" 終了するには、この画面で Ctrl + C")
    print("=" * 60)

    # 起動直後にブラウザを開く（reloader無効時のみ）
    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)

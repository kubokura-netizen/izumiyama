# -*- coding: utf-8 -*-
"""
領収書パイプライン共通ロジック（読取・確定の両方で使う）。
 ・設定(settings.json)の読み込み … インストール時に IN/OUT を相手環境へ
 ・ファイル名生成  YYYYMMDD_相手先_品目_金額円.pdf
 ・年月フォルダ/月シート判定
 ・重複名の回避（_2, _3 …）
※ 送信・課金なし。ローカルのファイル操作のみ。
"""
import os
import re
import json
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SETTINGS_PATH = os.path.join(ROOT, "settings.json")

DEFAULTS = {
    # 読取対象（相手の環境に合わせて設定）。相対なら ROOT 基準。
    "input_dir": "01_input",
    # 年月フォルダとマスターExcelを置く親。空なら 02_output を使う。
    "output_root": "",
    # 蓄積先の経費入力表.xlsx。既定は 03_経費入力表蓄積\\経費入力表.xlsx。
    # ここへ月シート「経費(N)」に蓄積していく（読取の下書き取込.xlsxは確認用に別途残る）。
    "master_xlsx": "03_経費入力表蓄積/経費入力表.xlsx",
    # 蓄積先が未作成なら、この元フォーマットからコピーして作る（保存用＋経費(1〜12)）。
    "template_xlsx": "97_テンプレート/経費入力表.xlsx",
    # 月シート名の書式（{m}=月番号）。例: 経費(8)
    "month_sheet_fmt": "経費({m})",
    # 原本の扱い: copy=コピーして年月へ / move=移動
    "file_action": "copy",
}


def _abs(root, p):
    if not p:
        return ""
    return p if os.path.isabs(p) else os.path.join(root, p)


def load_settings():
    """settings.json を読み、絶対パスに整えて返す。無ければ既定。"""
    cfg = dict(DEFAULTS)
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            cfg.update({k: v for k, v in json.load(f).items() if v is not None})
    except Exception:
        pass
    out = os.path.join(ROOT, "02_output")
    cfg["input_dir"] = _abs(ROOT, cfg.get("input_dir") or "01_input")
    cfg["output_root"] = _abs(ROOT, cfg.get("output_root")) or out
    cfg["master_xlsx"] = _abs(ROOT, cfg.get("master_xlsx")) \
        or os.path.join(ROOT, "03_経費入力表蓄積", "経費入力表.xlsx")
    cfg["template_xlsx"] = _abs(ROOT, cfg.get("template_xlsx")) \
        or os.path.join(ROOT, "97_テンプレート", "経費入力表.xlsx")
    return cfg


def save_settings(cfg):
    keep = {k: cfg.get(k, DEFAULTS[k]) for k in DEFAULTS}
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(keep, f, ensure_ascii=False, indent=2)


# ---- 日付 --------------------------------------------------------------
def parse_date(s):
    """'2026-05-23' 等 → datetime。取れなければ None。"""
    if not s:
        return None
    if isinstance(s, datetime.datetime):
        return s
    m = (re.search(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", str(s))
         or re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", str(s)))
    if not m:
        return None
    try:
        return datetime.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def ym_folder(dt):
    """datetime → '2026年5月'（年月フォルダ名）。"""
    return "%d年%d月" % (dt.year, dt.month)


def month_sheet_name(dt, fmt):
    return fmt.replace("{m}", str(dt.month))


# ---- ファイル名 --------------------------------------------------------
_BAD = re.compile(r'[\\/:*?"<>|\r\n\t]')


def sanitize(s, limit=40):
    s = _BAD.sub("", str(s or "")).strip()
    s = re.sub(r"\s+", " ", s)
    return s[:limit].strip()


def build_filename(fields, suffix=""):
    """YYYYMMDD_相手先_品目_金額円.pdf を作る。日付か金額が無ければ None。
       suffix は '請求書'/'領収証' 等の枝（任意）。"""
    dt = parse_date(fields.get("日付"))
    amt = re.sub(r"[^0-9]", "", str(fields.get("金額") or ""))
    if not dt or not amt:
        return None
    aite = sanitize(fields.get("相手先"), 24) or "相手先不明"
    item = sanitize(fields.get("内容"), 24) or "品目不明"
    base = "%04d%02d%02d_%s_%s_%s円" % (dt.year, dt.month, dt.day, aite, item, amt)
    if suffix:
        base += "_" + sanitize(suffix, 8)
    return base + ".pdf"


def unique_path(folder, filename):
    """folder/filename が既存なら _2, _3 … を付けて衝突回避したフルパスを返す。"""
    stem, ext = os.path.splitext(filename)
    cand = os.path.join(folder, filename)
    i = 2
    while os.path.exists(cand):
        cand = os.path.join(folder, "%s_%d%s" % (stem, i, ext))
        i += 1
    return cand

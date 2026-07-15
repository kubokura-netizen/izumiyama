# -*- coding: utf-8 -*-
"""
マスターテンプレの「緑（転記対象）」を {{トークン}} へ置換する（トークン方式への移行）。

方針:
  ・緑ラン連続スパン単位で、結合テキストを docs_config / sop_replacements の
    find（サンプル値）と照合し、一致部分を {{トークン}} へ置換する。
  ・トークン名は token_for(source)（ヒアリング項目名ベース）。
  ・一致しない緑（要判断・未マッピング）はそのまま残し、レポートに出す。
  ・すでに {{...}} 化済みの緑は find に一致しないので自動的にスキップ（冪等）。
  ・docx は python-docx で往復（全パーツ保持）。xlsx 様式は別処理（XML手術）。

使い方:
  python tokenize_master.py            # 全docxマスターをトークン化
  出力レポート: 03_logs/_tokenize_report.txt
"""
import os, io, json, glob, sys

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.dirname(SRC_DIR)
BASE = os.path.dirname(DATA)
TPL = os.path.join(DATA, "テンプレート")
DOCS_CONFIG = os.path.join(DATA, "マッピング", "docs_config.json")
SOP_CONFIG = os.path.join(DATA, "マッピング", "sop_replacements.json")

sys.path.insert(0, SRC_DIR)
from transcribe_sop import _iter_all_paragraphs, _is_green_rgb
from transcribe import token_for   # トークン名の唯一の定義（テンプレ表記と穴埋めキーを一致させる）


def _finds_for_docx(doc_cfg):
    """[(find_str, token, file_contains)] を長い順で返す。"""
    pairs = []
    for rep in doc_cfg.get("docx", []):
        tok = token_for(rep.get("source", {}))
        if not tok:
            continue
        finds = rep.get("find", [])
        if isinstance(finds, str):
            finds = [finds]
        fc = rep.get("file_contains", "")
        for f in finds:
            if f:
                pairs.append((f, tok, fc))
    pairs.sort(key=lambda x: len(x[0]), reverse=True)
    return pairs


def _finds_for_sop(sop_cfg):
    pairs = []
    for rep in sop_cfg.get("replacements", []):
        src = rep.get("source", {})
        # SOPの日付はfindで年月日/月日を切替 → tokenも分ける
        finds = rep.get("find", [])
        if isinstance(finds, str):
            finds = [finds]
        for f in finds:
            if not f:
                continue
            if src.get("t") == "today":
                tok = "作業日" if ("年" in f) else "作業日月日"
                if "/" in f:
                    tok = "作業日スラッシュ"
            else:
                tok = token_for(src)
            if tok:
                pairs.append((f, tok, ""))
    pairs.sort(key=lambda x: len(x[0]), reverse=True)
    return pairs


def _green_spans(p):
    """段落 p の連続緑ランのスパンを [(start_idx, end_idx_exclusive)] で返す。"""
    runs = p.runs
    spans = []
    i = 0
    n = len(runs)
    while i < n:
        if _is_green_rgb(runs[i].font.color):
            j = i
            while j < n and _is_green_rgb(runs[j].font.color):
                j += 1
            spans.append((i, j))
            i = j
        else:
            i += 1
    return spans


def tokenize_docx_file(path, pairs, report, fbase):
    from docx import Document
    d = Document(path)
    applied = 0
    for p in _iter_all_paragraphs(d):
        for (s, e) in _green_spans(p):
            runs = p.runs[s:e]
            T = "".join(r.text for r in runs)
            if not T.strip():
                continue
            # 既存トークン {{...}} は保護し、その外側のサンプル値だけを単一パスで置換。
            # 単一パス（re.sub）なので、置換後に生成した {{…}} が別findで再マッチしない＝冪等。
            import re as _re2
            applicable = [(f, tok) for (f, tok, fc) in pairs if (not fc or fc in fbase) and f]
            applicable.sort(key=lambda x: len(x[0]), reverse=True)   # 長い順優先
            fmap = {}
            for f, tok in applicable:
                fmap.setdefault(f, tok)
            if fmap:
                alt = _re2.compile("|".join(_re2.escape(f) for f in fmap))
                segs = _re2.split(r"(\{\{[^{}]*\}\})", T)
                for k in range(0, len(segs), 2):     # 偶数index＝トークン外
                    segs[k] = alt.sub(lambda m: "{{%s}}" % fmap[m.group(0)], segs[k])
                newT = "".join(segs)
            else:
                newT = T
            if newT != T:
                runs[0].text = newT
                for r in runs[1:]:
                    r.text = ""
                applied += 1
            # 置換後もサンプル値らしき緑が残るか（トークンでない生テキスト）→ レポート
            residual = newT
            # トークン部を除去して残りを見る
            import re as _re
            residual_wo = _re.sub(r"\{\{[^{}]*\}\}", "", residual).strip()
            if residual_wo:
                report.setdefault(fbase, []).append(residual_wo)
    if applied:
        d.save(path)
    return applied


def main():
    docs = json.load(io.open(DOCS_CONFIG, encoding="utf-8"))
    sop = json.load(io.open(SOP_CONFIG, encoding="utf-8"))
    report = {}
    total = 0
    log = []

    for dk, doc in docs.get("documents", {}).items():
        folder = os.path.join(TPL, doc["folder"])
        if not os.path.isdir(folder):
            continue
        pairs = _finds_for_docx(doc)
        for f in sorted(glob.glob(os.path.join(folder, "*.docx"))):
            fb = os.path.basename(f)
            if fb.startswith("~$"):
                continue
            a = tokenize_docx_file(f, pairs, report, fb)
            total += a
            if a:
                log.append("[%s] %s: %d spans tokenized" % (dk, fb, a))

    # SOP
    sop_folder = os.path.join(TPL, sop.get("template_folder", "SOP"))
    if os.path.isdir(sop_folder):
        pairs = _finds_for_sop(sop)
        for f in sorted(glob.glob(os.path.join(sop_folder, "*.docx"))):
            fb = os.path.basename(f)
            if fb.startswith("~$"):
                continue
            a = tokenize_docx_file(f, pairs, report, fb)
            total += a
            if a:
                log.append("[SOP] %s: %d spans tokenized" % (fb, a))

    outp = os.path.join(BASE, "03_logs", "_tokenize_report.txt")
    with io.open(outp, "w", encoding="utf-8") as o:
        o.write("=== トークン化サマリ: 合計 %d spans ===\n" % total)
        for line in log:
            o.write(line + "\n")
        o.write("\n=== マッピングされなかった残存緑（要判断）===\n")
        for fb, items in report.items():
            uniq = sorted(set(items))
            o.write("\n## %s (%d種)\n" % (fb, len(uniq)))
            for it in uniq:
                s = it.replace("\n", "\\n")
                o.write("   %r\n" % (s[:120]))
    print("done total=%d report=%s" % (total, outp))


if __name__ == "__main__":
    main()

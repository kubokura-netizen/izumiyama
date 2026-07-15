# -*- coding: utf-8 -*-
"""
フォルダ型テンプレ（2種関節系PRP / 3種筋腱靭帯系PRP / SOP）への転記エンジン。

各テンプレは「フォルダ」で、中に
  ・様式第一の二（.xlsx）… セル転記
  ・その他の書類（.docx）… サンプル値の一括置換（find→ヒアリング値）
  ・PDF・論文・追加xlsx 等 … 無変更で同梱（参照用）
を含む。

処理:
  1) 99_data/マッピング/docs_config.json を読み込む。
  2) テンプレフォルダを 02_output/(フォルダ名)_(ヒアリング名)/ へフォルダごとコピー。
  3) 様式xlsx はセル転記（resolve→安全書込）。
  4) 直下の各 .docx はサンプル値を一括置換（本文・表・ヘッダ/フッタ、図・書式は保持）。
  5) サブフォルダ内（論文PDF等）は対象外＝無変更で同梱。

※ Excel処理は transcribe.py（openpyxl）、Word処理は transcribe_sop.py の置換関数を再利用。
"""
import os, io, json, glob, shutil

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.dirname(SRC_DIR)
DOCS_CONFIG = os.path.join(DATA, "マッピング", "docs_config.json")


def _clean(s):
    return "" if s is None else str(s).strip()


def _long(path):
    """Windowsの260文字パス上限を回避する拡張パス（\\\\?\\）へ変換。
       深いサブフォルダ内の長いPDF名で copytree が WinError 3 になるのを防ぐ。
       他OSではそのまま返す。"""
    if os.name != "nt":
        return path
    ap = os.path.abspath(path)
    if ap.startswith("\\\\?\\"):
        return ap
    if ap.startswith("\\\\"):                     # UNCパス
        return "\\\\?\\UNC\\" + ap[2:]
    return "\\\\?\\" + ap


# ---- Excelセルの色マーク（緑=転記済み） ----
def _is_green_cell(cell):
    try:
        col = cell.font.color
        if col is not None and getattr(col, "rgb", None):
            s = str(col.rgb).upper()
            return s.endswith("00B050") or s.endswith("008000")
    except Exception:
        pass
    return False


def _set_cell_color(cell, hexcolor):
    from openpyxl.styles import Font
    f = cell.font
    cell.font = Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic,
                     underline=f.underline, strike=f.strike, vertAlign=f.vertAlign,
                     color=hexcolor)


GREEN = "FF00B050"
BLACK = "FF000000"


def _build_docx_pairs(entries, hearing, kits, TX):
    """docx置換ペアを生成（空値は除外、長い順）。ハイブリッド方式：
       ・主：テンプレの {{ヒアリング項目名}} を実値へ穴埋め（トークン方式）。
       ・副：旧サンプル値(find)→実値の置換も併載（トークン化漏れの緑を取りこぼさない）。
       トークン化済みの緑は {{...}} なのでサンプルfindは発火せず二重置換にならない。"""
    pairs = []
    for rep in entries:
        src = rep.get("source", {})
        val, _ = TX.resolve(src, hearing, kits)
        val = _clean(val)
        if not val:
            continue
        fc = rep.get("file_contains", "")   # 空=全docx対象／指定=そのファイル名を含むもののみ
        desc = rep.get("desc", "")
        tok = TX.token_for(src)
        if tok:
            pairs.append(("{{%s}}" % tok, val, desc, fc))   # 主：トークン穴埋め
        finds = rep.get("find", [])
        if isinstance(finds, str):
            finds = [finds]
        for find in finds:                                  # 副：サンプル値フォールバック
            if find and val != find:
                pairs.append((find, val, desc, fc))
    pairs.sort(key=lambda t: len(t[0]), reverse=True)
    return pairs


def _xml_escape(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _col_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _write_xlsx_cells_xmlsurgery(xlsx_path, sheet_title, coord_values, run_log):
    """openpyxl保存を使わず、セル値だけをシートXMLへ直接書き込む（図形・customXml等を保持）。
       既存 <c> は inlineStr で更新、無ければ列順に挿入。他パーツはバイトコピー。"""
    import re as _re
    import zipfile as _zip
    if not coord_values:
        return
    try:
        z = _zip.ZipFile(xlsx_path, "r")
        names = z.namelist()
        data = {n: z.read(n) for n in names}
        z.close()
    except Exception as e:
        run_log.append("[xmlsurgery] 読込失敗、openpyxl保存にフォールバック: %r" % e)
        raise

    wbxml = data.get("xl/workbook.xml", b"").decode("utf-8", "ignore")
    relsxml = data.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8", "ignore")
    # シート名 → r:id → ターゲットXML
    sheet_path = None
    rid = None
    for tag in _re.findall(r"<sheet\b[^>]*/>", wbxml):
        nm = _re.search(r'name="([^"]*)"', tag)
        ri = _re.search(r'r:id="([^"]*)"', tag)
        if nm and ri and nm.group(1) == sheet_title:
            rid = ri.group(1); break
    if rid is None:
        m = _re.search(r'<sheet\b[^>]*r:id="([^"]*)"', wbxml)
        rid = m.group(1) if m else None
    if rid:
        rm = _re.search(r'<Relationship\b[^>]*Id="%s"[^>]*Target="([^"]*)"' % _re.escape(rid), relsxml)
        if rm:
            tgt = rm.group(1).lstrip("/")
            sheet_path = tgt if tgt.startswith("xl/") else "xl/" + tgt
    if not sheet_path or sheet_path not in data:
        cand = [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        if not cand:
            raise RuntimeError("シートXMLが見つからない")
        sheet_path = sorted(cand)[0]

    xml = data[sheet_path].decode("utf-8")

    def _cell_xml(coord, s_attr, val):
        if val is None or val == "":
            return '<c r="%s"%s/>' % (coord, s_attr)
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
            coord, s_attr, _xml_escape(val))

    missing = {}
    for coord, val in coord_values.items():
        pat = _re.compile(r'<c r="%s"([^>/]*)(/>|>.*?</c>)' % _re.escape(coord), _re.S)
        m = pat.search(xml)
        if m:
            sm = _re.search(r'\ss="(\d+)"', m.group(1))
            s_attr = ' s="%s"' % sm.group(1) if sm else ""
            xml = xml[:m.start()] + _cell_xml(coord, s_attr, val) + xml[m.end():]
        else:
            missing[coord] = val

    # 既存行に無いセルは列順で挿入（無ければ行ごと挿入）
    for coord, val in missing.items():
        mm = _re.match(r"([A-Z]+)(\d+)", coord)
        if not mm:
            continue
        col, rownum = mm.group(1), int(mm.group(2))
        cellxml = _cell_xml(coord, "", val)
        rowpat = _re.compile(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % rownum, _re.S)
        rm2 = rowpat.search(xml)
        if rm2:
            inner = rm2.group(2)
            # 列順の挿入位置を探す
            pos = None
            for cm in _re.finditer(r'<c r="([A-Z]+)\d+"', inner):
                if _col_num(cm.group(1)) > _col_num(col):
                    pos = cm.start(); break
            inner2 = (inner[:pos] + cellxml + inner[pos:]) if pos is not None else (inner + cellxml)
            xml = xml[:rm2.start()] + rm2.group(1) + inner2 + rm2.group(3) + xml[rm2.end():]
        else:
            # 行が無い → sheetData末尾へ（Excelは順不同を許容）
            newrow = '<row r="%d">%s</row>' % (rownum, cellxml)
            xml = xml.replace("</sheetData>", newrow + "</sheetData>", 1)

    data[sheet_path] = xml.encode("utf-8")

    # inlineStr化で共有文字列への参照数が減るため、sharedStrings の count を実数に合わせる
    # （count不整合＝Excelの「修復されたレコード」警告の原因）。sharedStrings本体は改変しない
    # （未参照のトークン文字列が残るが表示されない。空<t>化＝別の破損原因になるため触らない）。
    ss_key = "xl/sharedStrings.xml"
    if ss_key in data:
        total_refs = 0
        for n in names:
            if _re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                content = xml if n == sheet_path else data[n].decode("utf-8", "ignore")
                total_refs += len(_re.findall(r'<c\b[^>]*\bt="s"', content))
        ss = data[ss_key].decode("utf-8")
        ss = _re.sub(r'(<sst\b[^>]*\bcount=")\d+(")', r"\g<1>%d\g<2>" % total_refs, ss, count=1)
        data[ss_key] = ss.encode("utf-8")

    tmp = xlsx_path + ".tmp"
    zo = _zip.ZipFile(tmp, "w", _zip.ZIP_DEFLATED)
    try:
        for n in names:
            zo.writestr(n, data[n])
    finally:
        zo.close()
    os.replace(tmp, xlsx_path)


def _fill_xlsx_tokens(ws, hearing, kits, TX, saisei):
    """様式xlsxセル内に残る {{項目名}} をヒアリング値で穴埋め（部分セル・旧トークン対応）。
       特殊トークン（採血量/治療価格/キット製造方法/委員会/場所）を名前で解決し、
       それ以外は「より→から」等を吸収してヒアリングlabel直接参照でフォールバック。"""
    import re as _re
    TOKENRE = _re.compile(r"\{\{([^{}]*)\}\}")
    MARKS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"

    def _kit_index(name):
        m = _re.search(r"メーカー([①-⑮])", name)
        if m:
            return MARKS.find(m.group(1))
        m = _re.search(r"メーカー([A-O])", name)
        if m:
            return ord(m.group(1)) - ord("A")
        return 0

    def _resolve_named(name):
        name = name.strip()
        n2 = name.replace("より", "から")           # 旧表記ゆれ吸収
        if name == "必要な採血量":
            return hearing.blood_volume()
        if name == "治療価格":
            return hearing.kit_price_block(saisei)
        if "治療費の設定" in name:
            # 様式は「①キット名を用いた治療：{価格}」形式のため、ブロックではなく単一価格を返す
            idx = 2 if "②" in name else (3 if "③" in name else 1)
            plabel = "再生医療%s（右記タブから選択）の治療費の設定（税別）" % MARKS[idx - 1]
            return TX.clean(hearing.lookup(plabel, "", 1)) or ""
        # キット製造方法：スロット番号（①/A等）付きは「その1キット単体」、無印は全キットのブロック
        def _kit_method(part):
            ks = hearing.prp_kits()
            i = _kit_index(name)
            has_mark = bool(_re.search(r"メーカー([①-⑮A-O])", name))
            if has_mark:
                return hearing.kit_method(ks[i], part) if i < len(ks) else ""
            return hearing.kit_block(part, numbered=True)
        if name == "採取方法" or ("採取" in name and "キット" in name):
            return _kit_method("採取")
        if name == "加工方法" or ("加工の方法" in name):
            return _kit_method("加工")
        if name == "投与方法" or ("投与の方法" in name):
            return _kit_method("投与")
        if "PRPキットメーカー" in name and "右記タブ" in name:
            ks = hearing.prp_kits(); i = _kit_index(name)
            return _re.sub(r"\s*（[^（）]*）\s*$", "", ks[i]).strip() if i < len(ks) else ""
        if "認定再生医療等委員会" in name and "認定番号" in name:
            cname = hearing.lookup("認定再生医療等委員会の名称", "審査委員会", 1)
            return hearing.committee_field(cname, 2)
        # 汎用：ヒアリングlabel直接参照（採血場所/投与場所/保険名称/問い合わせ先/委員会名称/電話/メール等）
        v = hearing.lookup(n2, "", 1)
        return TX.clean(v) if v not in (None, "") else ""

    n = 0
    for row in ws.iter_rows():
        for c in row:
            if not isinstance(c.value, str) or "{{" not in c.value:
                continue
            def _rep(m):
                val = _resolve_named(m.group(1))
                return val if val else m.group(0)   # 未解決は残す（後段のclear_tokensで空欄化）
            newv = TOKENRE.sub(_rep, c.value)
            # 採血量の注記がテンプレ固定文と値で重複した場合は1つに畳む
            newv = _re.sub(r"(（使用キットにより異なる）)[（(]使用キットにより異なる[）)]", r"\1", newv)
            if newv != c.value:
                TX.safe_set(ws, c.coordinate, newv)
                n += 1
    return n


def _fill_all_docx_tokens(d, hearing, kits, saisei, TX, iter_paras):
    """docx内に残る任意の {{項目名}} を汎用リゾルバで穴埋め（前世代トークンの取りこぼし防止）。
       段落全体が単一トークン→複数行(改行br)で差込／文中インライン→ラン置換（1行化）。"""
    import re as _re
    from transcribe_sop import _run_green, _replace_in_paragraph
    TOKEN = _re.compile(r"\{\{([^{}]*)\}\}")
    STANDALONE = _re.compile(r"^\s*\{\{([^{}]*)\}\}\s*$")
    cache = {}

    def val_for(nm):
        if nm not in cache:
            cache[nm] = TX.resolve_token_by_name(nm, hearing, kits, saisei) or ""
        return cache[nm]

    n = 0
    for p in iter_paras(d):
        if "{{" not in p.text:
            continue
        m = STANDALONE.match(p.text)
        if m:
            v = val_for(m.group(1).strip())
            if v == "":
                continue                       # 空は _clear_doc_tokens に任せる
            lines = v.split("\n")
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            run = p.add_run(lines[0]); _run_green(run)
            for ln in lines[1:]:
                run.add_break(); run = p.add_run(ln); _run_green(run)
            n += 1
        else:
            pairs = []
            for nm in set(x.strip() for x in TOKEN.findall(p.text)):
                v = val_for(nm)
                if v != "":
                    pairs.append(("{{%s}}" % nm, v.replace("\n", " ")))
            if pairs:
                n += _replace_in_paragraph(p, pairs)
    return n


def _clear_doc_tokens(d, iter_paras):
    """穴埋めされずに残った {{...}} を空欄化（未充足トークンをクライアントに見せない）。"""
    import re as _re
    TOKEN_RE = _re.compile(r"\{\{[^{}]*\}\}")
    n = 0
    for p in iter_paras(d):
        for r in p.runs:
            if r.text and "{{" in r.text and "}}" in r.text:
                nt = TOKEN_RE.sub("", r.text)
                if nt != r.text:
                    r.text = nt
                    n += 1
        # ラン跨ぎの残存トークンも段落テキストで検出して除去
        full = "".join(r.text for r in p.runs)
        if "{{" in full and "}}" in full and p.runs:
            new = TOKEN_RE.sub("", full)
            if new != full:
                p.runs[0].text = new
                for r in p.runs[1:]:
                    r.text = ""
                n += 1
    return n


def run_docs(hearing, hearing_path, dir_tpl, dir_output, run_log, run_dt):
    """フォルダ型テンプレ一式を転記出力。戻り値: (出力フォルダ一覧, log_rows)。"""
    try:
        from docx import Document
    except ImportError:
        run_log.append("[docs] python-docx 未導入のためWord書類はスキップ")
        Document = None
    import openpyxl
    import transcribe as TX
    try:
        from transcribe_sop import (_replace_in_paragraph, _iter_all_paragraphs,
                                     _replace_lowlevel, _fill_token_multiline, _set_cell_multiline,
                                     reset_doc_green_to_black)
    except Exception as e:
        run_log.append("[docs] Word置換関数の読込失敗: %r" % e)
        _replace_in_paragraph = _iter_all_paragraphs = _replace_lowlevel = None
        _fill_token_multiline = _set_cell_multiline = None

    if not os.path.exists(DOCS_CONFIG):
        run_log.append("[docs] docs_config.json が無いためスキップ")
        return [], []
    cfg = json.load(io.open(DOCS_CONFIG, encoding="utf-8"))

    kits = hearing.prp_kits()
    base = os.path.splitext(os.path.basename(hearing_path))[0]
    out_folders = []
    log_rows = []
    green_report = []   # 元が緑ターゲットだが未転記だったセル（要確認候補）

    for doc_key, doc in cfg.get("documents", {}).items():
        tpl_folder = os.path.join(dir_tpl, doc["folder"])
        if not os.path.isdir(tpl_folder):
            run_log.append("[%s] テンプレフォルダ無し: %s" % (doc_key, tpl_folder))
            continue
        out_folder = os.path.join(dir_output, "%s_%s" % (doc["folder"], base))
        try:
            if os.path.exists(out_folder):
                shutil.rmtree(_long(out_folder))
            # 拡張パス（\\?\）で複製し、深い階層の長いファイル名でも 260 文字上限に阻まれないようにする
            shutil.copytree(_long(tpl_folder), _long(out_folder))   # PDF・論文・追加xlsx等も無変更で同梱
        except PermissionError as e:
            run_log.append("[%s] ★スキップ：出力先のファイルが開かれています。"
                           "02_output内のWord/Excel（特に前回の出力）を閉じてから再実行してください。（%s）"
                           % (doc_key, e))
            continue
        except Exception as e:
            run_log.append("[%s] コピー失敗でスキップ: %r" % (doc_key, e))
            continue

        # --- 様式Excel（セル転記） ---
        n_x_done = n_x_check = 0
        exc = doc.get("excel")
        if exc:
            xpath = os.path.join(out_folder, exc["file"])
            if os.path.exists(xpath):
                wb = openpyxl.load_workbook(xpath)
                sheet = exc.get("sheet")
                ws = wb[sheet] if sheet in wb.sheetnames else wb[wb.sheetnames[0]]
                # 図形保持のため、書込みは最後にXML手術で行う。転記前の全セル値を控える。
                orig_vals = {c.coordinate: c.value for r in ws.iter_rows() for c in r if c.value is not None}
                # 転記前：テンプレの緑セルを記録し、一旦すべて黒に戻す
                orig_green = {}
                for row in ws.iter_rows():
                    for c in row:
                        if _is_green_cell(c):
                            orig_green[c.coordinate] = _clean(c.value)
                            _set_cell_color(c, BLACK)
                written = set()   # 転記したセル（後で緑にする）
                for e in exc.get("entries", []):
                    if e.get("_disabled"):
                        continue
                    val, st = TX.resolve(e["source"], hearing, kits)
                    if st == TX.ST_UNMAP or val in (None, ""):
                        continue
                    val = TX.apply_tf(val, e.get("tf", "text"))
                    if val in (None, ""):
                        continue
                    if TX.safe_set(ws, e["cell"], val):
                        written.add(e["cell"])
                        if st == TX.ST_CHECK or "確認対象" in _clean(e.get("note")):
                            n_x_check += 1
                        else:
                            n_x_done += 1
                        log_rows.append(TX.make_row(run_dt, doc_key, e, val, st, e.get("note")))
                # 医師の可変転記（最大N名。B列ラベルで医師ブロックを動的検出し、氏名/所属を順に）
                df = exc.get("doctor_fill")
                if df:
                    from openpyxl.utils import get_column_letter as _gl
                    blabel = df.get("block_label", "")
                    col = int(df.get("col", 12))
                    noff = int(df.get("name_offset", 1))
                    aoff = df.get("affil_offset")
                    maxn = int(df.get("max", 15))
                    blocks = [r for r in range(1, ws.max_row + 1)
                              if isinstance(ws.cell(r, 2).value, str) and blabel and blabel in ws.cell(r, 2).value]
                    docs = hearing.doctors()
                    med, _mst = TX.resolve({"t": "hearing",
                                            "label": "医療機関/名称（診療所開設届上）",
                                            "section": "法人/医療機関"}, hearing, kits)
                    for i, dname in enumerate(docs[:maxn]):
                        if i >= len(blocks):
                            run_log.append("[%s] 医師%d名目以降はテンプレの医師ブロック不足で未転記（ブロック追加が必要）"
                                           % (doc_key, i + 1))
                            break
                        br = blocks[i]
                        nc = "%s%d" % (_gl(col), br + noff)
                        TX.safe_set(ws, nc, dname); written.add(nc)
                        if aoff is not None and med:
                            ac = "%s%d" % (_gl(col), br + int(aoff))
                            TX.safe_set(ws, ac, med); written.add(ac)
                        n_x_done += 1
                    if docs:
                        run_log.append("[%s] 医師転記=%d名（テンプレ医師ブロック=%d）"
                                       % (doc_key, min(len(docs), len(blocks)), len(blocks)))
                # セル内の一部だけ置換（保険名称/問い合わせ先/採取方法など。定型文は保持）
                for ce in exc.get("cell_edits", []):
                    cur = ws[ce["cell"]].value
                    cur = "" if cur is None else str(cur)
                    val, _st = TX.resolve(ce.get("source", {}), hearing, kits)
                    val = _clean(val)
                    if not val:
                        continue
                    op = ce.get("op")
                    new = cur
                    # トークン方式：セル内の {{項目名}} を先に穴埋め（部分セルのトークン化対応）
                    _tok = TX.token_for(ce.get("source", {}))
                    if _tok and ("{{%s}}" % _tok) in new:
                        new = new.replace("{{%s}}" % _tok, val)
                    if op == "replace_find":            # 指定文字列を丸ごと置換
                        for fnd in ce.get("find", []):
                            if fnd and fnd in new:
                                new = new.replace(fnd, val)
                    elif op == "replace_from":          # マーカー以降を置換（前は保持）
                        marker = ce.get("marker", "")
                        idx = cur.find(marker) if marker else -1
                        if idx >= 0:
                            new = cur[:idx] + val
                    elif op == "replace_price":         # 治療費の①②…行を「キット×治療費」ブロックへ置換
                        lines = new.split("\n")
                        out = []; done = False
                        for ln in lines:
                            if "を用いた治療" in ln:      # 既存の価格行（サンプル/トークン問わず）
                                if not done:
                                    out.append(val); done = True
                            else:
                                out.append(ln)
                        new = "\n".join(out) if done else new
                    if new != cur and TX.safe_set(ws, ce["cell"], new):
                        written.add(ce["cell"])
                        n_x_done += 1
                        log_rows.append(TX.make_row(run_dt, doc_key,
                                                    {"sheet": exc.get("sheet"), "cell": ce["cell"],
                                                     "var": ce.get("desc", "cell_edit")},
                                                    val, TX.ST_CHECK, ce.get("desc")))
                # 転記したセルを緑にマーク／元が緑だが未転記のセルはリストへ（黒のまま）
                for coord in written:
                    try:
                        _set_cell_color(ws[coord], GREEN)
                    except Exception:
                        pass
                for coord, ov in orig_green.items():
                    if coord not in written:
                        green_report.append((doc_key, sheet, coord, ov))
                # 部分セル・旧トークンの {{…}} をヒアリング値で穴埋め（clear_tokensの前に）
                _saisei = 2 if "3種" in doc_key else 1
                _fill_xlsx_tokens(ws, hearing, kits, TX, _saisei)
                TX.clear_tokens(wb)   # なお残る未充足 {{…}} を空欄化（例: 3種様式の{{医師2_氏名}}）
                # 変更セルの差分を算出し、XML手術で書込み（openpyxl保存を避け図形を保持）
                diff = {}
                for r in ws.iter_rows():
                    for c in r:
                        if orig_vals.get(c.coordinate) != c.value:
                            diff[c.coordinate] = c.value
                try:
                    _write_xlsx_cells_xmlsurgery(xpath, ws.title, diff, run_log)
                except Exception as e:
                    run_log.append("[%s] XML手術失敗→openpyxl保存にフォールバック（図形が失われる可能性）: %r"
                                   % (doc_key, e))
                    wb.save(xpath)
            else:
                run_log.append("[%s] 様式xlsx無し: %s" % (doc_key, exc["file"]))

        # --- 直下の各docx（一括置換） ---
        n_files = n_repl = n_kit = 0
        if Document and _replace_in_paragraph:
            pairs = _build_docx_pairs(doc.get("docx", []), hearing, kits, TX)
            # キット製造方法トークン（{{採取方法}}等 → ①②③連番の複数行本文）
            kit_tokens = []
            for kt in doc.get("docx_kit", []):
                val, _ = TX.resolve(kt.get("source", {}), hearing, kits)
                kit_tokens.append((kt["token"], val or ""))
            for f in sorted(glob.glob(os.path.join(out_folder, "*.docx"))):
                fbase = os.path.basename(f)
                if fbase.startswith("~$"):
                    continue
                try:
                    d = Document(f)
                except Exception as ex:
                    run_log.append("[%s] docx読込失敗 %s: %r" % (doc_key, fbase, ex))
                    continue
                reset_doc_green_to_black(d)   # テンプレの緑を一旦黒へ（転記箇所だけ後で緑になる）
                # file_contains 指定があるものは対象ファイルのみに適用
                simple_pairs = [(fnd, v) for fnd, v, _, fc in pairs if (not fc or fc in fbase)]
                cnt = kc = 0
                for p in _iter_all_paragraphs(d):
                    cnt += _replace_in_paragraph(p, simple_pairs)
                cnt += _replace_lowlevel(d, simple_pairs)   # テキストボックス等も置換
                for tok, val in kit_tokens:                 # キットトークン差し込み
                    kc += _fill_token_multiline(d, tok, val)
                # 表フィル（略歴書など：行位置ベースでヒアリング略歴書シートから転記）
                for spec in doc.get("docx_table", []):
                    if spec.get("file_contains", "") not in fbase:
                        continue
                    ti = spec.get("table", 0); vcol = spec.get("value_col", 1)
                    tables = d.tables
                    if ti >= len(tables):
                        continue
                    tbl = tables[ti]
                    for rowdef in spec.get("rows", []):
                        ri = rowdef.get("row")
                        if ri is None or ri >= len(tbl.rows):
                            continue
                        val2, st2 = TX.resolve(rowdef.get("source", {}), hearing, kits)
                        if val2 in (None, ""):
                            continue
                        cells = tbl.rows[ri].cells
                        if vcol < len(cells):
                            try:
                                _set_cell_multiline(cells[vcol], val2)
                                kc += 1
                            except Exception as ex:
                                run_log.append("[docs] %s 表フィル失敗 r%s: %r" % (fbase, ri, ex))
                # 前世代を含む任意の {{項目名}} を汎用リゾルバで穴埋め（取りこぼし防止）
                _saisei_d = 2 if "3種" in doc_key else 1
                cnt += _fill_all_docx_tokens(d, hearing, kits, _saisei_d, TX, _iter_all_paragraphs)
                # 穴埋め後、なお残る未充足 {{...}} を空欄化（トークン方式の後始末）
                leftover = _clear_doc_tokens(d, _iter_all_paragraphs)
                if cnt or kc or leftover:
                    d.save(f)
                    n_files += 1
                    n_repl += cnt
                    n_kit += kc
            # 置換サマリをログ行に
            for fnd, v, desc, fc in pairs:
                log_rows.append(TX.make_row(run_dt, doc_key,
                                            {"sheet": "(docx一括置換)", "cell": fnd, "var": desc},
                                            v, TX.ST_DONE, "サンプル値→ヒアリング値"))

        out_folders.append(out_folder)
        run_log.append("[%s] 出力フォルダ: %s ／ 様式転記=%d(確認%d) ／ docx置換=%dファイル・%d件 ／ キット差込=%d"
                       % (doc_key, os.path.basename(out_folder),
                          n_x_done, n_x_check, n_files, n_repl, n_kit))

    # 緑ターゲットだが未転記だったセルをレポート出力（要確認）
    if green_report:
        rep = os.path.join(dir_output, "_緑なのに未転記セル一覧.txt")
        try:
            with io.open(rep, "w", encoding="utf-8") as g:
                g.write("緑（転記ターゲット）だが今回転記されなかったセル一覧\n")
                g.write("※ 転記漏れの可能性。不要な緑なら無視可。出力側では黒に戻しています。\n\n")
                for doc_key2, sh, coord, ov in green_report:
                    g.write("[%s] %s!%s  現値=%r\n" % (doc_key2, sh, coord, ov))
            run_log.append("[緑チェック] 未転記の緑セル %d件 → %s" % (len(green_report), os.path.basename(rep)))
        except Exception as e:
            run_log.append("[緑チェック] レポート出力失敗: %r" % e)
    else:
        run_log.append("[緑チェック] 未転記の緑セルなし")
    return out_folders, log_rows


if __name__ == "__main__":
    import datetime
    from transcribe import Hearing, find_hearing, DIR_TPL, DIR_OUTPUT
    hp = find_hearing()
    hearing = Hearing(hp, "ヒアリングシート（PRP）")
    log = []
    outs, _ = run_docs(hearing, hp, DIR_TPL, DIR_OUTPUT, log, datetime.datetime.now())
    print("=== フォルダ型転記 完了 ===")
    print("入力 :", hp)
    for o in outs:
        print("出力 :", o)
    for line in log:
        print("  -", line)

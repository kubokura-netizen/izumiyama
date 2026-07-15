# -*- coding: utf-8 -*-
"""
e-再生医療（https://saiseiiryo.mhlw.go.jp/…）等のWebフォームへ、ヒアリングシートの値を
半自動で入力する補助ツール（方式B：ブラウザ起動→手動ログイン→自動入力→人が確認・送信）。

※ 送信は行いません。入力（下書き）までを自動化し、最終確認と送信は必ず人が行ってください。
※ 政府ポータルの自動操作は利用規約の確認が前提です。ログインは人が手動で行います。

使い方:
  1) 準備（初回のみ）:
       pip install playwright
       playwright install chromium
  2) フォーム項目の抽出（対応表を作るため）:
       py 99_data/src/web_fill.py --dump
     → ブラウザが開くのでログインし、対象フォームを表示 → コンソールでEnter
     → 03_logs/web_fields_dump.txt に入力欄の一覧（selector候補・ラベル）が出力される
  3) 対応表を作る:
       03_logs/web_fields_dump.txt を見ながら 99_data/マッピング/web_mapping.json の
       fields[].selector と source を埋める（source書式は docs_config と同じ）
  4) 自動入力:
       py 99_data/src/web_fill.py
     → ログイン→フォーム表示→Enter で各欄に自動入力（緑枠でハイライト）→ 人が確認して送信
"""
import os, sys, io, json, glob, datetime

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.dirname(SRC_DIR)
BASE = os.path.dirname(DATA)
sys.path.insert(0, SRC_DIR)

import transcribe as TX   # Hearing / resolve を再利用

MAPPING = os.path.join(DATA, "マッピング", "web_mapping.json")
PROFILE = os.path.join(DATA, "_web_profile")          # ログイン維持用プロファイル
LOGDIR = TX.resolve_dir("03_logs", create=True)


def find_hearing():
    """01_input の最新ヒアリング(.xlsx)を返す（web_fill 独自。--flag を誤検出しない）。"""
    d = TX.resolve_dir("01_input")
    cands = [c for c in glob.glob(os.path.join(d, "*.xlsx"))
             if not os.path.basename(c).startswith("~$")]
    cands.sort(key=os.path.getmtime, reverse=True)
    if not cands:
        raise FileNotFoundError("01_input にヒアリングシート(.xlsx)を入れてください")
    return cands[0]


def load_mapping():
    if not os.path.exists(MAPPING):
        return {}
    with io.open(MAPPING, encoding="utf-8") as f:
        return json.load(f)


# フォーム上の全入力欄を抽出するJS（name/id/type/ラベル/選択肢）
# ラジオ/チェックは name でグルーピングし、各選択肢のラベル文言（有/無・該当/非該当等）も収集。
DUMP_JS = r"""
() => {
  const optText = (el) => {
    // ラジオ/チェックの選択肢文言（隣接ラベル or 直後テキスト）
    if (el.id) { const l = document.querySelector('label[for="'+CSS.escape(el.id)+'"]'); if (l) return l.innerText.trim(); }
    let p = el.closest('label'); if (p) return p.innerText.trim();
    let sib = el.nextElementSibling; if (sib && sib.innerText) return sib.innerText.trim().slice(0,20);
    if (el.value) return el.value;
    return '';
  };
  const questionText = (el) => {
    // 設問ラベル（表の左セル/直近の見出し的テキスト）
    let td = el.closest('td,th'); if (td && td.previousElementSibling) return (td.previousElementSibling.innerText||'').trim().slice(0,60);
    let row = el.closest('tr,div'); if (row) { const t=(row.innerText||'').trim().split('\n')[0]; if (t) return t.slice(0,60); }
    if (el.getAttribute('aria-label')) return el.getAttribute('aria-label').trim();
    return '';
  };
  const sel = (el) => {
    if (el.id) return '#'+CSS.escape(el.id);
    if (el.name) return el.tagName.toLowerCase()+'[type="'+(el.type||'')+'"][name="'+el.name+'"]';
    return el.tagName.toLowerCase();
  };
  const seenGroup = {};
  const out = [];
  document.querySelectorAll('input, select, textarea').forEach(el => {
    const t = (el.type||el.tagName).toLowerCase();
    if (['hidden','submit','button','image','reset'].includes(t)) return;
    if (t==='radio' || t==='checkbox') {
      const key = t+':'+(el.name||el.id);
      if (seenGroup[key]) { seenGroup[key].choices.push(optText(el)); return; }
      const rec = {tag:'input', type:t, name:el.name||'', id:el.id||'',
                   selector: el.name? 'input[type="'+t+'"][name="'+el.name+'"]' : sel(el),
                   label: questionText(el), placeholder:'', choices:[optText(el)]};
      seenGroup[key]=rec; out.push(rec); return;
    }
    let opts = '';
    if (el.tagName.toLowerCase()==='select') opts = Array.from(el.options).map(o=>o.text.trim()).join(' | ');
    out.push({tag: el.tagName.toLowerCase(), type: t, name: el.name||'', id: el.id||'',
              selector: sel(el), label: questionText(el), placeholder: el.placeholder||'', options: opts, choices:[]});
  });
  return out;
}
"""


def dump_fields(page):
    fields = page.evaluate(DUMP_JS)
    path = os.path.join(LOGDIR, "web_fields_dump.txt")
    with io.open(path, "w", encoding="utf-8") as f:
        f.write("Webフォーム 入力欄ダンプ  (%s)\n" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        f.write("URL: %s\n" % page.url)
        f.write("この selector と label を見て web_mapping.json を作成してください。\n\n")
        for i, fl in enumerate(fields, 1):
            f.write("%2d) [%s/%s] selector=%s\n" % (i, fl["tag"], fl["type"], fl["selector"]))
            if fl.get("label"):       f.write("     設問: %s\n" % fl["label"])
            if fl.get("placeholder"): f.write("     placeholder: %s\n" % fl["placeholder"])
            if fl.get("options"):     f.write("     選択肢(select): %s\n" % fl["options"])
            if fl.get("choices"):     f.write("     選択肢(radio/check): %s\n" % " / ".join([x for x in fl["choices"] if x]))
    print("→ 入力欄 %d件を出力: %s" % (len(fields), path))


import re as _re


def _pref_split(addr):
    """住所文字列を (都道府県, 残り) に分割。"""
    addr = TX.clean(addr)
    m = _re.match(r"(北海道|東京都|京都府|大阪府|.{1,3}?[県府])", addr)
    pref = m.group(1) if m else ""
    return pref, (addr[len(pref):] if pref else addr)


def load_output_ctx(mapping):
    """02_output の最新の出力フォルダ（output_folder_contains で絞る）を探し、
       様式xlsxシート ws とフォルダパス folder を返す。Webの参照元＝アウトプット。
       無ければ (None, '')。"""
    import openpyxl
    outdir = TX.resolve_dir("02_output")
    want = mapping.get("output_folder_contains", "2種関節系PRP")
    folders = [d for d in glob.glob(os.path.join(outdir, "*"))
               if os.path.isdir(d) and want in os.path.basename(d)]
    if not folders:
        return None, ""
    folder = max(folders, key=os.path.getmtime)
    ws = None
    xls = glob.glob(os.path.join(folder, "01.*.xlsx"))
    if xls:
        wb = openpyxl.load_workbook(xls[0], data_only=True)
        sh = mapping.get("output_sheet", "1 の２提供計画（治療）")
        ws = wb[sh] if sh in wb.sheetnames else wb[wb.sheetnames[0]]
    return ws, folder


def _iter_block_texts(doc):
    """docxを文書順（段落＋表）で走査し、各ブロックのテキストを返す。表は行を改行/セルを｜で連結。"""
    try:
        from docx.text.paragraph import Paragraph
        from docx.table import Table
    except Exception:
        for p in doc.paragraphs:
            yield p.text
        return
    for child in doc.element.body.iterchildren():
        tag = child.tag
        if tag.endswith("}p"):
            yield Paragraph(child, doc).text
        elif tag.endswith("}tbl"):
            tbl = Table(child, doc)
            rows = []
            for row in tbl.rows:
                cells = [c.text.strip() for c in row.cells]
                rows.append(" | ".join([c for c in cells if c]))
            yield "\n".join([r for r in rows if r.strip()])


def _docx_section(out_folder, file_contains, heading, until=None):
    """出力フォルダ内のWord（file_containsで特定）から、heading の次ブロック〜until手前 の本文を抽出。
       段落だけでなく表も文書順で拾う（適格性基準など表本文も取りこぼさない）。"""
    try:
        from docx import Document
    except Exception:
        return ""
    files = [f for f in glob.glob(os.path.join(out_folder, "*.docx"))
             if file_contains and file_contains in os.path.basename(f)]
    if not files:
        return ""
    doc = Document(files[0])
    blocks = list(_iter_block_texts(doc))
    start = None
    for i, t in enumerate(blocks):
        if heading and heading in t:
            start = i + 1
            break
    if start is None:
        return ""
    body = []
    for t in blocks[start:]:
        if until and t.strip().startswith(until):
            break
        body.append(t)
    return "\n".join(x for x in body if x.strip()).strip()


def _resolve_value(fld, hearing, kits, out_ws=None, out_folder=""):
    """フィールドの入力値を決める。★出力(アウトプット)を最優先：
       cell=様式xlsxのセル / docx=出力Wordの見出しセクション。無ければヒアリング(source)。
       web専用type: pref/addr_body/today(year|month|day)。cell_tf: pref/addr_body でセル値を分割。"""
    # ① アウトプット様式xlsxのセル優先
    cell = fld.get("cell")
    if cell and out_ws is not None:
        raw = TX.clean(out_ws[cell].value)
        tf = fld.get("cell_tf")
        if tf in ("pref", "addr_body"):
            pref, body = _pref_split(raw)
            if pref or body:
                return pref if tf == "pref" else body
        elif tf == "zip":
            z = raw.lstrip("〒 　").strip()
            if z:
                return z
        elif raw:
            return raw
    # ② アウトプットWordのセクション
    dx = fld.get("docx")
    if dx and out_folder:
        v = TX.clean(_docx_section(out_folder, dx.get("file_contains", ""),
                                   dx.get("heading", ""), dx.get("until")))
        if v:
            return v
    # ③ ヒアリング（フォールバック）
    src = fld.get("source")
    if src:
        t = src.get("t")
        if t in ("pref", "addr_body"):
            inner = src.get("source") or {"t": "hearing",
                    "label": "医療機関/住所（診療所開設届上）", "section": "法人/医療機関"}
            addr, _ = TX.resolve(inner, hearing, kits)
            pref, body = _pref_split(addr)
            return pref if t == "pref" else body
        if t == "today" and src.get("fmt") in ("year", "month", "day"):
            import datetime
            d = datetime.datetime.now()
            return {"year": str(d.year), "month": "%02d" % d.month, "day": str(d.day)}[src.get("fmt")]
        v, _ = TX.resolve(src, hearing, kits)
        v = TX.clean(v)
        # トークン名指定（採血量・治療価格・キット製造方法等の汎用解決）
        if not v and t == "token":
            nm = src.get("name", "")
            v = TX.clean(TX.resolve_token_by_name(nm, hearing, kits, fld.get("saisei", 1)))
            if not v and nm in ("採取方法", "加工方法", "投与方法"):   # 採取の非整合を補完
                part = {"採取方法": "採取", "加工方法": "加工", "投与方法": "投与"}[nm]
                v = TX.clean(hearing.kit_block(part, numbered=True))
        return v
    return TX.clean(fld.get("value", ""))


def _locate(page, fld):
    """selector優先、無ければ label（アクセシブルラベル）で要素を特定。"""
    selector = (fld.get("selector") or "").strip()
    if selector:
        return page.locator(selector).first
    label = (fld.get("label") or "").strip()
    if label:
        try:
            loc = page.get_by_label(label, exact=False).first
            if loc.count() > 0:
                return loc
        except Exception:
            pass
        return page.get_by_text(label, exact=False).first
    return None


def fill_one_page(page, mapping, hearing, kits, filled_keys, out_ws=None, out_folder=""):
    """現在表示中のページの、まだ埋めていないフィールドを入力する。"""
    done = 0
    logs = []
    for idx, fld in enumerate(mapping.get("fields", [])):
        if not isinstance(fld, dict):
            continue                              # 説明用の文字列などはスキップ
        key = fld.get("desc", "") + "#" + str(idx)
        if key in filled_keys:
            continue
        if not (fld.get("selector") or fld.get("label")):
            continue
        val = _resolve_value(fld, hearing, kits, out_ws, out_folder)
        desc = fld.get("desc", fld.get("selector") or fld.get("label"))
        ftype = (fld.get("type") or "text").lower()
        try:
            if ftype == "radio":
                # 選ぶ選択肢ラベル＝ choice(固定) or 値そのもの（有/無・該当/非該当）
                choice = fld.get("choice") or val
                grp = (fld.get("selector") or "").strip()
                if not choice or not grp:
                    continue
                # グループ内(name一致)の各ラジオのラベル文言を見て、choiceに一致するものだけをチェック
                radios = page.locator(grp)
                cnt = radios.count()
                picked = False
                for i in range(cnt):
                    r = radios.nth(i)
                    lbl = r.evaluate(
                        "el => { if(el.id){const l=document.querySelector('label[for=\"'+el.id+'\"]'); if(l) return l.innerText.trim();}"
                        " let p=el.closest('label'); if(p) return p.innerText.trim();"
                        " let s=el.nextElementSibling; if(s&&s.innerText) return s.innerText.trim();"
                        " return el.value||''; }")
                    if choice and (choice == (lbl or "").strip() or choice in (lbl or "")):
                        r.check()
                        r.evaluate("el => { const l=el.closest('label')||el.parentElement; if(l){l.style.outline='2px solid #00B050';} }")
                        picked = True
                        break
                if not picked:
                    try:
                        page.locator("%s[value='%s']" % (grp, choice)).first.check(); picked = True
                    except Exception:
                        pass
                if picked:
                    logs.append("選択(radio): %s → %s" % (desc, choice))
                    filled_keys.add(key); done += 1
                else:
                    logs.append("radio未選択(候補なし): %s → %s" % (desc, choice))
                continue
            loc = _locate(page, fld)
            if loc is None or loc.count() == 0:
                continue                         # このページには無い（別セクションかも）→スキップ
            if not val:
                continue
            tag = loc.evaluate("el => el.tagName.toLowerCase()")
            if tag == "select" or ftype == "select":
                try:
                    loc.select_option(label=val)
                except Exception:
                    loc.select_option(value=val)
            elif ftype == "checkbox":
                (loc.check() if str(val) not in ("", "0", "false", "無", "×") else loc.uncheck())
            else:
                loc.fill(str(val))
            loc.evaluate("el => { el.style.outline='2px solid #00B050'; el.style.background='#eaffea'; }")
            logs.append("入力: %s = %s" % (desc, str(val)[:40]))
            filled_keys.add(key); done += 1
        except Exception as e:
            logs.append("失敗: %s %r" % (desc, e))
    for l in logs:
        print("  -", l)
    return done


def fill_fields(page, mapping, hearing, kits):
    """複数セクション対応：現ページを入力→ユーザーが次セクションへ→また入力…を繰り返す。"""
    out_ws, out_folder = load_output_ctx(mapping)
    if out_folder:
        print("参照元(アウトプット): %s ／ 様式xlsx=%s"
              % (os.path.basename(out_folder), "あり" if out_ws is not None else "なし"))
    else:
        print("※ アウトプット未検出 → ヒアリングから取得します（先に 転記実行.bat を実行推奨）")
    filled = set()
    total = 0
    while True:
        print("\n=== このページを自動入力 ===")
        n = fill_one_page(page, mapping, hearing, kits, filled, out_ws, out_folder)
        total += n
        print("  （このページで %d件入力・累計 %d件）" % (n, total))
        print("★ 内容を確認してください。ツールは送信しません。")
        ans = input("次のセクションを表示したら Enter（続けて入力）／ 終了は q + Enter : ").strip().lower()
        if ans == "q":
            break
    print("\n合計 %d件を入力しました。最終確認のうえ、画面から送信してください。" % total)


def main():
    mode_dump = "--dump" in sys.argv
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright が未導入です。次を実行してください:")
        print("  pip install playwright")
        print("  playwright install chromium")
        return

    mapping = load_mapping()
    url = mapping.get("url") or "https://saiseiiryo.mhlw.go.jp/"
    hearing_sheet = mapping.get("hearing_sheet", "ヒアリングシート（PRP）")

    hearing = kits = None
    if not mode_dump:
        hp = find_hearing()
        hearing = TX.Hearing(hp, hearing_sheet)
        kits = hearing.prp_kits()
        print("ヒアリング:", os.path.basename(hp))

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(PROFILE, headless=False,
                                                    viewport={"width": 1280, "height": 900})
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
        print("\n▼ ブラウザが開きました（このサイトはログイン不要）。")
        if mode_dump:
            input("  対象フォームを表示したら、このウィンドウで Enter（入力欄を抽出します）…")
            dump_fields(page)
        else:
            input("  plan01フォームを表示したら、このウィンドウで Enter（自動入力を開始）…")
            print("  ※ 複数ページのフォームは、入力後に「次の項目へ」で次を表示→Enter で続けて入力できます。")
            fill_fields(page, mapping, hearing, kits)
        ctx.close()


if __name__ == "__main__":
    main()

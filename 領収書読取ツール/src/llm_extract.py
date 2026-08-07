# -*- coding: utf-8 -*-
"""
Ollama（ローカル画像対応LLM）で領収書画像 → 項目JSON を抽出する部品。
  ・完全ローカル/オフライン（http://127.0.0.1:11434）。API課金なし。
  ・Tesseract(OCR)のルール方式が苦手な「どの数字が合計か」等を、画像を見て判断させる。
"""
import os
import io
import re
import json
import base64
import urllib.request

OLLAMA_URL = "http://127.0.0.1:11434/api/generate"
DEFAULT_MODEL = os.environ.get("RECEIPT_LLM_MODEL", "qwen2.5vl:3b")

PROMPT = (
    "画像は日本の領収書・レシート・納付書です。下記キーだけのJSONを返す。JSON以外は出力しない。\n"
    "{\n"
    '  "日付": "YYYY-MM-DD",\n'
    '  "相手先": "店名・発行者名",\n'
    '  "金額": 0,\n'
    '  "内容": "但し書き/品目の短い要約",\n'
    '  "種別": "下記から1つ",\n'
    '  "手書き": false\n'
    "}\n"
    "【金額】最重要。領収の合計金額（税込）を必ず整数で入れる。『合計』『計』『¥』『￥』の直後の数字。\n"
    "  例: ¥250→250、¥1,000→1000、合計 ¥26,777→26777。電話番号・受付番号・No・年月日は絶対に金額にしない。\n"
    "  金額が本当に読めない時だけ 0。\n"
    "【種別】次から最も近い1語だけ: 印紙代/証明書代/交通費/消耗品費/通信費/租税公課/社会保険料/その他\n"
    "【手書き】紙面の本文が手書き文字中心のときだけ true。印刷レシートに赤い印やメモがあっても false。\n"
    "【日付】領収日。令和は西暦に直す(令和7年=2025)。読めなければ \"\"。"
)


def _shrink(image_path, max_side=1400):
    """大きすぎる切り出しはCPU推論が遅いので長辺を縮小してPNGバイト列を返す。"""
    try:
        from PIL import Image
        im = Image.open(image_path)
        w, h = im.size
        s = max(w, h)
        if s > max_side:
            im = im.resize((int(w * max_side / s), int(h * max_side / s)))
        buf = io.BytesIO()
        im.convert("RGB").save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        with open(image_path, "rb") as f:
            return f.read()


def extract(image_path, model=DEFAULT_MODEL, timeout=180):
    """画像1枚 → dict（日付/相手先/金額/内容/種別/手書き）。失敗時は {'_error':...}。"""
    b64 = base64.b64encode(_shrink(image_path)).decode("ascii")
    payload = {
        "model": model, "prompt": PROMPT, "images": [b64],
        "stream": False, "format": "json", "options": {"temperature": 0},
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(OLLAMA_URL, data=data,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            resp = json.loads(r.read().decode("utf-8"))
    except Exception as e:
        return {"_error": "Ollama呼び出し失敗: %r" % e}
    text = resp.get("response", "") or ""
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
        return {"_error": "JSON解析失敗", "_raw": text[:300]}


def ready(model=DEFAULT_MODEL, timeout=5):
    """Ollamaサーバーが応答し、モデルが入っているか。"""
    try:
        req = urllib.request.Request("http://127.0.0.1:11434/api/tags")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            tags = json.loads(r.read().decode("utf-8"))
        names = [m.get("name", "") for m in tags.get("models", [])]
        base = model.split(":")[0]
        return any(base in n for n in names), names
    except Exception as e:
        return False, ["(サーバー未応答: %r)" % e]


if __name__ == "__main__":
    import sys
    ok, names = ready()
    print("Ollama:", "OK" if ok else "NG", "／ models:", names)
    for p in sys.argv[1:]:
        print("---", os.path.basename(p))
        print(json.dumps(extract(p), ensure_ascii=False, indent=2))

#!/bin/bash
# ============================================================
#  PRP 自動転記ツール ダッシュボード 起動（Mac）
#   このファイルをダブルクリックすると、
#   必要な準備（初回のみ）→ サーバ起動 → ブラウザ表示 まで自動。
# ============================================================
cd "$(dirname "$0")"
DASH="98_dashboard"
VENV=".venv_dashboard"

# python3 の存在確認
if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 が見つかりません。https://www.python.org からインストールしてください。"
  read -n 1 -s -r -p "何かキーを押すと閉じます"
  exit 1
fi

# 初回のみ: 仮想環境を作って必要ライブラリを入れる
if [ ! -d "$VENV" ]; then
  echo "初回セットアップ中（1〜2分）… 必要なライブラリを準備します。"
  python3 -m venv "$VENV"
  "$VENV/bin/python" -m pip install --upgrade pip >/dev/null
  "$VENV/bin/python" -m pip install -r "$DASH/requirements.txt"
fi

echo ""
echo "ダッシュボードを起動します。ブラウザが自動で開きます。"
echo "終了するには、この黒い画面で Control + C を押してください。"
echo ""
"$VENV/bin/python" "$DASH/app.py"

# -*- coding: utf-8 -*-
"""
フォルダ設定（settings.json）を対話で作成/更新する。
相手の環境（読取フォルダ・保存先・経費入力表.xlsx）に合わせるために使う。
"""
import os
import pipeline as P


def ask(label, cur):
    print("\n%s" % label)
    if cur:
        print("  現在: %s" % cur)
    v = input("  入力（空Enter=変更なし）> ").strip().strip('"')
    return v or cur


def main():
    cfg = P.load_settings()
    print("=== 領収書ツール フォルダ設定 ===")
    print("（空Enterでその項目は変更しません）")

    input_dir = ask("① 読取フォルダ（PDFを入れる場所）※通常は 01_input のままでOK",
                    cfg.get("input_dir"))
    output_root = ask("② 保存先（年月フォルダと経費入力表を置く親フォルダ）\n"
                      "   例: C:\\Users\\xxx\\Desktop\\領収証データ",
                      cfg.get("output_root"))
    master = ask("③ 経費入力表.xlsx のフルパス（蓄積先）\n"
                 "   例: C:\\Users\\xxx\\Desktop\\領収証データ\\経費入力表.xlsx",
                 cfg.get("master_xlsx"))
    act = ask("④ 原本の扱い  copy=コピー(原本を残す) / move=移動  [copy推奨]",
              cfg.get("file_action") or "copy")
    act = "move" if str(act).lower().startswith("m") else "copy"

    save = dict(cfg)
    save["input_dir"] = input_dir
    save["output_root"] = output_root
    save["master_xlsx"] = master
    save["file_action"] = act
    P.save_settings(save)

    print("\n保存しました → %s" % P.SETTINGS_PATH)
    eff = P.load_settings()
    print("  読取フォルダ : %s" % eff["input_dir"])
    print("  保存先       : %s" % eff["output_root"])
    print("  経費入力表   : %s" % eff["master_xlsx"])
    print("  原本の扱い   : %s" % eff["file_action"])


if __name__ == "__main__":
    main()

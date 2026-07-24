# -*- coding: utf-8 -*-
"""
to_pdf.py ― Word/Excel を「白黒(グレースケール)PDF」に変換する部品。

流れ:
  ① Word/Excel → PDF（MS Office を COM 自動化して変換。要 pywin32）
  ② PDF を白黒化：Ghostscript があればそれ（ベクター保持=高品質・軽量）、
                  無ければ PyMuPDF で代替（ページを画像化してグレースケール）。

※ Office が必要（Word/Excel がインストールされていること）。
※ 送信・申請などは一切しない。ファイルを変換するだけ。
"""
import os
import glob
import shutil
import subprocess


def check_ready():
    """(ok, メッセージ)。PDF変換に必要な部品(pywin32/PyMuPDF)が揃っているか。"""
    import importlib.util as u
    miss = []
    if not u.find_spec("win32com"):
        miss.append("pywin32")
    if not u.find_spec("fitz"):
        miss.append("PyMuPDF")
    if miss:
        return False, ("PDF変換に必要な部品が未導入です（%s）。web転記実行.bat の『3: Setup』"
                       "またはダッシュボードの『初回準備』を実行してください。" % " / ".join(miss))
    return True, "OK"


def find_ghostscript():
    """Ghostscript(gswin64c 等)のパスを返す。無ければ ""。PATH→Program Files の順に探す。"""
    for name in ("gswin64c", "gswin64", "gswin32c", "gs"):
        p = shutil.which(name)
        if p:
            return p
    pats = [r"C:\Program Files\gs\*\bin\gswin64c.exe",
            r"C:\Program Files (x86)\gs\*\bin\gswin32c.exe",
            r"C:\Program Files\gs\*\bin\gswin32c.exe"]
    for pat in pats:
        hits = glob.glob(pat)
        if hits:
            return sorted(hits)[-1]
    return ""


# ---- ① Office → PDF -------------------------------------------------------
def _word_to_pdf(src, pdf):
    import win32com.client as win32
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(src), ReadOnly=True)
        try:
            doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)   # 17 = wdExportFormatPDF
        finally:
            doc.Close(False)
    finally:
        word.Quit()


def _excel_to_pdf(src, pdf):
    import win32com.client as win32
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(src), ReadOnly=True)
        try:
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf))     # 0 = xlTypePDF
        finally:
            wb.Close(False)
    finally:
        xl.Quit()


def office_to_pdf(src, pdf):
    ext = os.path.splitext(src)[1].lower()
    if ext in (".docx", ".doc", ".docm"):
        _word_to_pdf(src, pdf)
    elif ext in (".xlsx", ".xls", ".xlsm"):
        _excel_to_pdf(src, pdf)
    else:
        raise ValueError("PDF変換に未対応の拡張子: %s" % ext)
    return pdf


# ---- ② PDF → 白黒 ---------------------------------------------------------
def _grayscale_gs(gs, src_pdf, dst_pdf):
    subprocess.run([gs, "-q", "-dBATCH", "-dNOPAUSE", "-dSAFER",
                    "-sDEVICE=pdfwrite",
                    "-sProcessColorModel=DeviceGray",
                    "-sColorConversionStrategy=Gray",
                    "-dOverrideICC=true",
                    "-o", os.path.abspath(dst_pdf), os.path.abspath(src_pdf)],
                   check=True)


def _grayscale_pymupdf(src_pdf, dst_pdf, dpi=200):
    import fitz
    doc = fitz.open(src_pdf)
    out = fitz.open()
    try:
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0),
                                  colorspace=fitz.csGRAY)
            np = out.new_page(width=page.rect.width, height=page.rect.height)
            np.insert_image(np.rect, pixmap=pix)
        out.save(dst_pdf, deflate=True, garbage=3)
    finally:
        out.close()
        doc.close()


def grayscale_pdf(src_pdf, dst_pdf):
    """PDFを白黒化。Ghostscriptがあればそれ、無ければPyMuPDF。使った方式名を返す。"""
    gs = find_ghostscript()
    if gs:
        _grayscale_gs(gs, src_pdf, dst_pdf)
        return "ghostscript"
    _grayscale_pymupdf(src_pdf, dst_pdf)
    return "pymupdf(raster)"


# ---- まとめ ---------------------------------------------------------------
def convert(src, dst_pdf, gray=True):
    """src(Word/Excel) → dst_pdf。gray=Trueで白黒化。
       戻り値: (出力パス, 使った白黒化方式 or 'color')。"""
    os.makedirs(os.path.dirname(os.path.abspath(dst_pdf)) or ".", exist_ok=True)
    if not gray:
        office_to_pdf(src, dst_pdf)
        return dst_pdf, "color"
    color = dst_pdf + ".__color.pdf"
    office_to_pdf(src, color)
    try:
        method = grayscale_pdf(color, dst_pdf)
    finally:
        try:
            os.remove(color)
        except OSError:
            pass
    return dst_pdf, method

"""
PDF読み込みスクリプト
使い方：
  python3 read_pdf.py 元資料.pdf

PDFの内容を解析してテキスト・構造を出力します。
その後 generate_roofplus_report.py の内容に反映します。
"""

import sys
import fitz  # PyMuPDF
import pdfplumber

def extract_with_pdfplumber(pdf_path):
    """テキスト・表を抽出"""
    print("=" * 60)
    print(f"PDF: {pdf_path}")
    print("=" * 60)

    with pdfplumber.open(pdf_path) as pdf:
        print(f"総ページ数: {len(pdf.pages)}\n")

        for i, page in enumerate(pdf.pages):
            print(f"\n{'─'*40}")
            print(f"【ページ {i+1}】")
            print(f"{'─'*40}")

            # テキスト抽出
            text = page.extract_text()
            if text:
                print(text)

            # 表の抽出
            tables = page.extract_tables()
            if tables:
                print(f"\n[表 {len(tables)}個 検出]")
                for j, table in enumerate(tables):
                    print(f"\n  表{j+1}:")
                    for row in table:
                        print("  | " + " | ".join(str(c or "") for c in row) + " |")

def extract_images(pdf_path):
    """各ページをPNG画像として保存（グラフ確認用）"""
    doc = fitz.open(pdf_path)
    print(f"\n\n画像として保存中...")
    for i, page in enumerate(doc):
        mat = fitz.Matrix(2.0, 2.0)  # 2倍解像度
        pix = page.get_pixmap(matrix=mat)
        img_path = f"pdf_page_{i+1:02d}.png"
        pix.save(img_path)
        print(f"  保存: {img_path}")
    print("完了。pdf_page_XX.png を確認してください。")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使い方: python3 read_pdf.py <PDFファイルのパス>")
        print("例:     python3 read_pdf.py /home/user/RoofPlus元資料.pdf")
        sys.exit(1)

    pdf_path = sys.argv[1]

    # テキスト・表の抽出
    extract_with_pdfplumber(pdf_path)

    # 各ページを画像として保存
    extract_images(pdf_path)

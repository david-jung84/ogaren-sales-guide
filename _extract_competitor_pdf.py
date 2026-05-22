"""슬립퍼 매트리스 3.3.3 경쟁사 비교자료 PDF에서 텍스트 + 이미지 추출"""
import fitz
import sys
from pathlib import Path

PDF = r"C:\Users\USER\Downloads\[슬립퍼 매트리스 3.3.3. 경쟁사 비교자료] ver.251212.pdf"
OUT_DIR = Path(r"C:\Users\USER\Desktop\오가렌_DX\세일즈가이드\appendix_pdf_extract")
OUT_DIR.mkdir(exist_ok=True)

doc = fitz.open(PDF)
print(f"Pages: {doc.page_count}")
print(f"Output dir: {OUT_DIR}")

# 페이지별 텍스트 + PNG 렌더링
for i in range(doc.page_count):
    page = doc[i]
    # 텍스트
    text = page.get_text()
    (OUT_DIR / f"page_{i+1:02d}.txt").write_text(text, encoding="utf-8")
    # PNG (200 DPI)
    mat = fitz.Matrix(200/72, 200/72)
    pix = page.get_pixmap(matrix=mat)
    pix.save(str(OUT_DIR / f"page_{i+1:02d}.png"))
    print(f"page {i+1:02d}: {len(text)} chars, {pix.width}x{pix.height}")

doc.close()
print("done")

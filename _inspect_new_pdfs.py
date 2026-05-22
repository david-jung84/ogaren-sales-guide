# -*- coding: utf-8 -*-
"""3개 신규 PDF 텍스트 + 페이지별 첫 줄 추출."""
import fitz
import sys, io
from pathlib import Path
import json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PDFS = [
    (r"C:\Users\USER\Downloads\매장 운영 매뉴얼_part v1_20260410.pdf", "store_ops"),
    (r"C:\Users\USER\Downloads\세일즈매뉴얼_서비스 part v1_20260410.pdf", "service"),
    (r"C:\Users\USER\Downloads\세일즈매뉴얼_제품part v1_20260410.pdf", "product"),
]

out_dir = Path(r"C:\Users\USER\Desktop\오가렌_DX\세일즈가이드\_pdf_inspect")
out_dir.mkdir(exist_ok=True)

for src, name in PDFS:
    d = fitz.open(src)
    print(f"\n===== {name} ({d.page_count}p) =====")
    pages_info = []
    for i in range(d.page_count):
        text = d.load_page(i).get_text("text")
        first_lines = [l.strip() for l in text.split("\n") if l.strip()]
        first = " | ".join(first_lines[:3])[:120]
        pages_info.append({"page": i+1, "head": first, "len": len(text)})
        if i < 40:  # 첫 40페이지는 출력
            print(f"  p{i+1:3d}: {first}")
    # full json
    (out_dir / f"{name}_pages.json").write_text(
        json.dumps(pages_info, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    # full text
    full = "\n\n=== PAGE BREAK ===\n\n".join(
        d.load_page(i).get_text("text") for i in range(d.page_count)
    )
    (out_dir / f"{name}_full.txt").write_text(full, encoding="utf-8")
    d.close()
print("\nDone.")

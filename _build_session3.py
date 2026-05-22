# -*- coding: utf-8 -*-
"""
세션 3 (쇼룸·백화점 교육자료) 자산 빌드:
  1) 각 PDF의 1페이지 → 16:9 썸네일 JPG 추출 → assets/session3/thumbs/
  2) 20MB 초과 PDF는 이미지 화질 다운스케일로 압축 → assets/download/session3/<name>_lite.pdf
     (원본은 그대로 두고 사이트는 lite를 링크. 원본이 20MB 이하면 그대로 사용)
  3) _session3_manifest.json 출력: session3.html 빌드 입력
"""
from __future__ import annotations
import fitz  # PyMuPDF
import io, json
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).parent
DL = ROOT / "assets" / "download" / "session3"
THUMBS = ROOT / "assets" / "session3" / "thumbs"
THUMBS.mkdir(parents=True, exist_ok=True)

# 목록: (slug, 카테고리, 제목, 설명, 원본파일명 또는 None)
ITEMS = [
    ("signature",     "lineup",  "시그니처 매트리스 라인업",        "단단함과 안정감의 시그니처 3종 — 라인업 교육자료",                "01_signature_lineup_251216.pdf"),
    ("hybrid",        "lineup",  "하이브리드 매트리스 라인업",       "포켓스프링 + 메모리폼 — 하이브리드 라인업 교육자료",              "02_hybrid_lineup_251216.pdf"),
    ("stable",        "lineup",  "스테이블 매트리스 라인업",         "안정적 지지력의 스테이블 라인 — 라인업 교육자료",                  "03_stable_lineup_251216.pdf"),
    ("balanced",      "lineup",  "밸런스드 매트리스 라인업",         "균형감 있는 밸런스드 라인 — 라인업 교육자료",                       "04_balanced_lineup_251216.pdf"),
    ("compete",       "compare", "슬립퍼 매트리스 3.3.3. 경쟁사 비교", "타사 라인업과의 스펙·가격·구성 비교 (2025-12-12)",                  "05_competitor_compare_251212.pdf"),
    ("nooer",         "brand",   "누어 교육자료 — 포근히, 탄탄히",   "누어 라인의 브랜드 스토리·라인업·소재 (2024-09-23)",                 "06_nooer_education_240923.pdf"),
    ("snc",           "brand",   "슬립퍼 & 누어 매트리스 비교",      "두 브랜드의 포지셔닝·소재·라인업 차이 (2024-02-20)",                "07_sleeper_nooer_compare_240220.pdf"),
    ("radon",         "safety",  "슬립퍼 매트리스 라돈 안전성",      "라돈 검사 데이터·인증 — 안전성 응대 자료 (2024-02-20)",             "08_radon_safety_240220.pdf"),
    ("classify",      "brand",   "슬립퍼 매트리스 분류표",           "전 라인업 분류·스펙 한장 정리 (2024-05-13)",                         "09_mattress_classification_240513.pdf"),
    ("cervical",      "pillow",  "경추베개 팜플렛",                  "경추베개 라인 — 매장 응대용 팜플렛 (2024-02-14)",                    "10_cervical_pillow_240214.pdf"),
    ("hardy",         "bed",     "하디 침대 교육용 자료",            "하디 프레임 — 구조·소재·구성 교육 슬라이드 (PPTX)",                  "11_hardy_bed_education.pptx"),
]

MAX_BYTES = 20 * 1024 * 1024
LITE_DPI = 110          # 다운스케일 dpi
LITE_JPEG_QUALITY = 70


def fmt_size(n: int) -> str:
    if n >= 1024 * 1024:
        return f"{n/1024/1024:.1f} MB"
    return f"{n/1024:.0f} KB"


def make_thumb(pdf_path: Path, out_path: Path) -> bool:
    try:
        doc = fitz.open(pdf_path)
        if doc.page_count == 0:
            return False
        page = doc.load_page(0)
        # 16:9 crop centered on rendered page
        zoom = 2.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        # crop to 16:9
        W, H = img.size
        target = W / H
        want = 16 / 9
        if target > want:
            new_w = int(H * want)
            x0 = (W - new_w) // 2
            img = img.crop((x0, 0, x0 + new_w, H))
        else:
            new_h = int(W / want)
            y0 = (H - new_h) // 2
            img = img.crop((0, y0, W, y0 + new_h))
        img.thumbnail((960, 540), Image.LANCZOS)
        img.save(out_path, "JPEG", quality=85, optimize=True)
        doc.close()
        return True
    except Exception as e:
        print(f"  ! thumb 실패 {pdf_path.name}: {e}")
        return False


def make_lite_pdf(src: Path, dst: Path) -> bool:
    """이미지 위주 PDF를 dpi 낮춰 재렌더링 → 사이즈 축소."""
    try:
        doc = fitz.open(src)
        out = fitz.open()
        for i in range(doc.page_count):
            page = doc.load_page(i)
            mat = fitz.Matrix(LITE_DPI / 72, LITE_DPI / 72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_bytes = io.BytesIO()
            Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB").save(
                img_bytes, "JPEG", quality=LITE_JPEG_QUALITY, optimize=True
            )
            img_bytes.seek(0)
            rect = fitz.Rect(0, 0, pix.width, pix.height)
            new_page = out.new_page(width=pix.width, height=pix.height)
            new_page.insert_image(rect, stream=img_bytes.read())
        out.save(dst, deflate=True, garbage=4)
        out.close(); doc.close()
        return True
    except Exception as e:
        print(f"  ! lite 실패 {src.name}: {e}")
        return False


manifest = []
for slug, cat, title, desc, fname in ITEMS:
    src = DL / fname
    if not src.exists():
        print(f"[SKIP] {fname} 없음"); continue

    ext = src.suffix.lower()
    size = src.stat().st_size
    print(f"[{slug}] {fname}  ({fmt_size(size)})")

    # 썸네일
    thumb_name = f"{slug}.jpg"
    thumb_path = THUMBS / thumb_name
    thumb_rel = None
    if ext == ".pdf":
        if make_thumb(src, thumb_path):
            thumb_rel = f"assets/session3/thumbs/{thumb_name}"
    else:
        thumb_rel = None  # PPTX는 썸네일 없음

    # 페이지 수
    pages = None
    if ext == ".pdf":
        try:
            d = fitz.open(src); pages = d.page_count; d.close()
        except Exception:
            pages = None

    # 다운로드 파일 결정 (PDF만 lite 생성)
    serve_name = fname
    if ext == ".pdf" and size > MAX_BYTES:
        lite_name = src.stem + "_lite.pdf"
        lite_path = DL / lite_name
        print(f"  → compress {src.name} → {lite_name}")
        if make_lite_pdf(src, lite_path):
            lite_size = lite_path.stat().st_size
            print(f"    {fmt_size(size)} → {fmt_size(lite_size)}")
            if lite_size < size:
                serve_name = lite_name
                size = lite_size
            # 원본은 git에서 제외 (수동 .gitignore로 처리)

    manifest.append({
        "slug": slug,
        "category": cat,
        "title": title,
        "desc": desc,
        "file": f"assets/download/session3/{serve_name}",
        "thumb": thumb_rel,
        "size": fmt_size(size),
        "size_bytes": size,
        "pages": pages,
        "ext": ext.lstrip("."),
    })

out_path = ROOT / "_session3_manifest.json"
out_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"\nOK manifest: {out_path}")
print(f"  items: {len(manifest)}")

# 큰 원본 PDF는 git에 안 올림 → 삭제 (lite만 남김)
print("\nDelete oversized originals (kept lite):")
for slug, cat, title, desc, fname in ITEMS:
    src = DL / fname
    if not src.exists():
        continue
    if src.suffix.lower() != ".pdf":
        continue
    lite = DL / (src.stem + "_lite.pdf")
    if lite.exists() and lite.stat().st_size < src.stat().st_size:
        sz = src.stat().st_size
        src.unlink()
        print(f"  - removed {src.name} ({sz/1024/1024:.1f} MB)")


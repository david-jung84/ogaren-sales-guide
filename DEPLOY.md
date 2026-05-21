# Deploy

## 최초 1회 — GitHub 저장소 생성 + Pages 활성화

1. GitHub 웹에서 새 저장소 생성: https://github.com/new
   - **Owner**: `david-jung84`
   - **Repository name**: `ogaren-sales-guide`
   - **Visibility**: Private (사내 자료) — 이 경우 GitHub Pages는 유료 플랜 필요. Public이면 무료.
   - README/`.gitignore` 추가 옵션은 모두 **체크 해제** (이미 로컬에 있음)

2. 로컬에서 origin 연결 + 푸시:

   ```powershell
   cd C:\Users\USER\Desktop\오가렌_DX\세일즈가이드
   git remote add origin https://github.com/david-jung84/ogaren-sales-guide.git
   git branch -M main
   git push -u origin main
   ```

3. GitHub Pages 활성화:
   - 저장소 → **Settings** → **Pages**
   - Source: **Deploy from a branch**
   - Branch: **main** / Folder: **/ (root)**
   - **Save**
   - 1~2분 후 `https://david-jung84.github.io/ogaren-sales-guide/` 접속 가능

## 갱신 워크플로우

```powershell
# 1) 컨텐츠 변경 (예: part6_revised.js 수정)
# 2) PPT 재빌드
node split_and_build.js
# 3) 사이트 재빌드 (PPT → PNG + HTML 생성)
python build_site.py
# 4) 커밋 + 푸시
git add .
git commit -m "update: <description>"
git push
```

푸시 후 1~2분 내 GitHub Pages 자동 갱신.

## 트러블슈팅

- **build_site.py가 PowerPoint COM에서 실패**: PPT 파일이 열려 있으면 안 됨. 모두 닫고 재시도.
- **Pages가 404를 반환**: Settings → Pages에서 Branch / Folder 설정 다시 확인. 푸시 후 1~2분 대기 필요.
- **PNG가 깨져 보임**: `assets/session1/` 또는 `assets/session2/`가 누락되지 않았는지 확인. `python build_site.py` 재실행.

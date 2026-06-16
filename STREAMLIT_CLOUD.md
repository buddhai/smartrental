# Streamlit Cloud 배포 가이드

## 사전 준비 (로컬에서 완료됨)

- `git init` + `main` 브랜치 초기 커밋
- `requirements.txt`, `.python-version` (3.11), `templates/*.xlsx` 포함
- 앱 진입점: **`streamlit_app.py`**

## 1단계 — GitHub에 올리기

PowerShell에서 프로젝트 폴더로 이동 후:

```powershell
cd "C:\Users\User\Documents\수익률자동화"

# gh 인식 안 될 때 (PowerShell 새로 열기 전 임시 해결)
$env:Path += ";C:\Program Files\GitHub CLI"

& "C:\Program Files\GitHub CLI\gh.exe" auth login
powershell -ExecutionPolicy Bypass -File deploy_github.ps1
```

이미 GitHub에 빈 저장소를 만들었다면:

```powershell
git remote add origin https://github.com/<사용자명>/<저장소명>.git
git push -u origin main
```

## 2단계 — Streamlit Cloud 연결

1. [https://share.streamlit.io](https://share.streamlit.io) 접속
2. **Sign in with GitHub** (위와 같은 계정)
3. **Create app** 클릭
4. 설정:
   - **Repository**: `사용자명/smartquote-rental` (본인 저장소)
   - **Branch**: `main`
   - **Main file path**: `streamlit_app.py`
5. **Deploy** 클릭

배포 URL 예: `https://smartquote-rental.streamlit.app`

## 3단계 — 배포 확인

- 앱이 열리면 상품 선택 → 시나리오 추가 → **수익률분석표 다운로드** 테스트
- 실패 시 Streamlit Cloud **Manage app → Logs** 에서 오류 확인

## 자주 나는 오류

| 증상 | 해결 |
|---|---|
| `템플릿 파일을 찾을 수 없습니다` | `templates/` xlsx가 Git에 포함됐는지 확인 |
| `ModuleNotFoundError` | `requirements.txt`에 `streamlit`, `openpyxl` 있는지 확인 |
| 엑셀 다운로드 느림 | 시나리오 수 줄이기 (Cloud 무료 CPU 제한) |

## 롤백

- Streamlit 앱만 제거: share.streamlit.io에서 앱 Delete
- 로컬 HTML 버전: `python server.py` 그대로 사용 (변경 없음)

## 코드 수정 후 재배포

```powershell
git add .
git commit -m "업데이트"
git push
```

Streamlit Cloud는 `main` 푸시 시 자동 재배포됩니다.

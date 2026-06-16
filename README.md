# 렌탈 수익률 분석 (SmartQuote)

공급사 협의·렌탈 수익률 산출 및 수익률분석표 엑셀 생성.

## Streamlit (권장 · Cloud 배포)

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

**Streamlit Cloud:** [share.streamlit.io](https://share.streamlit.io) → GitHub 연결 → Main file: `streamlit_app.py`

## 로컬 HTML + Flask (기존)

```bash
pip install flask flask-cors openpyxl
python server.py
```

브라우저: http://localhost:5000

## 템플릿 재생성

```bash
python generate_templates.py
```

`templates/` 폴더에 12~60개월 xlsx가 있어야 엑셀 다운로드가 동작합니다.

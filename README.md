# Vocab&Click Streamlit App

`02_Vocab`의 어휘 분석 스크립트를 공동 작업자가 웹 UI로 사용할 수 있게 정리한 앱입니다.

## 구성

- `streamlit_app.py`: 가이드와 어휘 추출 UI
- `03_vocab&click.py`: 기존 CLI 분석 로직
- `merged_cefr.csv`: CEFR 단어 리스트
- `Story_Confirmed.xlsx`: 기본 입력 파일
- `analysis_checkpoint.json`: 재사용 가능한 분석 결과 캐시
- `requirements.txt`: Streamlit 배포/실행 의존성

## 로컬 실행

```powershell
cd "C:\Users\IM_1500\Desktop\스토리 텍스트\02_Vocab"
.\.venv_vocab_click\Scripts\python.exe -m pip install -r requirements.txt
.\.venv_vocab_click\Scripts\streamlit.exe run streamlit_app.py
```

## Streamlit 배포 메모

GitHub에 올릴 때는 API 키를 커밋하지 말고, 배포 환경의 secrets에 `GEMINI_API_KEY`로 등록하세요.

### GitHub Desktop으로 올리기

1. GitHub Desktop을 엽니다.
2. `File` -> `Add local repository...`를 선택합니다.
3. 이 폴더를 선택합니다: `02_Vocab/vocab_click_app`
4. `Publish repository`를 누릅니다.
5. 저장소 이름은 예: `vocab-click-app`
6. 공개 범위는 콘텐츠 보호가 필요하면 `Private`, 여러 사람이 쉽게 접근해야 하면 `Public`으로 선택합니다.

### Streamlit Community Cloud

1. https://share.streamlit.io 에 접속합니다.
2. `Create app`을 누릅니다.
3. GitHub repo, branch `main`, main file path `streamlit_app.py`를 선택합니다.
4. `Advanced settings` 또는 app settings에서 secrets에 아래 값을 등록합니다.

```toml
GEMINI_API_KEY = "your-api-key"
```

커밋 제외 권장:

- `.venv*`
- `~$*.xlsx`
- `.streamlit/secrets.toml`
- `.env`
- timestamp가 붙은 임시 결과 파일

## 입력 형식

어휘 추출 입력 XLSX에는 다음 컬럼이 필요합니다.

- `ID`
- `Title`
- `Normal Ver.`
- `Easy Ver.`
- `Difficult Ver.`

`Easy Ver.`나 `Difficult Ver.`가 비어 있으면 해당 버전의 어휘가 누락되므로 앱이 실행 전에 오류로 표시합니다.

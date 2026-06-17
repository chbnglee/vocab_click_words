# Vocab&Click Streamlit App

`02_Vocab`의 어휘 분석 스크립트를 공동 작업자가 웹 UI로 사용할 수 있게 정리한 앱입니다.

## 구성

- `streamlit_app.py`: 가이드와 어휘 추출 UI
- `03_vocab&click.py`: 기존 CLI 분석 로직
- `merged_cefr.csv`: CEFR 단어 리스트
- `streamlit_app.py`: 웹앱 진입점
- `merged_cefr.csv`: CEFR 단어 리스트
- `03_vocab&click.py`: Gemini 분석 프롬프트/파서 로직
- `requirements.txt`: Streamlit 배포/실행 의존성

## 로컬 실행

```powershell
cd "C:\Users\IM_1500\Desktop\스토리 텍스트\02_Vocab"
.\.venv_vocab_click\Scripts\python.exe -m pip install -r requirements.txt
.\.venv_vocab_click\Scripts\streamlit.exe run streamlit_app.py
```

## Streamlit 배포 메모

API 키는 repo나 Streamlit secrets에 저장하지 않습니다. 각 사용자가 어휘 추출을 실행할 때 앱 화면에서 직접 입력합니다. 가이드 탭과 CEFR 검색은 API 키 없이 사용할 수 있습니다.

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
4. 별도 secrets 설정은 필요하지 않습니다.

커밋 제외 권장:

- `.venv*`
- `~$*.xlsx`
- `.streamlit/secrets.toml`
- `.env`
- `Story_Confirmed*.xlsx`
- `analysis_checkpoint*.json`
- timestamp가 붙은 임시 결과 파일

## 입력 형식

어휘 추출 입력 XLSX에는 다음 컬럼이 필요합니다.

- `ID`
- `Title`
- `Normal Ver.`
- `Easy Ver.`
- `Difficult Ver.`

`Easy Ver.`나 `Difficult Ver.`가 비어 있으면 해당 버전의 어휘가 누락되므로 앱이 실행 전에 오류로 표시합니다.

## Checkpoint JSON

`analysis_checkpoint.json`은 이전 분석 결과를 저장한 캐시입니다. 같은 ID를 다시 분석하지 않고 이어서 작업하고 싶을 때 앱에 업로드합니다. 처음 실행하거나 전부 새로 분석하려면 업로드하지 않아도 됩니다.

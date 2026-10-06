# 티스토리 자동 발행 매크로

키워드 입력 → 네이버 블로그 참고글 선택 → Gemini 원고 생성 → 참고글 유사도 검사/재작성
→ 카드 이미지 생성·업로드 → 티스토리 에디터 자동 주입 (발행 버튼은 직접 누름)

## 설치

```bash
pip install -r requirements.txt
python tistory_macro.py
```

같은 폴더에 다음 리소스 파일을 두세요 (없어도 실행은 됩니다):

| 파일 | 용도 | 없으면 |
|---|---|---|
| `template.png`, `template_2.png` | 카드 이미지 배경 | 흰 배경 |
| `NanumSquareNeo-cBd/dEb/eHv.ttf` | 카드 글꼴 | 맑은 고딕 (Windows) |
| `icon.png`, `icon.ico` | 창 아이콘 | 기본 아이콘 |

## 처음 실행

1. 앱이 뜨면 **설정 창**이 자동으로 열립니다. API 키 4개를 넣고 저장하세요.
   (`config.json`에 저장되며 git에는 올라가지 않습니다)
2. 티스토리 계정은 설정 창에서 한 줄에 `이름 URL` 형식으로 추가/삭제합니다.

## 사용법

1. 계정을 고르고 키워드를 한 줄에 하나씩 입력 → **실행 시작** (`Ctrl+Enter`)
2. 키워드마다 참고할 글을 체크 → **다음** (`Enter`)
   - `열기`로 원문 미리보기, 아래 칸에 URL 직접 추가 가능
   - **참고글 직접 선택**을 끄면 상위 N개를 자동으로 사용
3. 원고를 만드는 동안 크롬이 먼저 열립니다. **그 사이에 로그인**하고 **로그인 완료**를 누르세요.
4. 글이 하나씩 에디터에 들어갑니다. 브라우저에서 발행 → **발행 완료 → 다음 글 주입**
5. 언제든 **중지** 가능. 모든 원고는 `output/날짜_계정/`에 HTML로 저장됩니다 (**📁 결과 폴더**).

## 원고 차별화 (참고글 유사도 검사)

- 원고를 만든 뒤, 참고글과 **8글자 이상 연속으로 겹치는 표현의 비율**을 계산합니다.
- 기준(`max_similarity`, 기본 12%)을 넘거나 참고글을 거의 그대로 옮긴 문장이 있으면,
  그 문장들을 짚어서 Gemini에 **재작성**을 요청합니다 (`rewrite_retries`, 기본 2회).
- 가장 덜 겹치는 버전을 사용하고, 그래도 기준을 넘으면 로그에 경고를 띄웁니다.
- 기준을 낮출수록 원고를 더 많이 바꿉니다 (설정 창에서 변경).

## 주요 설정 (`config.json`)

| 키 | 기본값 | 설명 |
|---|---|---|
| `gemini_models` | `["gemini-2.5-flash", "gemini-2.5-flash-lite"]` | 앞에서부터 시도 |
| `chrome_version_main` | `null` | 크롬 버전 오류가 날 때만 숫자로 지정 |
| `max_similarity` | `0.12` | 참고글 겹침 허용치 |
| `rewrite_retries` | `2` | 재작성 최대 횟수 |
| `min_body_chars` | `2500` | 목표 분량 |
| `search_count` | `5` | 네이버 검색 결과 개수 |
| `wait_timeout_min` | `30` | 로그인/발행 대기 시간(분), 0이면 무제한 |

API 키는 환경변수 `NAVER_CLIENT_ID`, `NAVER_CLIENT_SECRET`, `GEMINI_API_KEY`, `IMGBB_API_KEY`로도 넣을 수 있습니다.

## exe 빌드 (선택)

```bash
pyinstaller --onefile --noconsole --icon icon.ico ^
  --add-data "template.png;." --add-data "template_2.png;." ^
  --add-data "NanumSquareNeo-cBd.ttf;." --add-data "NanumSquareNeo-dEb.ttf;." ^
  --add-data "NanumSquareNeo-eHv.ttf;." --add-data "icon.png;." tistory_macro.py
```

`config.json`, `output/`, `logs/`는 exe와 같은 폴더에 생깁니다.

# 티스토리 자동 발행 매크로

키워드 입력 → 네이버 블로그 참고글 선택 → AI 원고 생성 (Claude Code 구독 또는 Gemini) → 참고글 유사도 검사/재작성
→ 카드 이미지 생성·업로드 → 티스토리 에디터 자동 주입 (발행 버튼은 직접 누름)

## 설치

```bash
pip install -r requirements.txt
python tistory_macro.py
```

같은 폴더에 다음 리소스 파일을 두세요 (없어도 실행은 됩니다):

| 파일 | 용도 | 없으면 |
|---|---|---|
| `template.png`, `template_2.png` | 카드 이미지 배경 (포함됨) | 흰 배경 |
| `NanumSquareNeo-cBd/dEb/eHv.ttf` | 카드 글꼴 | 맑은 고딕 (Windows) |
| `icon.png`, `icon.ico` | 창 아이콘 | 기본 아이콘 |

## 원고 작성 AI: Claude Code (기본, API 비용 없음)

PC에 설치된 Claude Code를 `claude -p`로 불러서 원고를 씁니다. **Claude 구독 사용량으로 처리**되므로
API 키나 별도 결제가 필요 없습니다. 하루 몇 개 정도 쓰는 용도에 맞습니다.

1. Claude Code 설치 (Windows PowerShell): `irm https://claude.ai/install.ps1 | iex`
2. 터미널에서 `claude` 를 한 번 실행해 **Claude 계정(구독)으로 로그인**
3. 앱의 **⚙ 설정 → Claude 연결 테스트**로 확인

- 글 하나에 1~3분 걸립니다 (재작성까지 하면 더 걸림).
- 구독 사용량 한도에 걸리면 로그에 안내가 나옵니다. **Gemini 키를 넣어두면 그때 자동으로 Gemini로 대신 씁니다.**
- 환경변수에 `ANTHROPIC_API_KEY`가 있으면 API로 과금될 수 있어서, 매크로가 실행할 때 자동으로 빼고 실행합니다.
- 모델은 설정에서 `sonnet` / `opus` 지정 가능 (비우면 Claude Code 기본값). opus는 사용량을 더 많이 씁니다.

## 처음 실행

1. 앱이 뜨면 **설정 창**이 자동으로 열립니다. 네이버 키 2개와 imgbb 키를 넣고 저장하세요.
   (Gemini 키는 Gemini를 쓰거나 백업용으로만 필요. 모두 `config.json`에 저장되며 git에는 올라가지 않습니다)
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
  그 문장들을 짚어서 AI에 **재작성**을 요청합니다 (`rewrite_retries`, 기본 2회).
- 가장 덜 겹치는 버전을 사용하고, 그래도 기준을 넘으면 로그에 경고를 띄웁니다.
- 기준을 낮출수록 원고를 더 많이 바꿉니다 (설정 창에서 변경).

## 주요 설정 (`config.json`)

| 키 | 기본값 | 설명 |
|---|---|---|
| `ai_provider` | `"claude_cli"` | `"claude_cli"` 또는 `"gemini"` |
| `claude_model` | `""` | 비우면 기본, `sonnet` / `opus` |
| `fallback_to_gemini` | `true` | Claude 실패 시 Gemini 키가 있으면 Gemini로 재시도 |
| `gemini_models` | `["gemini-2.5-flash", "gemini-2.5-flash-lite"]` | 앞에서부터 시도 |
| `chrome_version_main` | `null` | 크롬 버전 오류가 날 때만 숫자로 지정 |
| `max_similarity` | `0.12` | 참고글 겹침 허용치 |
| `rewrite_retries` | `2` | 재작성 최대 횟수 |
| `min_body_chars` | `2500` | 목표 분량 |
| `search_count` | `5` | 네이버 검색 결과 개수 |
| `wait_timeout_min` | `30` | 로그인/발행 대기 시간(분), 0이면 무제한 |

API 키는 환경변수 `NAVER_CLIENT_ID`, `NAVER_CLIENT_SECRET`, `GEMINI_API_KEY`, `IMGBB_API_KEY`로도 넣을 수 있습니다.

## exe 빌드

윈도우에서 **`build.bat` 더블클릭** → `dist\TistoryMacro.exe` 생성.
템플릿 2개, 폰트 3개, 아이콘이 exe 안에 같이 묶입니다.

`config.json`, `output/`, `logs/`는 exe와 같은 폴더에 생깁니다.

## 이미지 저장 위치

- **블로그에 들어가는 이미지:** imgbb에 업로드된 주소(`https://i.ibb.co/...`)로 본문에 삽입 (만료 설정 없음)
- **내 PC 사본:** `output/날짜_계정/images/글번호_이미지번호.jpg` (원고 HTML과 같은 폴더, 앱의 📁 결과 폴더 버튼)

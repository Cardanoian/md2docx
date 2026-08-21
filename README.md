# md2docx

마크다운(`.md`) 파일을 **Word(`.docx`)** 와 **한글(`.hwpx`)** 로 변환하는 파이썬 스크립트입니다.

Word 변환은 [**pandoc**](https://pandoc.org/)을 호출하고, 한글 변환은 순수 파이썬(`md2hwpx`)으로
HWPX를 만듭니다. 마크다운 형식(헤더·표·리스트·인용·각주 등)과 로컬 이미지 임베드를 지원하며,
Word 쪽은 코드 구문 강조와 목차 자동 생성까지 포함합니다. 한글 파일을 **생성하는 데**
한컴오피스는 필요 없습니다.

> 설치부터 OS별 사용법, 트러블슈팅까지 전 과정을 단계별로 안내하는 **상세 매뉴얼**은
> [manual.md](manual.md)를 참고하세요. 이 문서는 핵심만 요약한 빠른 안내서입니다.

---

## 주요 기능

- **마크다운 형식 전체 보존** — 헤더, 볼드/이탤릭, 중첩 리스트, 표, 인용, 각주, 링크,
  수평선, 체크박스 등 ([GitHub 스타일 확장 마크다운](https://github.github.com/gfm/) 기반)
- **로컬 이미지 자동 임베드** — 마크다운 파일 위치를 기준으로 상대경로를 자동 해석하고,
  이미지를 결과 파일 내부에 포함해 단일 파일로 완결
- **코드 구문 강조** — Word 변환 시 pandoc 내장 하이라이터로 8가지 테마 지원
- **목차(TOC) 자동 생성** — Word 변환의 `--toc` 옵션으로 3단계까지 자동 생성
- **세부 스타일 제어** — 참조 문서(`.docx` / `.hwpx`)로 폰트·헤딩·코드블록 등 커스터마이징
- **한글(HWPX) 변환** — `python md2hwpx.py 문서.md` 로 한컴 설치 없이 `.hwpx` 생성
- **간편한 설치** — `pip install -r requirements.txt` 한 번으로 pandoc·HWPX 변환기까지 함께 설치
  (Linux / macOS / Windows 동일 절차)

---

## 요구사항

- **Python 3.10 이상** (스크립트는 3.10+ 문법을 사용합니다)
- **pandoc 바이너리** — Word 변환용. `requirements.txt`의 `pypandoc-binary`로 함께 설치됩니다.
  시스템에 pandoc을 직접 설치(`brew`/`apt`/`winget`)한 경우 `md2docx.py`가 그쪽을 **우선** 사용합니다.
- **md2hwpx** — 한글 변환용. 같은 `requirements.txt`로 함께 설치됩니다.

---

## 빠른 시작

### Ubuntu / macOS

```bash
cd ~/md2docx
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip && pip install -r requirements.txt
python md2docx.py docs/report.md --toc
python md2hwpx.py docs/report.md
```

### Windows (PowerShell)

```powershell
cd $HOME\md2docx
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip ; pip install -r requirements.txt
python md2docx.py docs\report.md --toc
python md2hwpx.py docs\report.md
```

> 가상환경(venv) 생성·활성화 과정에서 막히는 경우(예: Ubuntu `python3-venv` 누락,
> Windows PowerShell 실행 정책 차단)는 [manual.md](manual.md)의 OS별 설치 가이드와
> 트러블슈팅을 참고하세요.

---

## 사용법

가상환경이 활성화된 상태(프롬프트에 `(.venv)` 표시)에서 실행합니다.

```bash
# Word: report.md → report.docx (같은 폴더에 생성)
python md2docx.py docs/report.md

# 한글: report.md → report.hwpx
python md2hwpx.py docs/report.md

# 출력 경로를 직접 지정
python md2docx.py docs/report.md -o out/report.docx
python md2hwpx.py docs/report.md -o out/report.hwpx

# Word만: 목차(TOC) 자동 생성 + 코드 하이라이트 테마 변경
python md2docx.py docs/report.md --toc --highlight-style breezedark
```

변환이 끝나면 `변환 완료: report.md → .../report.docx`(또는 `.hwpx`) 메시지가 출력됩니다.

---

## 옵션

공통 옵션은 두 스크립트 모두에 있습니다. `--toc` 와 `--highlight-style` 은 Word 전용입니다.

| 옵션                    | 설명                                                       |
| ----------------------- | ---------------------------------------------------------- |
| `input`                 | (필수) 입력 마크다운 파일 경로                             |
| `-o`, `--output`        | 출력 경로 (생략 시 입력과 같은 이름의 `.docx` / `.hwpx`)   |
| `--toc`                 | (docx) 문서 앞에 목차 자동 생성 (3단계까지)                |
| `--highlight-style`     | (docx) 코드 하이라이트 테마 선택 (기본: `tango`)           |
| `--reference-doc`       | 세부 스타일 제어용 참조 문서 적용                          |
| `--resource-path`       | 이미지 등 리소스를 추가로 탐색할 폴더 (여러 번 지정 가능)  |
| `--make-reference DEST` | 참조 문서 템플릿을 DEST에 생성하고 종료                    |
| `-h`, `--help`          | 도움말 출력                                                |

**하이라이트 테마 목록** (docx):
`pygments`, `tango`(기본값), `espresso`, `zenburn`, `kate`, `monochrome`(흑백),
`breezedark`(다크), `haddock`

---

## 세부 스타일 제어 (참조 문서)

폰트, 헤딩 색상, 코드블록 배경/글꼴 같은 세부 스타일까지 통제하려면 "참조 문서
(reference document)" 방식을 사용합니다.

```bash
# Word 템플릿
python md2docx.py --make-reference template.docx
python md2docx.py docs/report.md --reference-doc template.docx

# 한글 템플릿
python md2hwpx.py --make-reference template.hwpx
python md2hwpx.py docs/report.md --reference-doc template.hwpx
```

`template.docx`는 워드에서 `Source Code` / `Heading 1~6` / `Normal` 스타일을 편집하면 됩니다.
`template.hwpx`는 한글에서 제목·본문·표 스타일을 편집한 뒤 재사용하면 됩니다.

---

## 폴더 구조

```
md2docx/
├── md2docx.py          # Word(.docx) 변환
├── md2hwpx.py          # 한글(.hwpx) 변환
├── requirements.txt    # 의존성 목록 (pypandoc-binary, md2hwpx)
├── manual.md           # 상세 사용 매뉴얼
├── README.md           # 이 문서
├── .venv/              # 가상환경 (직접 생성)
└── docs/               # 변환할 마크다운과 이미지를 모아두는 작업 폴더(예시)
    ├── report.md
    └── images/
        └── sample.png
```

> **이미지 경로 규칙**: 마크다운 안의 `![](images/sample.png)` 같은 상대경로는
> **마크다운 파일이 있는 위치**를 기준으로 자동 해석됩니다. 즉 `report.md`와 `images/`
> 폴더가 같은 디렉터리에 있으면, 어느 위치에서 스크립트를 실행하든 이미지가 정상적으로
> 들어갑니다. 이미지가 다른 폴더에 있다면 `--resource-path` 로 탐색 경로를 추가하세요.

---

## 트러블슈팅

자주 발생하는 문제는 다음과 같습니다. 자세한 해결법은 [manual.md](manual.md)의 트러블슈팅
섹션을 참고하세요.

| 증상                                    | 빠른 해결                                                                          |
| --------------------------------------- | ---------------------------------------------------------------------------------- |
| `pandoc을 찾을 수 없습니다`             | 가상환경 활성화 여부 확인 후 `pip install -r requirements.txt` 재실행              |
| `md2hwpx 패키지를 찾을 수 없습니다`     | 가상환경 활성화 후 `pip install -r requirements.txt` 재실행                        |
| 이미지가 결과에 안 들어감               | 이미지 경로가 마크다운 파일 기준 상대경로인지 확인, 필요 시 `--resource-path` 추가 |
| Ubuntu에서 `python3 -m venv` 실패       | `sudo apt install -y python3-venv` 후 재시도                                       |
| Windows PowerShell에서 활성화 차단      | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` 실행        |
| `python` 명령이 인식되지 않음 (Windows) | `python` 대신 `py` 사용 (venv 활성화 후에는 `python` 동작)                         |
| 한글이 깨지거나 폰트가 어색함           | 참조 문서에서 `Normal`/`Source Code` 스타일을 한글 글꼴로 지정                     |

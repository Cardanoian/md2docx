# md2docx 사용 매뉴얼

마크다운(`.md`) 파일을 Word 문서(`.docx`)로 변환하는 파이썬 스크립트입니다.
마크다운 형식 전체(헤더·표·리스트·인용·각주 등), 로컬 이미지 임베드,
코드 구문 강조를 모두 지원합니다.

이 매뉴얼은 **파이썬이 이미 설치돼 있다는 가정** 하에, Ubuntu 24.04 LTS / macOS /
Windows 각각에서 가상환경(venv) 생성부터 실제 사용까지 모든 과정을 안내합니다.

---

## 0. 동작 방식 한눈에 보기

이 스크립트는 내부적으로 **pandoc**이라는 문서 변환 엔진을 호출합니다.
따라서 준비물은 두 가지뿐입니다.

1. `md2docx.py` — 변환 스크립트 (파이썬 표준 라이브러리만 사용)
2. **pandoc 바이너리** — `requirements.txt`(`pypandoc-binary`)로 함께 설치

`pip install -r requirements.txt` 한 번이면 pandoc까지 같이 깔리므로,
세 OS의 설치 절차가 거의 동일합니다. (시스템에 pandoc을 직접 설치한 경우,
스크립트는 그쪽을 우선 사용합니다.)

### 권장 폴더 구조

```
md2docx/
├── md2docx.py          # 변환 스크립트
├── requirements.txt    # 의존성 목록
├── .venv/              # 가상환경 (아래에서 생성)
└── docs/               # 변환할 마크다운과 이미지를 모아두는 작업 폴더(예시)
    ├── report.md
    └── images/
        └── sample.png
```

> **이미지 경로 규칙**: 마크다운 안의 `![](images/sample.png)` 같은 상대경로는
> **마크다운 파일이 있는 위치**를 기준으로 자동 해석됩니다. 즉 `report.md`와
> `images/` 폴더가 같은 디렉터리에 있으면, 어느 위치에서 스크립트를 실행하든
> 이미지가 정상적으로 들어갑니다.

---

## 1. 사전 확인 — 파이썬 버전

터미널(Windows는 PowerShell)에서 파이썬이 3.10 이상인지 확인합니다.

```bash
# Linux / macOS
python3 --version

# Windows
py --version
```

`Python 3.10.x` 이상이면 됩니다. (스크립트는 3.10+ 문법을 사용합니다.)

---

## 2. 운영체제별 설치

세 OS 모두 흐름은 동일합니다:
**① 작업 폴더로 이동 → ② venv 생성 → ③ venv 활성화 → ④ 의존성 설치 → ⑤ 확인**

스크립트(`md2docx.py`)와 `requirements.txt`를 먼저 같은 폴더에 넣어두세요.

### 2-A. Ubuntu 24.04 LTS

Ubuntu는 venv 생성을 위해 `python3-venv` 패키지가 필요할 수 있습니다.
(없으면 venv 생성 시 오류가 납니다.)

```bash
# (필요 시) venv 모듈 설치
sudo apt update
sudo apt install -y python3-venv

# ① 작업 폴더로 이동
cd ~/md2docx

# ② 가상환경 생성
python3 -m venv .venv

# ③ 활성화 (프롬프트 앞에 (.venv) 표시됨)
source .venv/bin/activate

# ④ 의존성 설치 (pandoc 포함)
pip install --upgrade pip
pip install -r requirements.txt

# ⑤ 설치 확인
python -c "import pypandoc; print('pandoc', pypandoc.get_pandoc_version())"
```

> **참고(PEP 668)**: Ubuntu 24.04는 시스템 파이썬에 직접 `pip install`을 하면
> `externally-managed-environment` 오류가 납니다. 위처럼 venv를 쓰면 이 문제가
> 발생하지 않으므로, 반드시 가상환경 안에서 설치하세요.

### 2-B. macOS

(Homebrew 또는 python.org 설치본 모두 동일합니다.)

```bash
# ① 작업 폴더로 이동
cd ~/md2docx

# ② 가상환경 생성
python3 -m venv .venv

# ③ 활성화
source .venv/bin/activate

# ④ 의존성 설치 (pandoc 포함)
pip install --upgrade pip
pip install -r requirements.txt

# ⑤ 설치 확인
python -c "import pypandoc; print('pandoc', pypandoc.get_pandoc_version())"
```

> macOS 기본 셸은 zsh이며 활성화 명령은 bash와 동일합니다.
> `python3`가 인식되지 않으면 [python.org](https://www.python.org/downloads/macos/)에서
> 설치하거나 `brew install python`을 사용하세요.

### 2-C. Windows

Windows는 활성화 명령이 셸(PowerShell vs 명령 프롬프트)에 따라 다릅니다.

#### PowerShell (권장)

```powershell
# ① 작업 폴더로 이동
cd $HOME\md2docx

# ② 가상환경 생성
py -m venv .venv

# ③ 활성화
.\.venv\Scripts\Activate.ps1

# ④ 의존성 설치 (pandoc 포함)
python -m pip install --upgrade pip
pip install -r requirements.txt

# ⑤ 설치 확인
python -c "import pypandoc; print('pandoc', pypandoc.get_pandoc_version())"
```

> **활성화 시 보안 오류가 나는 경우**
> `이 시스템에서 스크립트를 실행할 수 없으므로...` 메시지가 뜨면,
> 현재 사용자에 한해 스크립트 실행을 허용한 뒤 다시 활성화하세요.
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
> ```

#### 명령 프롬프트(cmd)를 쓰는 경우

```bat
cd %USERPROFILE%\md2docx
py -m venv .venv
.\.venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
python -c "import pypandoc; print('pandoc', pypandoc.get_pandoc_version())"
```

---

## 3. (선택) pandoc을 시스템에 직접 설치

`requirements.txt`만으로 충분하지만, 시스템 전역에서 pandoc을 쓰고 싶다면
아래처럼 설치할 수 있습니다. 설치돼 있으면 스크립트가 이쪽을 **우선** 사용합니다.

| OS      | 명령                      |
| ------- | ------------------------- |
| Ubuntu  | `sudo apt install pandoc` |
| macOS   | `brew install pandoc`     |
| Windows | `winget install pandoc`   |

설치 확인: `pandoc --version`

---

## 4. 기본 사용법

가상환경이 활성화된 상태(프롬프트에 `(.venv)` 표시)에서 실행합니다.

```bash
# 가장 단순한 변환: report.md → report.docx (같은 폴더에 생성)
python md2docx.py docs/report.md

# 출력 경로를 직접 지정
python md2docx.py docs/report.md -o out/report.docx

# 목차(TOC) 자동 생성 + 코드 하이라이트 테마 변경
python md2docx.py docs/report.md --toc --highlight-style breezedark
```

변환이 끝나면 `변환 완료: report.md → .../report.docx` 메시지가 출력됩니다.

---

## 5. 옵션 전체

| 옵션                    | 설명                                                     |
| ----------------------- | -------------------------------------------------------- |
| `input`                 | (필수) 입력 마크다운 파일 경로                           |
| `-o`, `--output`        | 출력 docx 경로 (생략 시 입력과 같은 이름의 `.docx`)      |
| `--toc`                 | 문서 앞에 목차 자동 생성 (3단계까지)                     |
| `--highlight-style`     | 코드 하이라이트 테마 선택                                |
| `--reference-doc`       | 세부 스타일 제어용 참조 문서 적용                        |
| `--resource-path`       | 이미지 등 리소스를 추가로 탐색할 폴더(여러 번 지정 가능) |
| `--make-reference DEST` | 참조 문서 템플릿을 DEST에 생성하고 종료                  |
| `-h`, `--help`          | 도움말 출력                                              |

### 하이라이트 테마 목록

`pygments`(기본 밝은 톤), `tango`(기본값), `espresso`, `zenburn`,
`kate`, `monochrome`(흑백), `breezedark`(다크), `haddock`

---

## 6. 코드/헤딩 스타일을 더 세밀하게 제어하기 (참조 문서)

폰트, 헤딩 색상, **코드블록 배경/글꼴** 같은 세부 스타일까지 통제하려면
"참조 문서(reference document)" 방식을 사용합니다. 3단계입니다.

```bash
# 1) 템플릿 추출
python md2docx.py --make-reference template.docx

# 2) template.docx를 워드/한컴오피스에서 열어 스타일 편집 후 저장
#    - 'Source Code' 스타일  → 코드블록 글꼴·배경
#    - 'Heading 1' ~ 'Heading 6' → 제목 색상·크기
#    - 'Normal' → 본문 기본 글꼴

# 3) 편집한 템플릿을 적용해 변환
python md2docx.py docs/report.md --reference-doc template.docx
```

한 번 만든 `template.docx`는 계속 재사용할 수 있어, 사내 표준 양식이나
보고서 서식을 일관되게 유지하는 데 유용합니다.

---

## 7. 작업 종료 / 재개

```bash
# 가상환경 빠져나오기 (모든 OS 공통)
deactivate
```

다음에 다시 작업할 때는 폴더로 이동해 활성화 명령만 실행하면 됩니다.
(venv 생성과 설치는 처음 한 번이면 충분합니다.)

| OS / 셸            | 재활성화 명령                  |
| ------------------ | ------------------------------ |
| Ubuntu / macOS     | `source .venv/bin/activate`    |
| Windows PowerShell | `.\.venv\Scripts\Activate.ps1` |
| Windows cmd        | `.\.venv\Scripts\activate.bat` |

---

## 8. 트러블슈팅

**`pandoc을 찾을 수 없습니다` 오류**
가상환경이 활성화돼 있는지(프롬프트에 `(.venv)`), 그리고
`pip install -r requirements.txt`가 성공했는지 확인하세요.

**이미지가 docx에 안 들어감**
마크다운의 이미지 경로가 마크다운 파일 기준 상대경로인지 확인하세요.
이미지가 다른 폴더에 있다면 `--resource-path /이미지/폴더/경로`로 추가 지정합니다.

**Ubuntu에서 `python3 -m venv` 실패**
`sudo apt install -y python3-venv` 후 다시 시도하세요.

**Windows PowerShell에서 활성화 차단**
`Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` 실행 후
다시 활성화하세요.

**`python` 명령이 인식되지 않음 (Windows)**
`python` 대신 `py`를 사용하세요. (단, venv 활성화 이후에는 `python`이 동작합니다.)

**한글이 깨지거나 폰트가 어색함**
참조 문서(6번)에서 `Normal`/`Source Code` 스타일의 글꼴을
Pretendard·맑은 고딕 등 설치된 한글 글꼴로 지정하면 해결됩니다.

---

## 9. 빠른 시작 요약 (복붙용)

```bash
# --- Ubuntu / macOS ---
cd ~/md2docx
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip && pip install -r requirements.txt
python md2docx.py docs/report.md --toc
```

```powershell
# --- Windows (PowerShell) ---
cd $HOME\md2docx
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip ; pip install -r requirements.txt
python md2docx.py docs\report.md --toc
```
# Rails 풀스택 교본 — 참조 문서(template.docx) 스타일 정의서

루비 온 레일즈 풀스택 개발 **교본**용 Word 템플릿(`template.docx`)에 적용할 스타일을
표로 정리한 문서입니다. [manual.md](manual.md) 6번 절의 "참조 문서" 워크플로에서,
`template.docx`를 Word/한컴오피스로 열어 **각 스타일을 아래 표대로 편집**할 때
그대로 보고 입력하는 체크리스트 역할을 합니다.

**디자인 방향**

- **강조 색**: Rails 레드(`#C5170A`) 계열 — 헤딩과 포인트에 사용
- **글꼴**: 본문 **Pretendard**, 코드 **D2Coding**
- **코드블록**: 연한 회색 GitHub 스타일 배경(`#F6F8FA`) + 얇은 테두리
- **표**: 모든 셀 0.5pt 옅은 회색(`#D0D7DE`) 격자 + 헤더 하단 1pt 레드선으로 헤더/본문 구분

---

## 1. 색상 팔레트

| 이름           | HEX      | 용도                                       |
| -------------- | -------- | ------------------------------------------ |
| Rails Red      | `C5170A` | 문서 제목, 제목 1, 목차 제목, 인용 좌측 바 |
| Deep Red       | `A21309` | 제목 2                                     |
| Inline Red     | `B0120A` | 인라인 코드 글자                           |
| Slate          | `374151` | 제목 3·4, 인용 본문                        |
| Gray           | `6B7280` | 제목 6, 부제, 캡션                         |
| Body Ink       | `1A1A1A` | 본문 글자                                  |
| Code Ink       | `24292E` | 코드블록 글자                              |
| Code BG        | `F6F8FA` | 코드블록 배경                              |
| Code Border    | `D0D7DE` | 코드블록 테두리                            |
| Table Grid     | `D0D7DE` | 표 셀 격자선 (코드 테두리와 같은 톤)       |
| Table Header BG| `F7DDDA` | 표 헤더행 배경 (레드 옅은 틴트)            |
| Inline Code BG | `F4F4F4` | 인라인 코드 음영                           |
| Quote Tint     | `FCF3F2` | 인용/팁 박스 배경                          |
| Link Blue      | `1D4ED8` | 하이퍼링크                                 |

---

## 2. 글꼴

| 용도      | 글꼴                  | 비고                                   |
| --------- | --------------------- | -------------------------------------- |
| 본문·헤딩 | Pretendard            | 현대적 시스템 UI 톤·장문 가독성 우수   |
| 코드      | D2Coding              | 한글 폭이 고정된 코딩용 글꼴           |

> 두 글꼴 모두 무료입니다. Pretendard는 [github.com/orioncactus/pretendard](https://github.com/orioncactus/pretendard)에서
> 받을 수 있습니다. 독자/공유 대상 PC에 없으면 Word가 비슷한 글꼴로 대체해 보여주므로
> 문서 자체는 깨지지 않습니다. 배포가 잦다면 나중에 글꼴 임베드를 고려할 수 있습니다.

---

## 3. 페이지

| 항목 | 값                      |
| ---- | ----------------------- |
| 용지 | A4 (210 × 297 mm)       |
| 여백 | 상하 25 mm · 좌우 25 mm |

---

## 4. 스타일별 설정 (Word에서 이 값대로 편집)

`마크다운 요소`가 변환될 때 어떤 Word 스타일을 거치는지와, 그 스타일에 넣을 값입니다.

| 마크다운 요소         | Word 스타일 이름            | 글꼴       | 크기   | 색(HEX)  | 기타                                               |
| --------------------- | --------------------------- | ---------- | ------ | -------- | -------------------------------------------------- |
| 문서 제목             | Title                       | Pretendard | 28pt   | `C5170A` | 굵게                                               |
| 부제                  | Subtitle                    | Pretendard | 14pt   | `6B7280` |                                                    |
| `#` 제목 1            | Heading 1 (제목 1)          | Pretendard | 20pt   | `C5170A` | 굵게, 아래 0.5pt 레드 보더                         |
| `##` 제목 2           | Heading 2 (제목 2)          | Pretendard | 16pt   | `A21309` | 굵게                                               |
| `###` 제목 3          | Heading 3 (제목 3)          | Pretendard | 14pt   | `374151` | 굵게                                               |
| `####` 제목 4         | Heading 4                   | Pretendard | 12pt   | `374151` | 굵게                                               |
| `#####` 제목 5        | Heading 5                   | Pretendard | 11pt   | `4B5563` | 굵게                                               |
| `######` 제목 6       | Heading 6                   | Pretendard | 10.5pt | `6B7280` | 굵게·이탤릭                                        |
| 본문 문단             | Normal / Body Text (표준)   | Pretendard | 10.5pt | `1A1A1A` | 줄간격 1.4                                         |
| 인라인 코드 `` `x` `` | Verbatim Char               | D2Coding   | 10pt   | `B0120A` | 음영 `F4F4F4`                                      |
| 코드블록 ` ``` `      | **Source Code (신규 추가)** | D2Coding   | 9.5pt  | `24292E` | 배경 `F6F8FA`, 사방 테두리 `D0D7DE`, 좌우 들여쓰기 |
| 인용/팁 `>`           | Block Text                  | Pretendard | 10.5pt | `374151` | 좌측 레드(`C5170A`) 바, 배경 `FCF3F2`              |
| 링크                  | Hyperlink                   | (상속)     | —      | `1D4ED8` | 밑줄                                               |
| 목차 제목             | TOC Heading                 | Pretendard | 18pt   | `C5170A` | 굵게                                               |
| 그림/표 캡션          | Caption / Image Caption     | Pretendard | 9pt    | `6B7280` | 이탤릭·가운데                                      |
| 표 `\|…\|`            | Table                       | Pretendard | 10.5pt | `1A1A1A` | 셀 0.5pt 회색(`D0D7DE`) 격자, 헤더 배경 `F7DDDA`·글자 `C5170A`·하단 1pt 레드선 |

> **Source Code** 스타일은 pandoc 기본 템플릿에 들어 있지 않습니다.
> Word에서 `template.docx`를 열고 **[스타일] 패널 → 새 스타일** 을 눌러
> 이름 `Source Code` (단락 스타일)로 새로 추가한 뒤 위 값을 적용하세요.
> 이 스타일이 없으면 코드블록은 밋밋한 기본 모양으로 나옵니다.

---

## 5. Word에서 적용하는 순서

[manual.md](manual.md) 6번 절과 동일한 3단계입니다.

```bash
# 1) 템플릿 추출
python md2docx.py --make-reference template.docx

# 2) template.docx를 Word/한컴오피스에서 열어
#    위 4번 표대로 각 스타일을 편집 (Source Code는 새로 추가) 후 저장

# 3) 편집한 템플릿을 적용해 변환
python md2docx.py docs/report.md --reference-doc template.docx --highlight-style kate
```

> 코드블록 배경이 밝은 회색이므로, 변환 시 **밝은 하이라이트 테마**(`kate`·`tango`·`pygments`)와
> 함께 쓰는 것을 권장합니다. (`breezedark` 같은 다크 테마는 밝은 배경과 어울리지 않습니다.)

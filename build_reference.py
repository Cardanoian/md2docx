#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_reference.py — Rails 풀스택 교재용 pandoc 참조 문서(template.docx) 생성

pandoc 기본 reference.docx의 스타일(styles.xml)·테마(theme1.xml)·페이지(document.xml)를
편집해 "스타일 정의서"의 디자인 토큰을 그대로 입힌다.

  추출:  pandoc --print-default-data-file reference.docx > base.docx
  실행:  python3 build_reference.py
  사용:  pandoc 본문.md --reference-doc=template.docx -o 결과.docx
"""
import zipfile
from lxml import etree

BASE = "base.docx"          # pandoc --print-default-data-file reference.docx 로 추출한 원본
OUT = "template.docx"       # 결과 참조 문서

# ── OOXML 네임스페이스 ───────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
def wq(tag): return f"{{{W}}}{tag}"
def aq(tag): return f"{{{A}}}{tag}"

# ── 디자인 토큰 (스타일 정의서 기준) ─────────────────────────────
# 폰트 — 정의서: 본문/헤딩 'Pretendard', 코드 'D2Coding'
BODY_FONT = "Pretendard"      # 본문·제목 한글 폰트 (현대적 시스템 UI 톤, 장문 가독성)
CODE_FONT = "D2Coding"        # 코드 폰트 (한글 주석까지 고정폭 정렬)

# 색상 팔레트 (RRGGBB, '#' 없이) — 정의서 1절 그대로
RAILS_RED      = "C5170A"     # 문서 제목·제목 1·목차 제목·인용 좌측 바
DEEP_RED       = "A21309"     # 제목 2
INLINE_RED     = "B0120A"     # 인라인 코드 글자
SLATE          = "374151"     # 제목 3·4, 인용 본문
H5_GRAY        = "4B5563"     # 제목 5
GRAY           = "6B7280"     # 제목 6·부제·캡션
BODY_INK       = "1A1A1A"     # 본문 글자
CODE_INK       = "24292E"     # 코드블록 글자
CODE_BG        = "F6F8FA"     # 코드블록 배경 (GitHub 라이트 톤)
CODE_BORDER    = "D0D7DE"     # 코드블록 테두리
TABLE_GRID     = "D0D7DE"     # 표 셀 격자선 (코드 테두리와 같은 톤)
INLINE_CODE_BG = "F4F4F4"     # 인라인 코드 음영
QUOTE_TINT     = "FCF3F2"     # 인용/팁 박스 배경
LINK_BLUE      = "1D4ED8"     # 하이퍼링크
# 표 헤더 배경 — 정의서에 미정의. RAILS_RED 계열의 옅은 틴트로 파생(원하면 조정).
TH_BG          = "F7DDDA"

# 크기 (half-point: 21 = 10.5pt) — 정의서 4절 그대로
SZ_BODY     = "21"   # 10.5pt
SZ_TITLE    = "56"   # 28pt
SZ_SUBTITLE = "28"   # 14pt
SZ_H1       = "40"   # 20pt
SZ_H2       = "32"   # 16pt
SZ_H3       = "28"   # 14pt
SZ_H4       = "24"   # 12pt
SZ_H5       = "22"   # 11pt
SZ_H6       = "21"   # 10.5pt
SZ_CODE     = "19"   # 9.5pt
SZ_INLINE   = "20"   # 10pt
SZ_TOC      = "36"   # 18pt
SZ_CAPTION  = "18"   # 9pt

# 줄간격 (line, 240 = 1.0) — 정의서: 본문 1.4
LINE_BODY = "336"    # 240 × 1.4

# 페이지 (twips, 1mm ≈ 56.69tw) — 정의서: A4 + 사방 25mm
PG_W, PG_H = "11906", "16838"   # A4 (210 × 297 mm)
PG_MARGIN  = "1417"             # 25mm
PG_HF      = "708"              # 머리글/바닥글 12.5mm


# ── 엘리먼트 빌더 ───────────────────────────────────────────────
def mk(tag, **attrs):
    """w: 네임스페이스 엘리먼트 생성. 속성도 w: 네임스페이스로 부여."""
    el = etree.SubElement(etree.Element(wq("_tmp")), wq(tag))
    el.getparent().remove(el)
    for k, v in attrs.items():
        el.set(wq(k), v)
    return el

def rfonts(font):
    """ascii(영문)·hAnsi(확장)·eastAsia(한글)·cs(복합문자) 모두 같은 폰트로."""
    return mk("rFonts", ascii=font, hAnsi=font, eastAsia=font, cs=font)

def shd(fill):
    """배경색 음영. val='clear' + fill 로 깔끔한 단색 배경."""
    return mk("shd", val="clear", color="auto", fill=fill)

def border(side, sz, space, color):
    return mk(side, val="single", sz=sz, space=space, color=color)

def pbdr(*sides):
    el = mk("pBdr")
    for s in sides:
        el.append(s)
    return el

def ordered_insert(parent, child, order):
    """스키마 순서(order: localname 리스트)를 지켜 child를 parent에 삽입한다."""
    ci = order.index(etree.QName(child).localname)
    for existing in parent:
        ln = etree.QName(existing).localname
        if ln in order and order.index(ln) > ci:
            existing.addprevious(child)
            return
    parent.append(child)


def set_block(style, *, pPr_children=None, rPr_children=None):
    """스타일의 pPr/rPr을 '스키마 순서를 지켜' 통째로 교체한다.

    기존 pPr/rPr을 제거하고 새로 만들되, w:name·basedOn·next 등 메타는 보존.
    스키마상 pPr은 rPr보다 앞에 와야 하므로 그 순서로 삽입한다.
    """
    if style is None:
        return
    for old in style.findall(wq("pPr")) + style.findall(wq("rPr")):
        style.remove(old)
    if pPr_children:
        ppr = mk("pPr")
        for c in pPr_children:
            ppr.append(c)
        style.append(ppr)
    if rPr_children:
        rpr = mk("rPr")
        for c in rPr_children:
            rpr.append(c)
        style.append(rpr)


def main():
    parser = etree.XMLParser(remove_blank_text=False)

    # base.docx에서 편집 대상 3개 파트를 읽는다.
    with zipfile.ZipFile(BASE) as z:
        styles_xml = z.read("word/styles.xml")
        theme_xml = z.read("word/theme/theme1.xml")
        document_xml = z.read("word/document.xml")

    # ============ 1) styles.xml 편집 ============
    sroot = etree.fromstring(styles_xml, parser)
    def style(sid):
        for s in sroot.findall(wq("style")):
            if s.get(wq("styleId")) == sid:
                return s
        return None

    # (1-a) 문서 기본값(docDefaults): 모든 글자의 기본 폰트를 한글 폰트로
    rpr_def = sroot.find(f"{wq('docDefaults')}/{wq('rPrDefault')}/{wq('rPr')}")
    if rpr_def is not None:
        for f in rpr_def.findall(wq("rFonts")):
            rpr_def.remove(f)
        rpr_def.insert(0, rfonts(BODY_FONT))
        for lang in rpr_def.findall(wq("lang")):  # 한국어 언어 태그
            lang.set(wq("eastAsia"), "ko-KR")

    # (1-b) Normal(본문): 폰트·크기·글자색 + 줄간격 1.4
    set_block(
        style("Normal"),
        pPr_children=[mk("spacing", after="120", line=LINE_BODY, lineRule="auto")],
        rPr_children=[rfonts(BODY_FONT), mk("color", val=BODY_INK), mk("sz", val=SZ_BODY), mk("szCs", val=SZ_BODY)],
    )

    # (1-c) Heading 1 (장): Rails 레드 + 하단 0.5pt 경계선
    set_block(
        style("Heading1"),
        pPr_children=[
            mk("keepNext"), mk("keepLines"),
            pbdr(border("bottom", "4", "4", RAILS_RED)),   # 0.5pt 빨간 밑줄 = 장 구분
            mk("spacing", before="480", after="120"),
            mk("outlineLvl", val="0"),
        ],
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("color", val=RAILS_RED), mk("sz", val=SZ_H1), mk("szCs", val=SZ_H1)],
    )

    # (1-d) Heading 2 (절): 딥 레드
    set_block(
        style("Heading2"),
        pPr_children=[mk("keepNext"), mk("keepLines"), mk("spacing", before="320", after="80"), mk("outlineLvl", val="1")],
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("color", val=DEEP_RED), mk("sz", val=SZ_H2), mk("szCs", val=SZ_H2)],
    )

    # (1-e) Heading 3 (소절): 슬레이트
    set_block(
        style("Heading3"),
        pPr_children=[mk("keepNext"), mk("keepLines"), mk("spacing", before="240", after="60"), mk("outlineLvl", val="2")],
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("color", val=SLATE), mk("sz", val=SZ_H3), mk("szCs", val=SZ_H3)],
    )

    # (1-f) Heading 4: 슬레이트(작게)
    set_block(
        style("Heading4"),
        pPr_children=[mk("keepNext"), mk("keepLines"), mk("spacing", before="200", after="40"), mk("outlineLvl", val="3")],
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("color", val=SLATE), mk("sz", val=SZ_H4), mk("szCs", val=SZ_H4)],
    )

    # (1-g) Heading 5: 그레이
    set_block(
        style("Heading5"),
        pPr_children=[mk("keepNext"), mk("keepLines"), mk("spacing", before="200", after="40"), mk("outlineLvl", val="4")],
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("color", val=H5_GRAY), mk("sz", val=SZ_H5), mk("szCs", val=SZ_H5)],
    )

    # (1-h) Heading 6: 그레이 + 이탤릭
    set_block(
        style("Heading6"),
        pPr_children=[mk("keepNext"), mk("keepLines"), mk("spacing", before="200", after="40"), mk("outlineLvl", val="5")],
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("i"), mk("iCs"), mk("color", val=GRAY), mk("sz", val=SZ_H6), mk("szCs", val=SZ_H6)],
    )

    # (1-i) Title(표지 제목): Rails 레드, 가장 큼, 가운데
    set_block(
        style("Title"),
        pPr_children=[mk("keepNext"), mk("keepLines"), mk("spacing", before="480", after="120"), mk("jc", val="center")],
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("color", val=RAILS_RED), mk("sz", val=SZ_TITLE), mk("szCs", val=SZ_TITLE)],
    )

    # (1-j) Subtitle(부제): 그레이, 가운데, 굵기 해제(Title 상속)
    set_block(
        style("Subtitle"),
        pPr_children=[mk("keepNext"), mk("keepLines"), mk("spacing", before="0", after="240"), mk("jc", val="center")],
        rPr_children=[rfonts(BODY_FONT), mk("b", val="0"), mk("bCs", val="0"), mk("color", val=GRAY), mk("sz", val=SZ_SUBTITLE), mk("szCs", val=SZ_SUBTITLE)],
    )

    # (1-k) Block Text(인용문): 좌측 레드 바 + 옅은 레드 배경 + 슬레이트 글자
    set_block(
        style("BlockText"),
        pPr_children=[
            pbdr(border("left", "24", "10", RAILS_RED)),   # 좌측 3pt 강조 바
            shd(QUOTE_TINT),
            mk("spacing", before="120", after="120"),
            mk("ind", left="360", right="360"),
        ],
        rPr_children=[mk("color", val=SLATE), mk("i"), mk("iCs")],
    )

    # (1-l) Verbatim Char(인라인 코드): D2Coding + 레드 글자 + 연한 배경
    set_block(
        style("VerbatimChar"),
        rPr_children=[rfonts(CODE_FONT), mk("color", val=INLINE_RED), mk("sz", val=SZ_INLINE), mk("szCs", val=SZ_INLINE), shd(INLINE_CODE_BG)],
    )

    # (1-m) Hyperlink(링크): 블루 + 밑줄
    set_block(
        style("Hyperlink"),
        rPr_children=[mk("color", val=LINK_BLUE), mk("u", val="single")],
    )

    # (1-n) Caption(캡션): 그레이 + 이탤릭 + 가운데 (ImageCaption은 상속)
    set_block(
        style("Caption"),
        pPr_children=[mk("spacing", before="0", after="120"), mk("jc", val="center")],
        rPr_children=[rfonts(BODY_FONT), mk("i"), mk("iCs"), mk("color", val=GRAY), mk("sz", val=SZ_CAPTION), mk("szCs", val=SZ_CAPTION)],
    )

    # (1-o) TOC Heading(목차 제목): Rails 레드 + 굵게 (pPr은 Heading1 상속)
    set_block(
        style("TOCHeading"),
        rPr_children=[rfonts(BODY_FONT), mk("b"), mk("bCs"), mk("color", val=RAILS_RED), mk("sz", val=SZ_TOC), mk("szCs", val=SZ_TOC)],
    )

    # (1-p) Source Code(코드 블록): 새 스타일 추가 — D2Coding + 박스 + 배경 + 코드 잉크
    sc = mk("style", type="paragraph", styleId="SourceCode")
    sc.append(mk("name", val="Source Code"))
    sc.append(mk("basedOn", val="Normal"))
    sc.append(mk("uiPriority", val="20"))
    set_block(
        sc,
        pPr_children=[
            mk("keepLines"),
            pbdr(  # 네 변 모두 얇은 박스
                border("top", "4", "6", CODE_BORDER), border("left", "4", "6", CODE_BORDER),
                border("bottom", "4", "6", CODE_BORDER), border("right", "4", "6", CODE_BORDER),
            ),
            shd(CODE_BG),
            mk("spacing", before="120", after="120", line="240", lineRule="auto"),
            mk("ind", left="120", right="120"),
        ],
        rPr_children=[rfonts(CODE_FONT), mk("color", val=CODE_INK), mk("sz", val=SZ_CODE), mk("szCs", val=SZ_CODE)],
    )
    sroot.append(sc)

    # (1-q) Table(표): 사방·내부 격자선 + 헤더행 레드 틴트 배경·레드 굵은 글자·레드 하단선
    tbl = style("Table")
    if tbl is not None:
        TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
                       "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
                       "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
                       "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
        BORDER_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV"]

        # 표 전체 격자선: 모든 셀 사방 + 내부 0.5pt(sz=4) 옅은 회색
        tblpr = tbl.find(wq("tblPr"))
        if tblpr is None:
            tblpr = mk("tblPr")
            ordered_insert(tbl, tblpr, ["pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"])
        for old in tblpr.findall(wq("tblBorders")):
            tblpr.remove(old)
        tb = mk("tblBorders")
        for side in BORDER_SIDES:
            tb.append(border(side, "4", "0", TABLE_GRID))
        ordered_insert(tblpr, tb, TBLPR_ORDER)

        fr = None
        for tp in tbl.findall(wq("tblStylePr")):
            if tp.get(wq("type")) == "firstRow":
                fr = tp
        if fr is not None:
            TBLSTYLE_ORDER = ["pPr", "rPr", "tblPr", "trPr", "tcPr"]
            # 헤더 글자: 레드 + 굵게 (스키마 순서 보장)
            for old in fr.findall(wq("rPr")):
                fr.remove(old)
            rpr = mk("rPr"); rpr.append(mk("b")); rpr.append(mk("bCs")); rpr.append(mk("color", val=RAILS_RED))
            ordered_insert(fr, rpr, TBLSTYLE_ORDER)
            # 헤더 셀: tcPr 없으면 생성 후 삽입
            tcpr = fr.find(wq("tcPr"))
            if tcpr is None:
                tcpr = mk("tcPr")
                ordered_insert(fr, tcpr, TBLSTYLE_ORDER)
            TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge",
                          "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
                          "tcFitText", "vAlign", "hideMark"]
            # 헤더 하단: 1pt(sz=8) 레드 선 — 색만으로 부족했던 헤더/본문 구분을 또렷하게
            for old in tcpr.findall(wq("tcBorders")):
                tcpr.remove(old)
            tcb = mk("tcBorders"); tcb.append(border("bottom", "8", "0", RAILS_RED))
            ordered_insert(tcpr, tcb, TCPR_ORDER)
            # 헤더 배경 음영
            for old in tcpr.findall(wq("shd")):
                tcpr.remove(old)
            ordered_insert(tcpr, shd(TH_BG), TCPR_ORDER)

    new_styles = etree.tostring(sroot, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ============ 2) theme1.xml 편집 (테마 폰트의 한글 fallback) ============
    troot = etree.fromstring(theme_xml, parser)
    for major_minor in ["majorFont", "minorFont"]:
        font = troot.find(f".//{aq(major_minor)}")
        if font is None:
            continue
        for child in font.findall(aq("latin")) + font.findall(aq("ea")):
            child.set("typeface", BODY_FONT)
    new_theme = etree.tostring(troot, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ============ 3) document.xml 편집 (A4 + 25mm 여백) ============
    droot = etree.fromstring(document_xml, parser)
    SECTPR_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr",
                    "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
                    "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
                    "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
                    "sectPrChange"]
    for sect in droot.findall(f".//{wq('sectPr')}"):
        for old in sect.findall(wq("pgSz")) + sect.findall(wq("pgMar")):
            sect.remove(old)
        ordered_insert(sect, mk("pgSz", w=PG_W, h=PG_H), SECTPR_ORDER)
        ordered_insert(sect, mk("pgMar", top=PG_MARGIN, right=PG_MARGIN,
                                bottom=PG_MARGIN, left=PG_MARGIN,
                                header=PG_HF, footer=PG_HF, gutter="0"), SECTPR_ORDER)
    new_document = etree.tostring(droot, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ============ 4) 재압축 ============
    with zipfile.ZipFile(BASE) as zin, zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            if item == "word/styles.xml":
                zout.writestr(item, new_styles)
            elif item == "word/theme/theme1.xml":
                zout.writestr(item, new_theme)
            elif item == "word/document.xml":
                zout.writestr(item, new_document)
            else:
                zout.writestr(item, zin.read(item))

    print(f"참조 문서 생성 완료: {OUT}")
    print(f"  본문/제목 폰트: {BODY_FONT}  |  코드 폰트: {CODE_FONT}")
    print(f"  페이지: A4(210×297mm), 여백 25mm")


if __name__ == "__main__":
    main()

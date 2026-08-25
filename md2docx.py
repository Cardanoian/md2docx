#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx.py — 마크다운(.md)을 Word 문서(.docx)로 변환하는 범용 스크립트

특징
  - 마크다운 형식 전체 보존: 헤더, 볼드/이탤릭, 리스트(중첩), 표, 인용,
    각주, 링크, 수평선, 체크박스 등
  - 로컬 이미지 자동 임베드: 마크다운 파일 기준 상대경로를 알아서 해석
  - 코드 구문 강조(syntax highlighting): pandoc 내장 하이라이터 사용
  - 세부 스타일 제어: 참조 문서(reference.docx)로 폰트/헤딩/코드블록 커스터마이징
  - 추가 의존성 없음: pandoc 바이너리만 있으면 동작 (pypandoc 불필요)

사용 예시
  # 기본 변환 (출력 파일명 생략 시 같은 이름의 .docx 생성)
  python md2docx.py report.md

  # 출력 경로 지정 + 목차 + 하이라이트 스타일 변경
  python md2docx.py report.md -o ./out/report.docx --toc --highlight-style breezedark

  # 세부 스타일 제어용 참조 문서 템플릿 생성 → 워드에서 스타일 편집 후 적용
  python md2docx.py --make-reference my-template.docx
  python md2docx.py report.md --reference-doc my-template.docx

요구사항
  - pandoc 바이너리가 필요합니다. 가장 간단한 방법:
        pip install -r requirements.txt
    (pypandoc-binary가 pandoc까지 함께 설치 → OS 무관 동일 절차)
  - 시스템에 직접 설치해도 됩니다(설치돼 있으면 그쪽을 우선 사용):
        macOS:   brew install pandoc
        Ubuntu:  sudo apt install pandoc
        Windows: winget install pandoc
"""

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

# pandoc docx 출력에서 지원하는 코드 하이라이트 스타일
VALID_HIGHLIGHT_STYLES = (
    "pygments",
    "tango",
    "espresso",
    "zenburn",
    "kate",
    "monochrome",
    "breezedark",
    "haddock",
)
DEFAULT_HIGHLIGHT_STYLE = "tango"


class ConversionError(Exception):
    """변환 준비 또는 실행 중 발생한 오류."""


def find_pandoc() -> str:
    """pandoc 실행 파일 경로를 반환한다.

    탐색 순서
      1순위: 시스템 PATH에 설치된 pandoc (brew/apt/winget 등으로 설치한 경우)
      2순위: pypandoc-binary로 함께 설치된 번들 pandoc
             (pip install -r requirements.txt 만으로 OS 무관하게 동작)
    둘 다 없으면 ConversionError를 발생시킨다.
    """
    pandoc = shutil.which("pandoc")
    if pandoc:
        return pandoc
    try:
        import pypandoc  # pip install pypandoc-binary

        return pypandoc.get_pandoc_path()
    except Exception:
        pass
    raise ConversionError(
        "오류: pandoc을 찾을 수 없습니다.\n"
        "  가장 간단한 해결: pip install -r requirements.txt\n"
        "  (pypandoc-binary가 pandoc 바이너리까지 함께 설치합니다)\n\n"
        "  직접 설치를 원하면:\n"
        "    설치 안내: https://pandoc.org/installing.html\n"
        "    macOS:   brew install pandoc\n"
        "    Ubuntu:  sudo apt install pandoc\n"
        "    Windows: winget install pandoc"
    )


def make_reference_doc(pandoc: str, dest: Path) -> Path:
    """pandoc 기본 reference.docx 템플릿을 추출한다.

    이 파일을 워드에서 열어 '본문/제목 1/Source Code' 등 스타일을 편집한 뒤
    --reference-doc 옵션으로 넘기면, 변환 결과에 그대로 반영된다.
    (코드블록 배경/폰트, 헤딩 색상, 기본 글꼴 등 세부 제어의 핵심 수단)
    """
    dest = dest.expanduser().resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(dest, "wb") as f:
            subprocess.run(
                [pandoc, "--print-default-data-file", "reference.docx"],
                stdout=f,
                check=True,
            )
    except subprocess.CalledProcessError as e:
        raise ConversionError(
            "오류: 참조 문서 템플릿 생성에 실패했습니다.\n"
            f"  메시지:\n{e.stderr}"
        ) from e
    return dest


def convert(
    pandoc: str,
    src: Path,
    out: Path,
    highlight_style: str,
    reference_doc: Path | None,
    toc: bool,
    extra_resource_paths: list[str],
    toc_depth: int = 3,
) -> Path:
    """마크다운 → docx 변환 수행. 생성된 출력 경로를 반환한다."""
    src = src.expanduser().resolve()
    if not src.is_file():
        raise ConversionError(f"오류: 입력 파일을 찾을 수 없습니다 → {src}")

    out = out.expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    # 로컬 이미지 해석 기준 경로:
    #   1순위 = 마크다운 파일이 있는 디렉터리 (상대경로 이미지의 기준점)
    #   추가로 사용자가 지정한 경로들도 탐색 대상에 포함
    # OS에 따라 경로 구분자가 다르므로 os.pathsep로 합친다.
    resource_paths = [str(src.parent), *extra_resource_paths]
    resource_path = os.pathsep.join(resource_paths)

    cmd = [
        pandoc,
        str(src),
        "-o",
        str(out),
        # 입력 포맷을 GitHub 스타일 확장 마크다운으로 명시
        # (표, 각주, 체크박스, 코드펜스 등 폭넓게 지원)
        "--from",
        "gfm+footnotes+tex_math_dollars+definition_lists",
        "--to",
        "docx",
        "--highlight-style",
        highlight_style,
        "--resource-path",
        resource_path,
        # 이미지/리소스를 docx 내부에 임베드 (외부 의존성 제거)
        "--embed-resources",
        # 끊어진 참조나 깨진 링크가 있어도 변환을 중단하지 않음
        "--standalone",
    ]
    if toc:
        cmd += ["--toc", f"--toc-depth={toc_depth}"]
    if reference_doc:
        ref = reference_doc.expanduser().resolve()
        if not ref.is_file():
            raise ConversionError(f"오류: 참조 문서를 찾을 수 없습니다 → {ref}")
        cmd += ["--reference-doc", str(ref)]

    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        raise ConversionError(
            "오류: pandoc 변환에 실패했습니다.\n"
            f"  명령: {' '.join(cmd)}\n"
            f"  메시지:\n{e.stderr}"
        ) from e

    return out


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="마크다운(.md)을 Word 문서(.docx)로 변환합니다 (pandoc 기반).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("input", nargs="?", help="입력 마크다운 파일 경로(.md)")
    p.add_argument(
        "--gui",
        action="store_true",
        help="그래픽 인터페이스(Tkinter)로 실행",
    )
    p.add_argument(
        "-o",
        "--output",
        help="출력 docx 경로 (생략 시 입력 파일과 같은 이름의 .docx)",
    )
    p.add_argument(
        "--highlight-style",
        default=DEFAULT_HIGHLIGHT_STYLE,
        choices=VALID_HIGHLIGHT_STYLES,
        help=f"코드 하이라이트 스타일 (기본: {DEFAULT_HIGHLIGHT_STYLE})",
    )
    p.add_argument(
        "--reference-doc",
        help="세부 스타일 제어용 참조 docx 경로",
    )
    p.add_argument(
        "--toc",
        action="store_true",
        help="문서 앞에 목차(Table of Contents) 자동 생성",
    )
    p.add_argument(
        "--resource-path",
        action="append",
        default=[],
        help="이미지 등 리소스를 추가로 탐색할 디렉터리(여러 번 지정 가능)",
    )
    p.add_argument(
        "--make-reference",
        metavar="DEST",
        help="참조 문서 템플릿(reference.docx)을 DEST 경로에 생성하고 종료",
    )
    return p


def _launch_gui() -> None:
    from gui import run_app

    run_app()


def main(argv: list[str] | None = None) -> None:
    args = build_parser().parse_args(argv)

    if args.gui or (not args.input and not args.make_reference):
        _launch_gui()
        return

    try:
        pandoc = find_pandoc()

        # 참조 문서 템플릿 생성 모드
        if args.make_reference:
            dest = make_reference_doc(pandoc, Path(args.make_reference))
            print(f"참조 문서 템플릿 생성 완료: {dest}")
            print("  → 워드에서 스타일(특히 'Source Code', 'Heading 1~6')을 편집한 뒤")
            print(f'    --reference-doc "{dest}" 로 적용하세요.')
            return

        src = Path(args.input)
        out = Path(args.output) if args.output else src.with_suffix(".docx")

        result = convert(
            pandoc=pandoc,
            src=src,
            out=out,
            highlight_style=args.highlight_style,
            reference_doc=Path(args.reference_doc) if args.reference_doc else None,
            toc=args.toc,
            extra_resource_paths=args.resource_path,
        )
        print(f"변환 완료: {src.name} → {result}")
    except ConversionError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()

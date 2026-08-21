#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2hwpx.py — 마크다운(.md)을 한글 문서(.hwpx)로 변환하는 범용 스크립트

특징
  - 마크다운 형식 보존: 헤더, 볼드/이탤릭, 리스트(중첩), 표, 인용,
    각주, 링크, 수평선, 코드블록 등
  - 로컬 이미지 자동 임베드: 마크다운 파일 기준 상대경로를 알아서 해석
  - 세부 스타일 제어: 참조 문서(reference.hwpx)로 폰트/헤딩/페이지 커스터마이징
  - 한컴오피스 설치 불필요: 순수 파이썬으로 HWPX를 생성

사용 예시
  # 기본 변환 (출력 파일명 생략 시 같은 이름의 .hwpx 생성)
  python md2hwpx.py report.md

  # 출력 경로 지정
  python md2hwpx.py report.md -o ./out/report.hwpx

  # 세부 스타일 제어용 참조 문서 템플릿 생성 → 한글에서 스타일 편집 후 적용
  python md2hwpx.py --make-reference my-template.hwpx
  python md2hwpx.py report.md --reference-doc my-template.hwpx

요구사항
  - pip install -r requirements.txt  (md2hwpx 패키지가 함께 설치됩니다)
"""

from __future__ import annotations

import argparse
import logging
import shutil
import sys
import tempfile
from pathlib import Path

# 이 파일이 md2hwpx.py 라서, 그대로 import 하면 설치된 패키지 대신
# 자기 자신을 불러온다. 패키지 import 전에 스크립트 디렉터리를 경로에서 뺀다.
_SCRIPT_DIR = str(Path(__file__).resolve().parent)
sys.path = [p for p in sys.path if p not in ("", _SCRIPT_DIR)]

logging.getLogger("md2hwpx").setLevel(logging.ERROR)


def require_md2hwpx():
    """md2hwpx 패키지를 불러온다. 없으면 설치 방법을 안내하고 종료한다."""
    try:
        import md2hwpx
        from md2hwpx.MarkdownToHwpx import MarkdownToHwpx
        from md2hwpx.config import ConversionConfig
        from md2hwpx.exceptions import HwpxError
        from md2hwpx.frontmatter_parser import (
            convert_metadata_to_pandoc_meta,
            parse_markdown_with_frontmatter,
        )
        from md2hwpx.marko_adapter import MarkoToPandocAdapter
    except ImportError:
        sys.exit(
            "오류: md2hwpx 패키지를 찾을 수 없습니다.\n"
            "  가장 간단한 해결: pip install -r requirements.txt"
        )
    return (
        md2hwpx,
        MarkdownToHwpx,
        ConversionConfig,
        HwpxError,
        convert_metadata_to_pandoc_meta,
        parse_markdown_with_frontmatter,
        MarkoToPandocAdapter,
    )


def default_reference_path(md2hwpx_mod) -> Path:
    pkg_dir = Path(md2hwpx_mod.__path__[0])
    return pkg_dir / "blank.hwpx"


def make_reference_doc(dest: Path) -> None:
    """기본 참조 HWPX 템플릿을 추출한다.

    이 파일을 한글에서 열어 스타일을 편집한 뒤 --reference-doc 으로 넘기면
    변환 결과에 반영된다.
    """
    md2hwpx_mod, *_ = require_md2hwpx()
    src = default_reference_path(md2hwpx_mod)
    if not src.is_file():
        sys.exit(f"오류: 기본 참조 템플릿을 찾을 수 없습니다 → {src}")

    dest = dest.expanduser().resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    print(f"참조 문서 템플릿 생성 완료: {dest}")
    print("  → 한글에서 스타일(제목, 본문, 표 등)을 편집한 뒤")
    print(f'    --reference-doc "{dest}" 로 적용하세요.')


def _merge_dir_contents(src_dir: Path, dest_dir: Path) -> None:
    """src_dir의 파일/폴더를 dest_dir에 병합 복사한다."""
    for item in src_dir.iterdir():
        dest = dest_dir / item.name
        if item.is_dir():
            shutil.copytree(item, dest, dirs_exist_ok=True)
        elif item.is_file():
            shutil.copy2(item, dest)


def convert(
    src: Path,
    out: Path,
    reference_doc: Path | None,
    extra_resource_paths: list[str],
) -> None:
    """마크다운 → hwpx 변환 수행."""
    (
        md2hwpx_mod,
        MarkdownToHwpx,
        ConversionConfig,
        HwpxError,
        convert_metadata_to_pandoc_meta,
        parse_markdown_with_frontmatter,
        MarkoToPandocAdapter,
    ) = require_md2hwpx()

    src = src.expanduser().resolve()
    if not src.is_file():
        sys.exit(f"오류: 입력 파일을 찾을 수 없습니다 → {src}")

    out = out.expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    if reference_doc:
        ref = reference_doc.expanduser().resolve()
        if not ref.is_file():
            sys.exit(f"오류: 참조 문서를 찾을 수 없습니다 → {ref}")
    else:
        ref = default_reference_path(md2hwpx_mod)
        if not ref.is_file():
            sys.exit(f"오류: 기본 참조 템플릿을 찾을 수 없습니다 → {ref}")

    work_src = src
    tmp_dir: tempfile.TemporaryDirectory[str] | None = None
    extras = [Path(p).expanduser().resolve() for p in extra_resource_paths if p]
    extras = [p for p in extras if p.is_dir()]

    try:
        if extras:
            tmp_dir = tempfile.TemporaryDirectory(prefix="md2hwpx_")
            work_dir = Path(tmp_dir.name)
            _merge_dir_contents(src.parent, work_dir)
            for extra in extras:
                _merge_dir_contents(extra, work_dir)
            work_src = work_dir / src.name
            if not work_src.is_file():
                sys.exit(f"오류: 작업 복사본을 만들 수 없습니다 → {work_src}")

        metadata, md_content = parse_markdown_with_frontmatter(str(work_src))
        adapter = MarkoToPandocAdapter(config=ConversionConfig())
        ast = adapter.parse(md_content)
        ast["meta"] = convert_metadata_to_pandoc_meta(metadata)

        MarkdownToHwpx.convert_to_hwpx(
            input_path=str(work_src),
            output_path=str(out),
            reference_path=str(ref),
            json_ast=ast,
            config=ConversionConfig(),
        )
    except HwpxError as e:
        sys.exit(f"오류: HWPX 변환에 실패했습니다.\n  메시지: {e}")
    finally:
        if tmp_dir is not None:
            tmp_dir.cleanup()

    print(f"변환 완료: {src.name} → {out}")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="마크다운(.md)을 한글 문서(.hwpx)로 변환합니다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("input", nargs="?", help="입력 마크다운 파일 경로(.md)")
    p.add_argument(
        "-o",
        "--output",
        help="출력 hwpx 경로 (생략 시 입력 파일과 같은 이름의 .hwpx)",
    )
    p.add_argument(
        "--reference-doc",
        help="세부 스타일 제어용 참조 hwpx 경로",
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
        help="참조 문서 템플릿(blank.hwpx)을 DEST 경로에 생성하고 종료",
    )
    return p


def main(argv: list[str] | None = None) -> None:
    args = build_parser().parse_args(argv)

    if args.make_reference:
        make_reference_doc(Path(args.make_reference))
        return

    if not args.input:
        sys.exit("오류: 입력 마크다운 파일을 지정하세요. (-h 로 도움말 확인)")

    src = Path(args.input)
    out = Path(args.output) if args.output else src.with_suffix(".hwpx")

    convert(
        src=src,
        out=out,
        reference_doc=Path(args.reference_doc) if args.reference_doc else None,
        extra_resource_paths=args.resource_path,
    )


if __name__ == "__main__":
    main()

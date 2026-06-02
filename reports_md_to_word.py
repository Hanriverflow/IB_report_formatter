#!/usr/bin/env python3
"""
Dedicated Reports Markdown to Word CLI Converter.

This script specializes in converting Markdown files (.md) in the 'reports/' folder
to professional IB-styled Word documents (.docx). It supports interactive selection,
individual file targeting, and batch conversion.

Usage:
    uv run convert-reports                          # Interactive selection mode
    uv run convert-reports hanwa_off_v3.md          # Convert specific file inside reports/
    uv run convert-reports --batch                  # Convert all markdown files in reports/
    uv run convert-reports reports/my_report.md     # Convert via path
"""

import os
import sys
import time
import argparse
import logging
from pathlib import Path
from typing import List, Optional, Tuple

# Re-use existing converter logic and configuration
from md_to_word import run_conversion
from cli_utils import setup_logging

# ANSI Colors for premium terminal aesthetics
COLOR_NAVY = "\033[1;34m"
COLOR_GREEN = "\033[1;32m"
COLOR_YELLOW = "\033[1;33m"
COLOR_RED = "\033[1;31m"
COLOR_GRAY = "\033[0;90m"
COLOR_CYAN = "\033[1;36m"
COLOR_RESET = "\033[0m"
COLOR_BOLD = "\033[1m"

logger = logging.getLogger("reports_md_to_word")


def print_header():
    """Prints a beautiful double-line bordered header matching the IB styling."""
    os.system('cls' if os.name == 'nt' else 'clear')
    print(f"{COLOR_NAVY}╔═════════════════════════════════════════════════════════════════╗{COLOR_RESET}")
    print(f"{COLOR_NAVY}║              📊  REPORTS MD ➔ WORD CONVERTER CLI  📊              ║{COLOR_RESET}")
    print(f"{COLOR_NAVY}║             Dedicated Markdown to Word Conversion               ║{COLOR_RESET}")
    print(f"{COLOR_NAVY}╚═════════════════════════════════════════════════════════════════╝{COLOR_RESET}")
    print()


def get_reports_dir() -> Path:
    """Get the absolute path to the reports/ directory."""
    root_dir = Path(__file__).resolve().parent
    return root_dir / "reports"


def scan_markdown_reports() -> Tuple[List[Path], int]:
    """
    Scans the reports/ folder for Markdown files.
    Filters out intermediate/generated files to keep the listing clean.
    """
    reports_dir = get_reports_dir()
    
    if not reports_dir.exists() or not reports_dir.is_dir():
        return [], 0

    all_md = sorted(reports_dir.glob("*.md"))
    
    # Filter out generated files
    md_files = [
        f for f in all_md 
        if not f.name.endswith("_formatted.md") and not f.name.endswith("_cleaned.md")
    ]
    
    filtered_count = len(all_md) - len(md_files)
    return md_files, filtered_count


def format_size(byte_size: int) -> str:
    """Formats file size in KB."""
    return f"{byte_size / 1024:.1f} KB"


def format_mtime(timestamp: float) -> str:
    """Formats modification time to yyyy-mm-dd HH:MM."""
    return time.strftime("%Y-%m-%d %H:%M", time.localtime(timestamp))


def run_interactive_menu(args: argparse.Namespace) -> int:
    """Renders the interactive file selection menu."""
    if os.name == 'nt':
        os.system('')  # Enable ANSI terminal styling under Windows CMD

    while True:
        print_header()
        
        md_files, filter_count = scan_markdown_reports()
        
        if not md_files:
            print(f"{COLOR_RED}[!] 변환 가능한 마크다운 파일이 reports/ 폴더에 존재하지 않습니다.{COLOR_RESET}")
            print(f"경로: {get_reports_dir()}")
            return 1
            
        print(f"{COLOR_BOLD}📁 [reports/] 폴더 내 원본 마크다운 파일 목록 (총 {len(md_files)}개){COLOR_RESET}")
        if filter_count > 0:
            print(f"{COLOR_GRAY}(* 자동 포맷된 임시 파일 {filter_count}개는 목록에서 숨김 처리됨){COLOR_RESET}")
        print()

        print(f"{COLOR_NAVY}┌────────────────────────────────────────────────────────────────────────┐{COLOR_RESET}")
        print(f"{COLOR_NAVY}│ 📄 원본 마크다운 파일 (*.md)                                           │{COLOR_RESET}")
        print(f"{COLOR_NAVY}└────────────────────────────────────────────────────────────────────────┘{COLOR_RESET}")
        
        for idx, f in enumerate(md_files, 1):
            size_str = format_size(f.stat().st_size)
            mtime_str = format_mtime(f.stat().st_mtime)
            print(f"  [{COLOR_GREEN}{idx:2d}{COLOR_RESET}]  {f.name:<40s}  {COLOR_GRAY}| {size_str:>8s} | {mtime_str}{COLOR_RESET}")
        print()
        
        print(f"  [{COLOR_RED}q{COLOR_RESET}]  종료 (Quit)")
        print()
        
        user_input = input("👉 변환할 파일 번호를 입력하세요: ").strip().lower()
        
        if user_input in ('q', 'quit', 'exit'):
            print(f"\n{COLOR_YELLOW}프로그램을 종료합니다. 감사합니다!{COLOR_RESET}\n")
            return 0
            
        if not user_input:
            continue
            
        try:
            choice_idx = int(user_input)
        except ValueError:
            print(f"{COLOR_RED}잘못된 입력입니다. 숫자 번호 또는 'q'를 입력해 주세요.{COLOR_RESET}")
            time.sleep(1.2)
            continue
            
        if choice_idx < 1 or choice_idx > len(md_files):
            print(f"{COLOR_RED}범위를 벗어난 번호입니다. 목록에 있는 번호를 입력해 주세요.{COLOR_RESET}")
            time.sleep(1.2)
            continue
            
        selected_file = md_files[choice_idx - 1]
        
        # Submenu for Conversion Options
        print_header()
        print(f"{COLOR_CYAN}📄 선택된 파일: {COLOR_BOLD}{selected_file.name}{COLOR_RESET}")
        print(f"{COLOR_GRAY}경로: {selected_file}{COLOR_RESET}\n")
        
        print(f"{COLOR_NAVY}┌─────────────────────────────────────────────────────────┐{COLOR_RESET}")
        print(f"{COLOR_NAVY}│       변환 옵션을 선택하세요 (Markdown ➔ Word DOCX)      │{COLOR_RESET}")
        print(f"{COLOR_NAVY}└─────────────────────────────────────────────────────────┘{COLOR_RESET}")
        print(f"  {COLOR_GREEN}[1]{COLOR_RESET} 기본 변환 (Standard Conversion)")
        print(f"      - 일반적인 마크다운 구조를 정교한 IB 스타일 Word 문서로 변환")
        print(f"  {COLOR_GREEN}[2]{COLOR_RESET} 자동 사전 포맷팅 적용 변환 (--format)")
        print(f"      - 단일 라인으로 뭉쳐 있거나 AI가 내보낸 비구조화 텍스트 정리 후 변환")
        print(f"  {COLOR_GREEN}[3]{COLOR_RESET} DeepResearch 정리 및 자동 포맷팅 변환 (최고 정밀 모드)")
        print(f"      - 딥리서치 특수 마커 정리, Footnote 인용 변환, 사전 포맷팅 모두 적용")
        print(f"  {COLOR_GRAY}[b]{COLOR_RESET} 뒤로 가기 (이전 메뉴로)")
        print()
        
        opt_choice = input("👉 선택하실 번호를 입력하세요: ").strip().lower()
        
        if opt_choice == 'b':
            continue
            
        # Clone args and apply selected option
        file_args = argparse.Namespace(**vars(args))
        file_args.input_file = str(selected_file)
        
        if opt_choice == '1':
            file_args.format = False
            file_args.deepresearch_cleaner = "off"
        elif opt_choice == '2':
            file_args.format = True
            file_args.deepresearch_cleaner = "off"
        elif opt_choice == '3':
            file_args.format = True
            file_args.deepresearch_cleaner = "auto"
            file_args.cite_mode = "footnote"
            file_args.cleaner_report = True
        else:
            print(f"{COLOR_RED}잘못된 선택입니다. 다시 입력해 주세요.{COLOR_RESET}")
            time.sleep(1)
            continue
            
        exit_code = run_conversion(selected_file, file_args)
        if exit_code == 0:
            print(f"\n{COLOR_GREEN}✓ 성공적으로 변환되었습니다!{COLOR_RESET}")
        else:
            print(f"\n{COLOR_RED}✗ 변환 중 오류가 발생했습니다.{COLOR_RESET}")
            
        input("\n계속하려면 [Enter] 키를 누르세요...")


def run_batch_conversion(args: argparse.Namespace) -> int:
    """Executes batch conversion for all .md files in reports/."""
    md_files, _ = scan_markdown_reports()
    
    if not md_files:
        logger.error("No markdown files found in reports/ directory.")
        return 1
        
    print(f"\n{COLOR_BOLD}🚀 reports/ 내의 모든 마크다운 파일 일괄 변환 시작 (총 {len(md_files)}개){COLOR_RESET}")
    print(f"{COLOR_GRAY}{'=' * 65}{COLOR_RESET}")
    
    success_count = 0
    failure_count = 0
    
    for f in md_files:
        print(f"{COLOR_CYAN}➔ 변환 중: {f.name}{COLOR_RESET}")
        file_args = argparse.Namespace(**vars(args))
        file_args.input_file = str(f)
        
        # Determine output file path
        if args.output_file:
            # For batch, if user specified output_file, we treat it as output directory
            out_dir = Path(args.output_file)
            out_dir.mkdir(parents=True, exist_ok=True)
            # Standard auto naming logic for output
            from md_to_word import generate_output_path as build_output_path
            out_path = build_output_path(f)
            file_args.output_file = str(out_dir / out_path.name)
        else:
            file_args.output_file = None
            
        exit_code = run_conversion(f, file_args)
        if exit_code == 0:
            success_count += 1
        else:
            failure_count += 1
            
    print(f"\n{COLOR_GRAY}{'=' * 65}{COLOR_RESET}")
    print(f"{COLOR_GREEN}✓ 변환 완료: 성공 {success_count}개{COLOR_RESET}, {COLOR_RED}실패 {failure_count}개{COLOR_RESET}")
    return 0 if failure_count == 0 else 1


def build_parser() -> argparse.ArgumentParser:
    """Builds the argument parser for reports-to-word utility."""
    parser = argparse.ArgumentParser(
        description="Convert Markdown reports in reports/ folder to professional IB-style Word documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    uv run convert-reports                         # Interactive selector
    uv run convert-reports hanwa_off_v3.md         # Direct convert (resolved inside reports/)
    uv run convert-reports hanwa_off_v3            # Omit extension
    uv run convert-reports --batch                 # Convert all reports
    uv run convert-reports hanwa_off_v3 --format   # Pre-format before converting
        """,
    )
    
    parser.add_argument(
        "input_file",
        nargs="?",
        help="Markdown report filename or path (looks in reports/ if relative name)",
    )
    
    parser.add_argument(
        "output_file",
        nargs="?",
        help="Output Word document path (optional, defaults to reports/ directory)",
    )
    
    # CLI mode flags
    mode_group = parser.add_argument_group("mode options")
    mode_group.add_argument(
        "-i",
        "--interactive",
        action="store_true",
        help="Force interactive report selector mode",
    )
    mode_group.add_argument(
        "-b",
        "--batch",
        "--all",
        action="store_true",
        dest="batch",
        help="Convert all markdown files inside the reports/ folder",
    )

    # Pre-processing options
    preprocess_group = parser.add_argument_group("pre-processing options")
    preprocess_group.add_argument(
        "-f",
        "--format",
        action="store_true",
        help="Auto-format markdown (single-line ➔ structured) before conversion",
    )
    preprocess_group.add_argument(
        "--deepresearch-cleaner",
        choices=["off", "auto", "on"],
        default="off",
        dest="deepresearch_cleaner",
        help="Apply OpenAI DeepResearch marker cleaner",
    )
    preprocess_group.add_argument(
        "--cite-mode",
        choices=["footnote", "inline", "strip"],
        default="footnote",
        help="How to transform cite markers when cleaner is enabled",
    )
    preprocess_group.add_argument(
        "--drop-unknown-markers",
        action="store_true",
        help="Drop unknown DeepResearch marker blocks",
    )
    preprocess_group.add_argument(
        "--cleaner-report",
        action="store_true",
        help="Print DeepResearch cleaner summary",
    )

    # Styling and section toggles
    section_group = parser.add_argument_group("section controls")
    section_group.add_argument(
        "--no-cover",
        action="store_true",
        help="Skip cover page",
    )
    section_group.add_argument(
        "--no-toc",
        action="store_true",
        help="Skip table of contents",
    )
    section_group.add_argument(
        "--no-disclaimer",
        action="store_true",
        help="Skip disclaimer page",
    )
    section_group.add_argument(
        "--separator-mode",
        choices=["auto", "rule", "page-break"],
        default="auto",
        help="Separator line rendering mode (rule, page-break, or auto)",
    )
    section_group.add_argument(
        "--style",
        choices=["classic", "ib-pro"],
        default="classic",
        help="Output style profile: classic (current output) or ib-pro (IB-grade styling)",
    )

    # Verbosity
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose (debug) logging",
    )

    return parser


def main():
    """Main execution entrypoint."""
    parser = build_parser()
    args = parser.parse_args()

    setup_logging(verbose=args.verbose)

    reports_dir = get_reports_dir()

    # Determine execution flow
    if args.interactive or (not args.input_file and not args.batch):
        sys.exit(run_interactive_menu(args))
        
    if args.batch:
        sys.exit(run_batch_conversion(args))

    # Resolve direct argument input file
    input_str = args.input_file
    input_path = Path(input_str)
    
    # Heuristics:
    # 1. Check if it's already an absolute or relative file path that exists
    if input_path.exists() and input_path.is_file():
        resolved_path = input_path
    else:
        # 2. Check in reports/ folder with original name
        resolved_path = reports_dir / input_str
        if not resolved_path.exists() or not resolved_path.is_file():
            # 3. Check in reports/ folder with .md appended
            resolved_path = reports_dir / f"{input_str}.md"
            if not resolved_path.exists() or not resolved_path.is_file():
                logger.error("Could not find file: %s (Also searched in reports/)", input_str)
                sys.exit(1)

    sys.exit(run_conversion(resolved_path, args))


if __name__ == "__main__":
    main()

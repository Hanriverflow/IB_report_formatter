#!/usr/bin/env python3
"""
Interactive CLI Report Selector and Converter for IB Report Formatter.

This script scans the 'reports/' folder for Markdown and Word documents,
renders a beautiful terminal dashboard, and allows users to select files
and convert them with customized rendering options.
"""

import os
import sys
import time
import subprocess
import logging
from pathlib import Path
from typing import List, Tuple, Optional

# Setup direct terminal prints with ANSI colors
COLOR_NAVY = "\033[1;34m"
COLOR_GREEN = "\033[1;32m"
COLOR_YELLOW = "\033[1;33m"
COLOR_RED = "\033[1;31m"
COLOR_GRAY = "\033[0;90m"
COLOR_CYAN = "\033[1;36m"
COLOR_RESET = "\033[0m"
COLOR_BOLD = "\033[1m"

def print_header():
    """Prints a beautiful double-line bordered header matching the IB styling."""
    os.system('cls' if os.name == 'nt' else 'clear')
    print(f"{COLOR_NAVY}╔═════════════════════════════════════════════════════════════════╗{COLOR_RESET}")
    print(f"{COLOR_NAVY}║                 📊  IB REPORT FORMATTER CLI  📊                 ║{COLOR_RESET}")
    print(f"{COLOR_NAVY}║          Interactive Report Selector & Converter Tool           ║{COLOR_RESET}")
    print(f"{COLOR_NAVY}╚═════════════════════════════════════════════════════════════════╝{COLOR_RESET}")
    print()

def scan_reports() -> Tuple[List[Path], List[Path], int, int]:
    """
    Scans the reports/ folder for Markdown and Word documents.
    Filters out intermediate/generated files to keep the listing clean.
    """
    root_dir = Path(__file__).resolve().parent
    reports_dir = root_dir / "reports"
    
    if not reports_dir.exists() or not reports_dir.is_dir():
        print(f"{COLOR_RED}[오류] 'reports/' 디렉토리가 존재하지 않습니다.{COLOR_RESET}")
        print(f"경로: {reports_dir}")
        return [], [], 0, 0

    all_md = sorted(reports_dir.glob("*.md"))
    all_docx = sorted(reports_dir.glob("*.docx"))

    # Smart filtering for generated files
    md_files = [
        f for f in all_md 
        if not f.name.endswith("_formatted.md") and not f.name.endswith("_cleaned.md")
    ]
    docx_files = [
        f for f in all_docx 
        if not f.name.endswith("_Report_Pro.docx")
    ]

    filtered_md_count = len(all_md) - len(md_files)
    filtered_docx_count = len(all_docx) - len(docx_files)

    return md_files, docx_files, filtered_md_count, filtered_docx_count

def format_size(byte_size: int) -> str:
    """Formats file size in KB."""
    return f"{byte_size / 1024:.1f} KB"

def format_mtime(timestamp: float) -> str:
    """Formats modification time to yyyy-mm-dd HH:MM."""
    return time.strftime("%Y-%m-%d %H:%M", time.localtime(timestamp))

def run_command(args: List[str]) -> bool:
    """Executes a command and displays the outputs clearly."""
    print(f"\n{COLOR_GRAY}[RUNNING] {' '.join(args)}{COLOR_RESET}\n")
    try:
        # Launch process using current Python interpreter
        result = subprocess.run(args, check=False)
        return result.returncode == 0
    except Exception as e:
        print(f"{COLOR_RED}[실행 오류] 명령 실행 중 오류가 발생했습니다: {e}{COLOR_RESET}")
        return False

def handle_markdown_conversion(file_path: Path):
    """Submenu for Markdown to Word conversion options."""
    root_dir = Path(__file__).resolve().parent
    script_path = root_dir / "md_to_word.py"
    
    while True:
        print_header()
        print(f"{COLOR_CYAN}📄 선택된 파일: {COLOR_BOLD}{file_path.name}{COLOR_RESET}")
        print(f"{COLOR_GRAY}경로: {file_path}{COLOR_RESET}\n")
        
        print(f"{COLOR_NAVY}┌─────────────────────────────────────────────────────────┐{COLOR_RESET}")
        print(f"{COLOR_NAVY}│       변환 옵션을 선택하세요 (Markdown → Word DOCX)      │{COLOR_RESET}")
        print(f"{COLOR_NAVY}└─────────────────────────────────────────────────────────┘{COLOR_RESET}")
        print(f"  {COLOR_GREEN}[1]{COLOR_RESET} 기본 변환 (Standard Conversion)")
        print(f"      - 일반적인 마크다운 구조를 정교한 IB 스타일 Word 문서로 변환")
        print(f"  {COLOR_GREEN}[2]{COLOR_RESET} 자동 사전 포맷팅 적용 변환 (--format)")
        print(f"      - 단일 라인으로 뭉쳐 있거나 AI가 내보낸 비구조화 텍스트 정리 후 변환")
        print(f"  {COLOR_GREEN}[3]{COLOR_RESET} DeepResearch 정리 및 자동 포맷팅 변환 (최고 정밀 모드)")
        print(f"      - 딥리서치 특수 마커 정리, Footnote 인용 변환, 사전 포맷팅 모두 적용")
        print(f"  {COLOR_GRAY}[b]{COLOR_RESET} 뒤로 가기 (이전 메뉴로)")
        print()
        
        choice = input("👉 선택하실 번호를 입력하세요: ").strip().lower()
        
        if choice == 'b':
            return
        
        cmd = [sys.executable, str(script_path), str(file_path)]
        
        if choice == '1':
            # Standard conversion
            pass
        elif choice == '2':
            cmd.append("--format")
        elif choice == '3':
            cmd.extend([
                "--format", 
                "--deepresearch-cleaner", "auto", 
                "--cite-mode", "footnote", 
                "--cleaner-report"
            ])
        else:
            print(f"{COLOR_RED}잘못된 선택입니다. 다시 입력해 주세요.{COLOR_RESET}")
            time.sleep(1)
            continue
            
        success = run_command(cmd)
        if success:
            print(f"\n{COLOR_GREEN}✓ 성공적으로 변환되었습니다!{COLOR_RESET}")
        else:
            print(f"\n{COLOR_RED}✗ 변환 중 오류가 발생했습니다.{COLOR_RESET}")
            
        input("\n계속하려면 [Enter] 키를 누르세요...")
        return

def handle_word_conversion(file_path: Path):
    """Submenu for Word to Markdown conversion options."""
    root_dir = Path(__file__).resolve().parent
    script_path = root_dir / "word_to_md.py"
    
    while True:
        print_header()
        print(f"{COLOR_CYAN}📄 선택된 파일: {COLOR_BOLD}{file_path.name}{COLOR_RESET}")
        print(f"{COLOR_GRAY}경로: {file_path}{COLOR_RESET}\n")
        
        print(f"{COLOR_NAVY}┌─────────────────────────────────────────────────────────┐{COLOR_RESET}")
        print(f"{COLOR_NAVY}│       변환 옵션을 선택하세요 (Word DOCX → Markdown)      │{COLOR_RESET}")
        print(f"{COLOR_NAVY}└─────────────────────────────────────────────────────────┘{COLOR_RESET}")
        print(f"  {COLOR_GREEN}[1]{COLOR_RESET} 기본 변환 (Standard MD Extraction)")
        print(f"      - 서식과 본문을 유지하여 일반 마크다운 파일로 추출")
        print(f"  {COLOR_GREEN}[2]{COLOR_RESET} LLM 최적화 평문 추출 (--strip --no-frontmatter)")
        print(f"      - 볼드, 이탤릭 등 불필요한 메타 마커와 frontmatter를 제외하고 깔끔한 평문 추출")
        print(f"  {COLOR_GREEN}[3]{COLOR_RESET} 이미지 별도 폴더 추출 포함 변환 (--extract-images)")
        print(f"      - 문서 내 포함된 이미지들을 별도 폴더에 저장하고 상대 경로 링크 삽입")
        print(f"  {COLOR_GREEN}[4]{COLOR_RESET} 이미지 Base64 인라인 임베딩 변환 (--embed-images-base64)")
        print(f"      - 이미지를 마크다운 내부에 Base64 데이터 스트림으로 포팅 (단일 파일 유지)")
        print(f"  {COLOR_GRAY}[b]{COLOR_RESET} 뒤로 가기 (이전 메뉴로)")
        print()
        
        choice = input("👉 선택하실 번호를 입력하세요: ").strip().lower()
        
        if choice == 'b':
            return
        
        cmd = [sys.executable, str(script_path), str(file_path)]
        
        if choice == '1':
            pass
        elif choice == '2':
            cmd.extend(["--strip", "--no-frontmatter"])
        elif choice == '3':
            cmd.append("--extract-images")
        elif choice == '4':
            cmd.append("--embed-images-base64")
        else:
            print(f"{COLOR_RED}잘못된 선택입니다. 다시 입력해 주세요.{COLOR_RESET}")
            time.sleep(1)
            continue
            
        success = run_command(cmd)
        if success:
            print(f"\n{COLOR_GREEN}✓ 성공적으로 변환되었습니다!{COLOR_RESET}")
        else:
            print(f"\n{COLOR_RED}✗ 변환 중 오류가 발생했습니다.{COLOR_RESET}")
            
        input("\n계속하려면 [Enter] 키를 누르세요...")
        return

def main():
    """Main CLI execution loop."""
    # Ensure Windows console ANSI escape codes are enabled
    if os.name == 'nt':
        os.system('')

    while True:
        print_header()
        
        md_files, docx_files, filter_md_count, filter_docx_count = scan_reports()
        
        # Display File Lists
        total_files = len(md_files) + len(docx_files)
        
        print(f"{COLOR_BOLD}📁 [reports/] 폴더 내 변환 가능한 원본 파일 목록 (총 {total_files}개){COLOR_RESET}")
        if filter_md_count > 0 or filter_docx_count > 0:
            print(f"{COLOR_GRAY}(* 자동 생성된 결과물 및 중간 단계 파일 {filter_md_count + filter_docx_count}개는 목록에서 숨김 처리됨){COLOR_RESET}")
        print()

        index_map = {}
        idx = 1

        # 1. Show Markdown files
        print(f"{COLOR_NAVY}┌────────────────────────────────────────────────────────────────────────┐{COLOR_RESET}")
        print(f"{COLOR_NAVY}│ 📄 Category A: Markdown 원본 파일 (*.md)                              │{COLOR_RESET}")
        print(f"{COLOR_NAVY}└────────────────────────────────────────────────────────────────────────┘{COLOR_RESET}")
        if not md_files:
            print(f"   {COLOR_GRAY}변환 가능한 마크다운 파일이 없습니다.{COLOR_RESET}")
        else:
            for f in md_files:
                size_str = format_size(f.stat().st_size)
                mtime_str = format_mtime(f.stat().st_mtime)
                print(f"  [{COLOR_GREEN}{idx:2d}{COLOR_RESET}]  {f.name:<40s}  {COLOR_GRAY}| {size_str:>8s} | {mtime_str}{COLOR_RESET}")
                index_map[idx] = (f, "md")
                idx += 1
        print()

        # 2. Show Word documents
        print(f"{COLOR_NAVY}┌────────────────────────────────────────────────────────────────────────┐{COLOR_RESET}")
        print(f"{COLOR_NAVY}│ 📘 Category B: Word 원본 문서 (*.docx)                                 │{COLOR_RESET}")
        print(f"{COLOR_NAVY}└────────────────────────────────────────────────────────────────────────┘{COLOR_RESET}")
        if not docx_files:
            print(f"   {COLOR_GRAY}변환 가능한 워드 문서가 없습니다.{COLOR_RESET}")
        else:
            for f in docx_files:
                size_str = format_size(f.stat().st_size)
                mtime_str = format_mtime(f.stat().st_mtime)
                print(f"  [{COLOR_GREEN}{idx:2d}{COLOR_RESET}]  {f.name:<40s}  {COLOR_GRAY}| {size_str:>8s} | {mtime_str}{COLOR_RESET}")
                index_map[idx] = (f, "docx")
                idx += 1
        print()

        print(f"  [{COLOR_RED}q{COLOR_RESET}]  종료 (Quit)")
        print()

        user_input = input("👉 변환할 파일 번호를 입력하세요: ").strip().lower()

        if user_input == 'q' or user_input == 'quit' or user_input == 'exit':
            print(f"\n{COLOR_YELLOW}프로그램을 종료합니다. 감사합니다!{COLOR_RESET}\n")
            break

        if not user_input:
            continue

        try:
            choice_idx = int(user_input)
        except ValueError:
            print(f"{COLOR_RED}잘못된 입력입니다. 숫자 번호 또는 'q'를 입력해 주세요.{COLOR_RESET}")
            time.sleep(1)
            continue

        if choice_idx not in index_map:
            print(f"{COLOR_RED}범위를 벗어난 번호입니다. 목록에 있는 번호를 입력해 주세요.{COLOR_RESET}")
            time.sleep(1)
            continue

        selected_file, file_type = index_map[choice_idx]
        
        if file_type == "md":
            handle_markdown_conversion(selected_file)
        elif file_type == "docx":
            handle_word_conversion(selected_file)

if __name__ == "__main__":
    main()

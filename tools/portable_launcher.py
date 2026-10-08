"""Small Korean desktop entry point using the existing strict conversion pipeline."""

import argparse
import json
import logging
import os
import queue
import sys
import threading
from datetime import datetime
from pathlib import Path

# Permit both direct source execution and the frozen executable.
if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from document_profiles import PROFILES, RenderOptions
from md_to_word import IBReportConverter

AUTO_PROFILE = "자동 (문서의 YAML 설정 사용)"


def application_root() -> Path:
    """Return the portable folder, independent of the working directory."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def convert_document(source: str, output_dir: str, profile: str | None = None) -> Path:
    """Convert one Markdown document without replacing an existing result.

    Args:
        source: Input Markdown path; companion files resolve relative to it.
        output_dir: Destination folder.
        profile: Explicit profile override, or None to preserve YAML choices.

    Returns:
        Actual saved path, including the engine's locked-file fallback name.
    """
    if not source.strip():
        raise ValueError("먼저 Markdown (.md) 파일을 선택하세요.")
    if not output_dir.strip():
        raise ValueError("저장 폴더를 선택하세요.")
    input_path = Path(source).resolve()
    if input_path.suffix.lower() != ".md":
        raise ValueError("입력 파일은 Markdown (.md) 형식이어야 합니다.")
    if profile is not None and profile not in PROFILES:
        raise ValueError(f"알 수 없는 문서 유형: {profile}")
    destination = Path(output_dir).resolve()
    output = destination / f"{input_path.stem}.docx"
    if output.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        output = destination / f"{input_path.stem}_{stamp}.docx"
    converter = IBReportConverter(
        str(input_path), str(output), RenderOptions(profile=profile, strict=True)
    )
    return Path(converter.convert()).resolve()


class QueueLogHandler(logging.Handler):
    """Send engine warnings to the GUI without accessing Tk from a worker."""

    def __init__(self, messages: queue.Queue):
        super().__init__(logging.WARNING)
        self.messages = messages

    def emit(self, record: logging.LogRecord) -> None:
        self.messages.put(("log", self.format(record)))


def launch_gui(self_test: bool = False) -> None:
    """Open the Korean file chooser and conversion window."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText

    root_path = application_root()
    window = tk.Tk()
    if self_test:
        window.withdraw()
    window.title("IB 문서변환기 — Markdown → Word")
    window.geometry("820x590")
    window.minsize(650, 470)
    frame = ttk.Frame(window, padding=20)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    frame.rowconfigure(7, weight=1)
    source = tk.StringVar()
    destination = tk.StringVar(value=str(root_path / "출력"))
    profile = tk.StringVar(value=AUTO_PROFILE)
    status = tk.StringVar(value="예제 .md 파일을 복사해 수정한 뒤 선택하세요.")
    messages: queue.Queue = queue.Queue()
    log_handler = QueueLogHandler(messages)
    logging.getLogger().addHandler(log_handler)
    busy = False

    def choose_source() -> None:
        selected = filedialog.askopenfilename(
            title="Markdown 파일 선택", initialdir=str(root_path / "예제"),
            filetypes=[("Markdown 문서", "*.md")],
        )
        if selected:
            source.set(selected)

    def choose_destination() -> None:
        selected = filedialog.askdirectory(title="저장 폴더 선택", initialdir=destination.get())
        if selected:
            destination.set(selected)

    def open_path(path: Path) -> None:
        try:
            if not path.exists():
                raise FileNotFoundError(f"파일 또는 폴더가 없습니다: {path}")
            if sys.platform == "win32":
                os.startfile(str(path))
            else:
                raise OSError("파일 열기 기능은 Windows에서 지원됩니다.")
        except OSError as exc:
            messagebox.showerror("열기 실패", str(exc), parent=window)

    def show_guide() -> None:
        guide = root_path / "시작하기.html"
        if not guide.exists():
            guide = root_path / "docs" / "distribution" / "시작하기.html"
        open_path(guide)

    def work(values: tuple[str, str, str | None]) -> None:
        try:
            messages.put(("done", str(convert_document(*values))))
        except Exception as exc:
            messages.put(("error", str(exc)))

    def convert() -> None:
        nonlocal busy
        if busy:
            return
        values = (source.get(), destination.get(), None if profile.get() == AUTO_PROFILE else profile.get())
        busy = True
        convert_button.configure(state="disabled")
        status.set("변환 중입니다. 완료될 때까지 기다려 주세요.")
        append_log("변환 시작: " + values[0])
        threading.Thread(target=work, args=(values,), daemon=True).start()

    def append_log(text: str) -> None:
        log.configure(state="normal")
        log.insert("end", text + "\n")
        log.see("end")
        log.configure(state="disabled")

    def poll() -> None:
        nonlocal busy
        while not messages.empty():
            kind, value = messages.get_nowait()
            append_log(value)
            if kind in {"done", "error"}:
                busy = False
                convert_button.configure(state="normal")
                status.set("변환 완료 — 아래 저장 경로를 확인하세요." if kind == "done" else "변환 실패 — 아래 원인을 확인하세요.")
                if kind == "error":
                    messagebox.showerror("변환 실패", value, parent=window)
        window.after(100, poll)

    def close() -> None:
        if busy:
            messagebox.showinfo("변환 중", "변환이 끝난 뒤 창을 닫아 주세요.", parent=window)
            return
        logging.getLogger().removeHandler(log_handler)
        window.destroy()

    ttk.Label(frame, text="Markdown 파일을 Word 문서로 변환", font=("맑은 고딕", 17, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))
    ttk.Label(frame, text="입력 파일").grid(row=1, column=0, sticky="w", padx=(0, 12))
    ttk.Entry(frame, textvariable=source).grid(row=1, column=1, sticky="ew", pady=6)
    ttk.Button(frame, text="파일 선택", command=choose_source).grid(row=1, column=2, padx=(8, 0))
    ttk.Label(frame, text="저장 폴더").grid(row=2, column=0, sticky="w")
    ttk.Entry(frame, textvariable=destination).grid(row=2, column=1, sticky="ew", pady=6)
    ttk.Button(frame, text="폴더 선택", command=choose_destination).grid(row=2, column=2, padx=(8, 0))
    ttk.Label(frame, text="문서 유형").grid(row=3, column=0, sticky="w")
    ttk.Combobox(frame, textvariable=profile, values=[AUTO_PROFILE, *PROFILES], state="readonly").grid(row=3, column=1, sticky="ew", pady=6)
    ttk.Label(frame, text="자동: YAML profile 사용 · 유형을 직접 고르면 YAML보다 우선합니다.\n엄격 모드: 누락·렌더링 오류가 있으면 저장하지 않습니다. 기존 결과는 새 이름으로 보존합니다.", wraplength=720).grid(row=4, column=0, columnspan=3, sticky="w", pady=10)
    buttons = ttk.Frame(frame)
    buttons.grid(row=5, column=0, columnspan=3, sticky="w", pady=8)
    convert_button = ttk.Button(buttons, text="Word로 변환", command=convert)
    convert_button.pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="시작하기 · 텀시트 작성법", command=show_guide).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="출력 폴더 열기", command=lambda: open_path(Path(destination.get()))).pack(side="left")
    ttk.Label(frame, textvariable=status, wraplength=720).grid(row=6, column=0, columnspan=3, sticky="w", pady=8)
    log = ScrolledText(frame, height=12, wrap="word", state="disabled", font=("맑은 고딕", 10))
    log.grid(row=7, column=0, columnspan=3, sticky="nsew")
    window.protocol("WM_DELETE_WINDOW", close)
    window.after(100, poll)
    if self_test:
        window.update_idletasks()
        close()
        return
    window.mainloop()


def main() -> int:
    """Launch the GUI, or perform a repeatable frozen-bundle smoke conversion."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--convert", metavar="MARKDOWN")
    parser.add_argument("--output-dir")
    parser.add_argument("--profile", choices=list(PROFILES))
    parser.add_argument("--result-json", type=Path)
    parser.add_argument("--self-test-ui", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.self_test_ui:
        try:
            launch_gui(self_test=True)
            result = {"ok": True, "tk_widgets": "constructed and destroyed while withdrawn"}
        except Exception as exc:
            result = {"ok": False, "error": str(exc)}
        if args.result_json:
            args.result_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        return 0 if result["ok"] else 1
    if args.convert:
        if not args.output_dir or not args.result_json:
            parser.error("--convert requires --output-dir and --result-json")
        try:
            saved = convert_document(args.convert, args.output_dir, args.profile)
            result = {"ok": True, "saved_path": str(saved)}
        except Exception as exc:
            result = {"ok": False, "error": str(exc)}
        args.result_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        return 0 if result["ok"] else 1
    launch_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Image, chart and LaTeX equation rendering for rendered documents.

Changelog (captions):
    - A figure with an ID gets its caption number before the description.

Changelog (alt text):
    - An inserted author image without a description adds a rendering warning
      (`IBDocumentRenderer.warnings`) that never rejects; the document is unchanged.
"""

import logging
import os
import re
from collections.abc import Sequence
from io import BytesIO
from pathlib import Path, PureWindowsPath

from docx.document import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches

from document_model import CaptionNumber, Chart, TextRun
from ib_render_fonts import FontStyler
from input_safety import (
    DEFAULT_LIMITS,
    AssetOutsideRootError,
    InputLimitError,
    InputLimits,
    check_image_source,
    decode_image,
    local_asset_path,
)
from md_parser import Image
from render_styles import STYLE, RasterFontPolicy

logger = logging.getLogger("ib_renderer")


class ImageRenderer:
    """
    Renders images in Word documents.

    Supports:
        - Base64 embedded images (data:image/... URI)
        - File path images (local files)
        - Fallback to placeholder if image cannot be loaded
    """

    # Maximum image width in inches (fits within typical IB report margins)
    MAX_WIDTH_INCHES: float = 5.5
    # Word's default left/right cell margin when a table does not set one.
    _DEFAULT_CELL_MARGIN_TWIPS = 108
    _URI_RE = re.compile(r"^(?:[a-zA-Z][a-zA-Z0-9+.-]*://|data:)", re.IGNORECASE)

    def __init__(self, doc: DocxDocument):
        self.doc = doc

    @staticmethod
    def local_path(
        path: str, base_dir: Path | None = None, asset_roots: Sequence[str | Path] = ()
    ) -> Path | None:
        """Return the existing local file an image destination names.

        Markdown destinations are URLs, so converters write `a%20b.png`; the
        path is tried as written first (a file may really contain `%`), then
        percent-decoded.

        Args:
            path: Destination as written.
            base_dir: Folder for a relative path; None keeps it cwd-relative.
            asset_roots: Operator-supplied allowed folders.

        Returns:
            The file, or None when neither form names an existing file.
        """
        return local_asset_path(path, base_dir, asset_roots)

    @classmethod
    def image_source(
        cls,
        image: Image,
        base_dir: Path | None = None,
        asset_roots: Sequence[str | Path] = (),
        *,
        limits: InputLimits = DEFAULT_LIMITS,
    ) -> str | BytesIO | None:
        """Return what python-docx inserts for an image: decoded bytes or a file path.

        Args:
            image: Parsed image; Base64 data or a `data:` URI wins over a path.
            base_dir: Folder for a relative path (see `local_path`).
            asset_roots: Operator-supplied allowed folders.
            limits: Immutable caller limits for bytes and header pixel dimensions.

        Returns:
            A stream or path, or None when the file does not exist.
        """
        data = image.base64_data
        if data is None and image.path.startswith("data:"):
            header, _, payload = image.path.partition(",")
            if header.endswith(";base64"):
                data = payload
        if data is not None:
            return decode_image(data, limits)
        if cls._URI_RE.match(image.path) and not Path(image.path).is_absolute():
            return None
        location = cls.local_path(image.path, base_dir, asset_roots)
        if location is not None:
            check_image_source(str(location), limits)
        return str(location) if location is not None else None

    @classmethod
    def fit_table_cell_images(cls, document: DocxDocument) -> None:
        """Shrink pictures wider than the text width of their table cell.

        Run after every table has its final grid: a cell's text width is its
        grid columns (gridSpan included) minus the table's cell margins. Height
        scales with width, so the aspect ratio is kept.

        Args:
            document: Rendered document.
        """
        for table in document.element.body.iter(qn("w:tbl")):
            grid = [int(width) for width in table.xpath("./w:tblGrid/w:gridCol/@w:w")]
            margins = []
            for side in ("left", "right"):
                values = table.xpath(f"./w:tblPr/w:tblCellMar/w:{side}/@w:w")
                margins.append(int(values[0]) if values else cls._DEFAULT_CELL_MARGIN_TWIPS)
            for row in table.xpath("./w:tr"):
                before = row.xpath("./w:trPr/w:gridBefore/@w:val")
                column = int(before[0]) if before else 0
                for cell in row.xpath("./w:tc"):
                    span_values = cell.xpath("./w:tcPr/w:gridSpan/@w:val")
                    span = int(span_values[0]) if span_values else 1
                    limit = (sum(grid[column : column + span]) - sum(margins)) * 635  # twips to EMU
                    column += span
                    if limit <= 0:
                        continue
                    for inline in cell.xpath(".//wp:inline"):
                        extents = inline.xpath("./wp:extent") + inline.xpath(".//a:xfrm/a:ext")
                        width = int(extents[0].get("cx"))
                        if width <= limit:
                            continue
                        height = int(extents[0].get("cy")) * limit // width
                        for extent in extents:
                            extent.set("cx", str(limit))
                            extent.set("cy", str(height))

    def render(self, image: Image, caption_prefix: Sequence[TextRun] = ()) -> bool:
        """
        Render an image to the Word document.

        Args:
            image: Image object with either base64_data or path
            caption_prefix: The caption number of a figure with an ID, written
                before the description in the caption line.

        Attempts to insert actual image; falls back to placeholder on failure.

        Returns:
            Whether an image was inserted successfully.
        """
        import tempfile

        inserted = False
        temp_file_path = None
        message = "Image could not be rendered"

        try:
            # ── Case 1: Base64 embedded image ──────────────────────────────────
            if image.base64_data:
                # Decode Base64 to bytes
                img_bytes = decode_image(
                    image.base64_data, getattr(self.doc.part, "_ib_limits", DEFAULT_LIMITS)
                ).getvalue()

                # Determine file extension from MIME type
                ext = self._mime_to_extension(image.mime_type)

                # Write to temporary file
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False, mode="wb") as f:
                    f.write(img_bytes)
                    temp_file_path = f.name

                # Insert into document
                self._insert_image(temp_file_path, image, caption_prefix, loaded=None)
                inserted = True
                logger.debug("Inserted Base64 image: %s", image.alt_text)

            # ── Case 2: File path image ────────────────────────────────────────
            elif image.path:
                # File-backed parses made the path absolute; others stay cwd-relative.
                source = self.image_source(
                    image,
                    getattr(self.doc.part, "_ib_source_dir", None),
                    getattr(self.doc.part, "_ib_asset_roots", ()),
                    limits=getattr(self.doc.part, "_ib_limits", DEFAULT_LIMITS),
                )
                if source is None:
                    logger.warning(
                        "Image file not found: %s — inserting placeholder",
                        image.path,
                    )
                else:
                    self._insert_image(source, image, caption_prefix, loaded=source)
                    inserted = True
                    logger.debug("Inserted file image: %s", image.path[:80])

        except InputLimitError:
            raise
        except AssetOutsideRootError as e:
            message = str(e)
            logger.warning("%s", message)
        except Exception as e:
            logger.warning(
                "Failed to insert image '%s': %s — inserting placeholder",
                image.alt_text,
                e,
            )

        finally:
            # Clean up temporary file
            if temp_file_path:
                try:
                    os.unlink(temp_file_path)
                except OSError:
                    pass

        # ── Fallback: Placeholder ──────────────────────────────────────────────
        if not inserted:
            self._render_placeholder(image.alt_text)
            errors = getattr(self.doc.part, "_ib_render_errors", None)
            if errors is not None:
                errors.append(message)
        return inserted

    def _insert_image(
        self,
        file_path: str | BytesIO,
        image: Image,
        caption_prefix: Sequence[TextRun] = (),
        *,
        loaded: str | BytesIO | None,
    ):
        """
        Insert image file into document with proper sizing.

        Args:
            file_path: Path to image file, or its bytes
            image: The parsed image; its alt text is the description and caption
            caption_prefix: Caption number runs written before the alt text.
            loaded: The file the image was loaded from; None for embedded data.
        """
        alt_text = image.alt_text
        # Add image with max width constraint
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        inline_shape = run.add_picture(file_path, width=Inches(self.MAX_WIDTH_INCHES))
        self._apply_alt_text(inline_shape, alt_text)
        self.warn_missing_description(self.doc.part, inline_shape, image, loaded)

        # Add caption below image
        if alt_text and caption_prefix:
            from ib_render_text import TextRenderer

            caption_para = self.doc.add_paragraph()
            caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            def style(run) -> None:
                FontStyler.apply_run_style(
                    run, font_size=STYLE.SMALL_SIZE, italic=True, color=STYLE.DARK_GRAY
                )

            for piece in [*caption_prefix, TextRun(text=alt_text)]:
                if isinstance(piece, CaptionNumber):
                    TextRenderer.write_caption_number(caption_para, piece, style)
                else:
                    style(caption_para.add_run(piece.text))
        elif alt_text:
            caption_para = self.doc.add_paragraph()
            caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption_run = caption_para.add_run(alt_text)
            FontStyler.apply_run_style(
                caption_run,
                font_size=STYLE.SMALL_SIZE,
                italic=True,
                color=STYLE.DARK_GRAY,
            )

        # Spacer
        self.doc.add_paragraph()

    @staticmethod
    def _apply_alt_text(inline_shape, alt_text: str):
        """Attach descriptive alt text to a Word inline image when available."""
        if not alt_text:
            return

        doc_pr = inline_shape._inline.docPr
        doc_pr.set("descr", alt_text)
        doc_pr.set("title", alt_text)

    @staticmethod
    def warn_missing_description(
        part: object, inline_shape, image: Image, loaded: str | BytesIO | None
    ) -> None:
        """Warn once for an inserted author image whose description is empty or blank.

        A warning only: it goes to the renderer's `warnings` and the log, never to
        the errors that strict mode rejects, and it writes nothing to the document.
        Pictures the engine makes itself (equations, diagrams, logos) have no
        description the author could write, so their renderers do not call this.

        Args:
            part: Document part carrying the render-scoped warning list.
            inline_shape: The picture just inserted.
            image: The parsed image; its description is checked.
            loaded: The file actually loaded, or None/bytes for embedded data. Only
                its last name is shown, never the destination as written: that can
                hold a percent-encoded folder path.
        """
        if image.alt_text.strip():
            return
        if isinstance(loaded, str):
            # Windows rules split on both separators and a drive, on every system.
            source = f"'{PureWindowsPath(loaded).name}'"
        else:
            source = "embedded image"
        name = inline_shape._inline.docPr.get("name")
        message = (
            f"Image without a description (alt text): {source}, Word name '{name}'. "
            'Write one as ![description](...) or <img alt="...">.'
        )
        logger.warning("%s", message)
        warnings = getattr(part, "_ib_render_warnings", None)
        if warnings is not None:
            warnings.append(message)

    def _render_placeholder(self, alt_text: str):
        """Render a placeholder when image cannot be loaded."""
        p = self.doc.add_paragraph(style=STYLE.STYLE_IB_BODY)
        run = p.add_run(f"[Image: {alt_text}]")
        FontStyler.apply_run_style(run, italic=True, color=STYLE.RED)

    @staticmethod
    def _mime_to_extension(mime_type: str) -> str:
        """Convert MIME type to file extension."""
        mime_map = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/jpg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/webp": ".webp",
            "image/svg+xml": ".svg",
            "image/tiff": ".tiff",
        }
        return mime_map.get(mime_type.lower(), ".png")


class ChartRenderer:
    """Insert a validated chart from memory, respecting the active section width."""

    def __init__(self, doc: DocxDocument) -> None:
        self.doc = doc

    def render(self, chart: Chart) -> None:
        """Render an enabled chart; the orchestrator handles diagnostic fallbacks.

        Args:
            chart: Parsed specification and original fence source.

        Raises:
            ValueError: Invalid chart data or unusable page geometry.
        """
        from chart_renderer import render_chart_png

        if chart.error or chart.spec is None:
            raise ValueError(chart.error or "chart specification is missing")
        section = self.doc.sections[-1]
        page_width = section.page_width
        left_margin, right_margin = section.left_margin, section.right_margin
        if page_width is None or left_margin is None or right_margin is None:
            raise ValueError("chart requires explicit page width and margins")
        width = min(6.0, (page_width - left_margin - right_margin) / 914400)
        if width <= 0:
            raise ValueError("chart requires a positive content width")
        png = render_chart_png(chart.spec, width_inches=width)
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            with BytesIO(png) as buffer:
                shape = paragraph.add_run().add_picture(buffer, width=Inches(width))
            ImageRenderer._apply_alt_text(shape, chart.spec.title or chart.label)
        except Exception:
            paragraph._p.getparent().remove(paragraph._p)
            raise


# ═══════════════════════════════════════════════════════════════════════════════
# LATEX RENDERER (NEW)
# ═══════════════════════════════════════════════════════════════════════════════


class LaTeXRenderer:
    """
    Renders LaTeX expressions to PNG images using matplotlib.

    This class provides static methods to convert LaTeX math expressions
    into image files suitable for embedding in Word documents.

    Dependencies:
        - matplotlib (optional, graceful degradation if unavailable)
    """

    # Flag to track if matplotlib is available
    _matplotlib_available: bool | None = None
    _NON_ASCII_RE = re.compile(r"[^\x00-\x7F]")
    _TEXT_COMMAND_RE = re.compile(r"\\(?:text|mathrm|operatorname)\{([^{}]*)\}")
    _FRAC_RE = re.compile(r"\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}")
    _MULTISPACE_RE = re.compile(r"\s+")
    _UNRESOLVED_COMMAND_RE = re.compile(r"\\[A-Za-z]+")
    _LATEX_SYMBOL_REPLACEMENTS: tuple[tuple[str, str], ...] = (
        (r"\Leftrightarrow", "⇔"),
        (r"\Leftarrow", "⇐"),
        (r"\Rightarrow", "⇒"),
        (r"\leftrightarrow", "↔"),
        (r"\leftarrow", "←"),
        (r"\rightarrow", "→"),
        (r"\approx", "≈"),
        (r"\alpha", "α"),
        (r"\beta", "β"),
        (r"\gamma", "γ"),
        (r"\delta", "δ"),
        (r"\epsilon", "ε"),
        (r"\varepsilon", "ε"),
        (r"\theta", "θ"),
        (r"\lambda", "λ"),
        (r"\mu", "μ"),
        (r"\pi", "π"),
        (r"\rho", "ρ"),
        (r"\sigma", "σ"),
        (r"\tau", "τ"),
        (r"\phi", "φ"),
        (r"\varphi", "φ"),
        (r"\omega", "ω"),
        (r"\Gamma", "Γ"),
        (r"\Delta", "Δ"),
        (r"\Theta", "Θ"),
        (r"\Lambda", "Λ"),
        (r"\Pi", "Π"),
        (r"\Sigma", "Σ"),
        (r"\Phi", "Φ"),
        (r"\Omega", "Ω"),
        (r"\cdot", "·"),
        (r"\times", "×"),
        (r"\sqrt", "√"),
        (r"\infty", "∞"),
        (r"\prod", "∏"),
        (r"\sum", "∑"),
        (r"\int", "∫"),
        (r"\geq", "≥"),
        (r"\leq", "≤"),
        (r"\neq", "≠"),
        (r"\pm", "±"),
        (r"\to", "→"),
        (r"\left", ""),
        (r"\right", ""),
        (r"\%", "%"),
    )
    _LATEX_SYMBOL_MAP = dict(_LATEX_SYMBOL_REPLACEMENTS)
    _LATEX_SYMBOL_RE = re.compile(
        "|".join(
            re.escape(source) + r"(?![A-Za-z])"
            for source, _ in sorted(
                _LATEX_SYMBOL_REPLACEMENTS,
                key=lambda item: len(item[0]),
                reverse=True,
            )
        )
    )

    @classmethod
    def is_available(cls) -> bool:
        """Check if matplotlib is available for LaTeX rendering."""
        if cls._matplotlib_available is None:
            try:
                from importlib.util import find_spec

                cls._matplotlib_available = find_spec("matplotlib") is not None
            except ImportError:
                cls._matplotlib_available = False
            if not cls._matplotlib_available:
                logger.info(
                    "matplotlib not installed — LaTeX will render as text. "
                    "Install with: pip install matplotlib"
                )
        return cls._matplotlib_available

    @classmethod
    def render_to_image(
        cls,
        expression: str,
        fontsize: int = 14,
        dpi: int = 150,
    ) -> str | None:
        """
        Render a LaTeX expression to a PNG image file.

        Args:
            expression: LaTeX math expression (without $ delimiters)
            fontsize: Font size for the equation
            dpi: Resolution of the output image

        Returns:
            Path to the temporary PNG file, or None if rendering fails
        """
        if not cls.is_available():
            return None

        if cls._NON_ASCII_RE.search(expression):
            display_text = cls.to_display_text(expression)
            if cls._UNRESOLVED_COMMAND_RE.search(display_text):
                return None
            image_path = cls._render_plain_text_to_image(display_text, fontsize=fontsize, dpi=dpi)
            if image_path:
                return image_path

        image_path = cls._render_mathtext_to_image(expression, fontsize=fontsize, dpi=dpi)
        if image_path:
            return image_path

        # A failed mathtext parse is a rendering failure. Removing braces and
        # drawing the remaining command as text cannot establish success.
        return None

    @classmethod
    def to_display_text(cls, expression: str) -> str:
        """Convert a LaTeX expression into a readable unicode text fallback."""
        display = expression.strip()

        previous = None
        while previous != display:
            previous = display
            display = cls._TEXT_COMMAND_RE.sub(lambda match: match.group(1), display)
            display = cls._FRAC_RE.sub(
                lambda match: f"({match.group(1)} / {match.group(2)})",
                display,
            )

        display = cls._LATEX_SYMBOL_RE.sub(
            lambda match: cls._LATEX_SYMBOL_MAP[match.group(0)],
            display,
        )

        display = display.replace("{", "").replace("}", "")
        display = cls._MULTISPACE_RE.sub(" ", display)
        return display.strip()

    @classmethod
    def _render_mathtext_to_image(
        cls,
        expression: str,
        fontsize: int,
        dpi: int,
    ) -> str | None:
        """Render mathtext-compatible LaTeX to a PNG image."""
        fig = None
        temp_path = None
        succeeded = False
        try:
            import tempfile

            from matplotlib.backends.backend_agg import FigureCanvasAgg
            from matplotlib.figure import Figure

            fig = Figure(figsize=(0.01, 0.01))
            FigureCanvasAgg(fig)
            ax = fig.subplots()
            fig.patch.set_alpha(0)
            ax.set_axis_off()

            latex_text = f"${expression}$"

            ax.text(
                0.5,
                0.5,
                latex_text,
                fontsize=fontsize,
                ha="center",
                va="center",
                transform=ax.transAxes,
            )

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False, mode="wb") as temp_file:
                temp_path = temp_file.name

            fig.savefig(
                temp_path,
                dpi=dpi,
                bbox_inches="tight",
                pad_inches=0.1,
                transparent=False,
                facecolor="white",
            )
            succeeded = True
            logger.debug("Rendered LaTeX to: %s", temp_path)
            return temp_path

        except Exception as err:
            logger.warning("LaTeX rendering failed: %s", err)
            return None
        finally:
            if fig is not None:
                fig.clear()
            if temp_path is not None and not succeeded:
                try:
                    os.unlink(temp_path)
                except OSError as err:
                    logger.warning("Cannot remove unsuccessful equation image: %s", err)

    @classmethod
    def _render_plain_text_to_image(
        cls,
        display_text: str,
        fontsize: int,
        dpi: int,
    ) -> str | None:
        """Render a readable unicode fallback image for non-ASCII equations."""
        fig = None
        temp_path = None
        succeeded = False
        try:
            import tempfile

            from matplotlib.backends.backend_agg import FigureCanvasAgg
            from matplotlib.figure import Figure
            from matplotlib.font_manager import FontProperties

            fig = Figure(figsize=(0.01, 0.01))
            FigureCanvasAgg(fig)
            fig.patch.set_facecolor("white")
            fig.patch.set_alpha(1)
            font_props = FontProperties(
                family=RasterFontPolicy.resolve(STYLE.KOREAN_FONT, display_text)
            )

            fig.text(
                0.5,
                0.5,
                display_text,
                fontsize=fontsize,
                ha="center",
                va="center",
                fontproperties=font_props,
            )

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False, mode="wb") as temp_file:
                temp_path = temp_file.name

            fig.savefig(
                temp_path,
                dpi=dpi,
                bbox_inches="tight",
                pad_inches=0.12,
                transparent=False,
                facecolor="white",
            )
            succeeded = True
            logger.debug("Rendered unicode equation to: %s", temp_path)
            return temp_path

        except Exception as err:
            logger.warning("Unicode equation rendering failed: %s", err)
            return None
        finally:
            if fig is not None:
                fig.clear()
            if temp_path is not None and not succeeded:
                try:
                    os.unlink(temp_path)
                except OSError as err:
                    logger.warning("Cannot remove unsuccessful equation image: %s", err)

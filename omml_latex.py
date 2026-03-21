"""
OMML (Office Math Markup Language) to LaTeX converter.

Converts Word equation XML (<m:oMath>, <m:oMathPara>) into LaTeX strings.
Adapted from Microsoft's markitdown project (MIT License):
  https://github.com/microsoft/markitdown

Usage:
    from omml_latex import pre_process_docx_math, extract_latex_from_paragraph

    # Option 1: Pre-process entire DOCX (modifies XML in-memory)
    processed_stream = pre_process_docx_math(docx_stream)

    # Option 2: Extract LaTeX from a single paragraph XML element
    latex_parts = extract_latex_from_paragraph(paragraph_element)
"""

import logging
import zipfile
from io import BytesIO
from typing import BinaryIO, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════════════
# OMML NAMESPACE & CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

OMML_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Characters that need escaping in LaTeX
CHARS = ("{", "}", "_", "^", "#", "&", "$", "%", "~")
BLANK = ""
BACKSLASH = "\\"
ALN = "&"
BRK = "\\\\"
FUNC_PLACE = "{fe}"

# ═══════════════════════════════════════════════════════════════════════════════
# LATEX SYMBOL DICTIONARIES
# ═══════════════════════════════════════════════════════════════════════════════

# Accent characters (top & bottom, over/under groups)
CHR: Dict[str, str] = {
    # Top accents
    "\u0300": "\\grave{{{0}}}",
    "\u0301": "\\acute{{{0}}}",
    "\u0302": "\\hat{{{0}}}",
    "\u0303": "\\tilde{{{0}}}",
    "\u0304": "\\bar{{{0}}}",
    "\u0305": "\\overbar{{{0}}}",
    "\u0306": "\\breve{{{0}}}",
    "\u0307": "\\dot{{{0}}}",
    "\u0308": "\\ddot{{{0}}}",
    "\u030c": "\\check{{{0}}}",
    "\u0338": "\\not{{{0}}}",
    "\u20d6": "\\overleftarrow{{{0}}}",
    "\u20d7": "\\vec{{{0}}}",
    "\u20db": "\\dddot{{{0}}}",
    "\u20e1": "\\overleftrightarrow{{{0}}}",
    # Bottom accents
    "\u0330": "\\wideutilde{{{0}}}",
    "\u0331": "\\underbar{{{0}}}",
    # Over | group
    "\u23b4": "\\overbracket{{{0}}}",
    "\u23dc": "\\overparen{{{0}}}",
    "\u23de": "\\overbrace{{{0}}}",
    # Under | group
    "\u23b5": "\\underbracket{{{0}}}",
    "\u23dd": "\\underparen{{{0}}}",
    "\u23df": "\\underbrace{{{0}}}",
}

# Big operators
CHR_BO: Dict[str, str] = {
    "\u220f": "\\prod",
    "\u2210": "\\coprod",
    "\u2211": "\\sum",
    "\u222b": "\\int",
    "\u22c0": "\\bigwedge",
    "\u22c1": "\\bigvee",
    "\u22c2": "\\bigcap",
    "\u22c3": "\\bigcup",
    "\u2a00": "\\bigodot",
    "\u2a01": "\\bigoplus",
    "\u2a02": "\\bigotimes",
}

# Text symbol replacements (Greek, relations, Latin italic)
T: Dict[str, str] = {
    "\u2192": "\\rightarrow ",
    # Greek letters
    "\U0001d6fc": "\\alpha ",
    "\U0001d6fd": "\\beta ",
    "\U0001d6fe": "\\gamma ",
    "\U0001d6ff": "\\delta ",
    "\U0001d700": "\\epsilon ",
    "\U0001d701": "\\zeta ",
    "\U0001d702": "\\eta ",
    "\U0001d703": "\\theta ",
    "\U0001d704": "\\iota ",
    "\U0001d705": "\\kappa ",
    "\U0001d706": "\\lambda ",
    "\U0001d707": "\\mu ",
    "\U0001d708": "\\nu ",
    "\U0001d709": "\\xi ",
    "\U0001d70b": "\\pi ",
    "\U0001d70c": "\\rho ",
    "\U0001d70d": "\\varsigma ",
    "\U0001d70e": "\\sigma ",
    "\U0001d70f": "\\tau ",
    "\U0001d710": "\\upsilon ",
    "\U0001d711": "\\phi ",
    "\U0001d712": "\\chi ",
    "\U0001d713": "\\psi ",
    "\U0001d714": "\\omega ",
    "\U0001d715": "\\partial ",
    "\U0001d716": "\\varepsilon ",
    "\U0001d717": "\\vartheta ",
    "\U0001d719": "\\varphi ",
    "\U0001d71a": "\\varrho ",
    "\U0001d71b": "\\varpi ",
    # Relation symbols
    "\u2190": "\\leftarrow ",
    "\u2191": "\\uparrow ",
    "\u2193": "\\downarrow ",
    "\u2194": "\\leftrightarrow ",
    "\u22ee": "\\vdots ",
    "\u22ef": "\\cdots ",
    "\u22f1": "\\ddots ",
    "\u2260": "\\ne ",
    "\u2264": "\\leq ",
    "\u2265": "\\geq ",
    "\u226a": "\\ll ",
    "\u226b": "\\gg ",
    "\u2208": "\\in ",
    "\u2209": "\\notin ",
    "\u220b": "\\ni ",
    # Ordinary symbols
    "\u221e": "\\infty ",
    # Binary relations
    "\u00b1": "\\pm ",
    "\u2213": "\\mp ",
    # Italic Latin uppercase
    "\U0001d434": "A", "\U0001d435": "B", "\U0001d436": "C", "\U0001d437": "D",
    "\U0001d438": "E", "\U0001d439": "F", "\U0001d43a": "G", "\U0001d43b": "H",
    "\U0001d43c": "I", "\U0001d43d": "J", "\U0001d43e": "K", "\U0001d43f": "L",
    "\U0001d440": "M", "\U0001d441": "N", "\U0001d442": "O", "\U0001d443": "P",
    "\U0001d444": "Q", "\U0001d445": "R", "\U0001d446": "S", "\U0001d447": "T",
    "\U0001d448": "U", "\U0001d449": "V", "\U0001d44a": "W", "\U0001d44b": "X",
    "\U0001d44c": "Y", "\U0001d44d": "Z",
    # Italic Latin lowercase
    "\U0001d44e": "a", "\U0001d44f": "b", "\U0001d450": "c", "\U0001d451": "d",
    "\U0001d452": "e", "\U0001d453": "f", "\U0001d454": "g", "\U0001d456": "i",
    "\U0001d457": "j", "\U0001d458": "k", "\U0001d459": "l", "\U0001d45a": "m",
    "\U0001d45b": "n", "\U0001d45c": "o", "\U0001d45d": "p", "\U0001d45e": "q",
    "\U0001d45f": "r", "\U0001d460": "s", "\U0001d461": "t", "\U0001d462": "u",
    "\U0001d463": "v", "\U0001d464": "w", "\U0001d465": "x", "\U0001d466": "y",
    "\U0001d467": "z",
}

# Function names
FUNC: Dict[str, str] = {
    "sin": "\\sin({fe})", "cos": "\\cos({fe})", "tan": "\\tan({fe})",
    "arcsin": "\\arcsin({fe})", "arccos": "\\arccos({fe})", "arctan": "\\arctan({fe})",
    "sinh": "\\sinh({fe})", "cosh": "\\cosh({fe})", "tanh": "\\tanh({fe})",
    "coth": "\\coth({fe})", "sec": "\\sec({fe})", "csc": "\\csc({fe})",
}

CHR_DEFAULT = {"ACC_VAL": "\\hat{{{0}}}"}
POS = {"top": "\\overline{{{0}}}", "bot": "\\underline{{{0}}}"}
POS_DEFAULT = {"BAR_VAL": "\\overline{{{0}}}"}

SUB = "_{{{0}}}"
SUP = "^{{{0}}}"

F = {
    "bar": "\\frac{{{num}}}{{{den}}}",
    "skw": r"^{{{num}}}/_{{{den}}}",
    "noBar": "\\genfrac{{}}{{}}{{0pt}}{{}}{{{num}}}{{{den}}}",
    "lin": "{{{num}}}/{{{den}}}",
}
F_DEFAULT = "\\frac{{{num}}}{{{den}}}"

D = "\\left{left}{text}\\right{right}"
D_DEFAULT = {"left": "(", "right": ")", "null": "."}

RAD = "\\sqrt[{deg}]{{{text}}}"
RAD_DEFAULT = "\\sqrt{{{text}}}"
ARR = "\\begin{{array}}{{c}}{text}\\end{{array}}"

LIM_FUNC = {"lim": "\\lim_{{{lim}}}", "max": "\\max_{{{lim}}}", "min": "\\min_{{{lim}}}"}
LIM_TO = ("\\rightarrow", "\\to")
LIM_UPP = "\\overset{{{lim}}}{{{text}}}"
M = "\\begin{{matrix}}{text}\\end{{matrix}}"


# ═══════════════════════════════════════════════════════════════════════════════
# OMML → LATEX CONVERTER
# ═══════════════════════════════════════════════════════════════════════════════


def _escape_latex(strs: str) -> str:
    last = None
    new_chr = []
    strs = strs.replace(r"\\", "\\")
    for c in strs:
        if (c in CHARS) and (last != BACKSLASH):
            new_chr.append(BACKSLASH + c)
        else:
            new_chr.append(c)
        last = c
    return BLANK.join(new_chr)


def _get_val(key, default=None, store=None):
    if store is None:
        store = CHR
    if key is not None:
        return key if not store else store.get(key, key)
    return default


class _Pr:
    """Common properties of an OMML element."""

    _val_tags = ("chr", "pos", "begChr", "endChr", "type")

    def __init__(self, elm: ET.Element):
        self._dict: Dict[str, Optional[str]] = {}
        self.text = self._process_children(elm)

    def __str__(self) -> str:
        return self.text

    def __getattr__(self, name: str):
        if name.startswith("_"):
            raise AttributeError(name)
        return self._dict.get(name, None)

    def _process_children(self, elm: ET.Element) -> str:
        parts = []
        for child in elm:
            if OMML_NS not in child.tag:
                continue
            stag = child.tag.replace(OMML_NS, "")
            if stag == "brk":
                self._dict["brk"] = BRK
                parts.append(BRK)
            elif stag in self._val_tags:
                val = child.get(f"{OMML_NS}val")
                self._dict[stag] = val
        return BLANK.join(parts)


class OmmlToLatex:
    """Convert an <m:oMath> XML element to LaTeX."""

    _direct_tags = ("box", "sSub", "sSup", "sSubSup", "num", "den", "deg", "e")

    def __init__(self, element: ET.Element):
        self._latex = self._process_children(element)

    def __str__(self) -> str:
        return self.latex

    @property
    def latex(self) -> str:
        return self._latex

    # ── child processing ──────────────────────────────────────────────────

    def _process_children(self, elm: ET.Element, include=None) -> str:
        parts = []
        for stag, text, _ in self._iter_children(elm, include):
            if text is not None:
                parts.append(str(text))
        return BLANK.join(parts)

    def _process_children_dict(self, elm: ET.Element, include=None) -> Dict:
        result: Dict = {}
        for stag, text, _ in self._iter_children(elm, include):
            result[stag] = text
        return result

    def _iter_children(self, elm: ET.Element, include=None):
        for child in elm:
            if OMML_NS not in child.tag:
                continue
            stag = child.tag.replace(OMML_NS, "")
            if include and stag not in include:
                continue
            handler = self._tag_handlers.get(stag)
            if handler:
                text = handler(self, child)
            elif stag in self._direct_tags:
                text = self._process_children(child)
            elif stag.endswith("Pr"):
                text = _Pr(child)
            else:
                continue
            if text is not None:
                yield (stag, text, child)

    # ── element handlers ──────────────────────────────────────────────────

    def _do_acc(self, elm):
        c = self._process_children_dict(elm)
        latex_s = _get_val(c["accPr"].chr, default=CHR_DEFAULT.get("ACC_VAL"), store=CHR)
        return latex_s.format(c["e"])

    def _do_bar(self, elm):
        c = self._process_children_dict(elm)
        pr = c["barPr"]
        latex_s = _get_val(pr.pos, default=POS_DEFAULT.get("BAR_VAL"), store=POS)
        return pr.text + latex_s.format(c["e"])

    def _do_d(self, elm):
        c = self._process_children_dict(elm)
        pr = c["dPr"]
        null = D_DEFAULT.get("null")
        s_val = _get_val(pr.begChr, default=D_DEFAULT.get("left"), store=T)
        e_val = _get_val(pr.endChr, default=D_DEFAULT.get("right"), store=T)
        return pr.text + D.format(
            left=null if not s_val else _escape_latex(s_val),
            text=c["e"],
            right=null if not e_val else _escape_latex(e_val),
        )

    def _do_sub(self, elm):
        return SUB.format(self._process_children(elm))

    def _do_sup(self, elm):
        return SUP.format(self._process_children(elm))

    def _do_f(self, elm):
        c = self._process_children_dict(elm)
        pr = c["fPr"]
        latex_s = _get_val(pr.type, default=F_DEFAULT, store=F)
        return pr.text + latex_s.format(num=c.get("num"), den=c.get("den"))

    def _do_func(self, elm):
        c = self._process_children_dict(elm)
        func_name = c.get("fName")
        return func_name.replace(FUNC_PLACE, c.get("e", ""))

    def _do_fname(self, elm):
        parts = []
        for stag, text, _ in self._iter_children(elm):
            if stag == "r":
                func = FUNC.get(str(text))
                if func:
                    parts.append(func)
                else:
                    parts.append(str(text))
            else:
                parts.append(str(text))
        result = BLANK.join(parts)
        return result if FUNC_PLACE in result else result + FUNC_PLACE

    def _do_groupchr(self, elm):
        c = self._process_children_dict(elm)
        pr = c["groupChrPr"]
        latex_s = _get_val(pr.chr)
        return pr.text + latex_s.format(c["e"])

    def _do_rad(self, elm):
        c = self._process_children_dict(elm)
        text = c.get("e", "")
        deg = c.get("deg")
        if deg:
            return RAD.format(deg=deg, text=text)
        return RAD_DEFAULT.format(text=text)

    def _do_eqarr(self, elm):
        return ARR.format(
            text=BRK.join(
                t for stag, t, _ in self._iter_children(elm, include=("e",))
            )
        )

    def _do_limlow(self, elm):
        c = self._process_children_dict(elm, include=("e", "lim"))
        latex_s = LIM_FUNC.get(c.get("e", ""))
        if not latex_s:
            return c.get("e", "") + "_{" + c.get("lim", "") + "}"
        return latex_s.format(lim=c.get("lim"))

    def _do_limupp(self, elm):
        c = self._process_children_dict(elm, include=("e", "lim"))
        return LIM_UPP.format(lim=c.get("lim"), text=c.get("e"))

    def _do_lim(self, elm):
        return self._process_children(elm).replace(LIM_TO[0], LIM_TO[1])

    def _do_m(self, elm):
        rows = []
        for stag, text, _ in self._iter_children(elm):
            if stag == "mr":
                rows.append(str(text))
        return M.format(text=BRK.join(rows))

    def _do_mr(self, elm):
        return ALN.join(
            t for stag, t, _ in self._iter_children(elm, include=("e",))
        )

    def _do_nary(self, elm):
        res = []
        bo = ""
        for stag, text, _ in self._iter_children(elm):
            if stag == "naryPr":
                bo = _get_val(text.chr, store=CHR_BO) or ""
            else:
                res.append(str(text))
        return bo + BLANK.join(res)

    def _do_r(self, elm):
        _str = []
        t_elem = elm.find(f"{OMML_NS}t")
        if t_elem is not None and t_elem.text:
            for s in t_elem.text:
                _str.append(T.get(s, s))
        return _escape_latex(BLANK.join(_str))

    _tag_handlers = {
        "acc": _do_acc, "r": _do_r, "bar": _do_bar,
        "sub": _do_sub, "sup": _do_sup, "f": _do_f,
        "func": _do_func, "fName": _do_fname,
        "groupChr": _do_groupchr, "d": _do_d,
        "rad": _do_rad, "eqArr": _do_eqarr,
        "limLow": _do_limlow, "limUpp": _do_limupp,
        "lim": _do_lim, "m": _do_m, "mr": _do_mr,
        "nary": _do_nary,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# PARAGRAPH-LEVEL EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════


def extract_latex_from_paragraph(para_xml: ET.Element) -> List[Tuple[str, bool]]:
    """Extract LaTeX equations from a paragraph XML element.

    Searches for <m:oMathPara> (block equations) and <m:oMath> (inline equations)
    within the given paragraph element.

    Returns:
        List of (latex_string, is_block) tuples.
    """
    results: List[Tuple[str, bool]] = []

    # Block equations: <m:oMathPara> contains one or more <m:oMath>
    for omath_para in para_xml.iter(f"{OMML_NS}oMathPara"):
        for omath in omath_para.iter(f"{OMML_NS}oMath"):
            try:
                latex = OmmlToLatex(omath).latex
                if latex.strip():
                    results.append((latex.strip(), True))
            except Exception as e:
                logger.debug("Failed to convert block OMML to LaTeX: %s", e)
        # Remove so we don't double-process in the inline pass
        parent = _find_parent(para_xml, omath_para)
        if parent is not None:
            parent.remove(omath_para)

    # Inline equations: <m:oMath> directly in paragraph
    for omath in para_xml.iter(f"{OMML_NS}oMath"):
        try:
            latex = OmmlToLatex(omath).latex
            if latex.strip():
                results.append((latex.strip(), False))
        except Exception as e:
            logger.debug("Failed to convert inline OMML to LaTeX: %s", e)

    return results


def _find_parent(root: ET.Element, target: ET.Element) -> Optional[ET.Element]:
    """Find the parent of a target element in the tree."""
    for parent in root.iter():
        for child in parent:
            if child is target:
                return parent
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# DOCX PRE-PROCESSOR (ZIP-level)
# ═══════════════════════════════════════════════════════════════════════════════

# Full namespace declaration for constructing valid XML roots
_MATH_ROOT_TEMPLATE = "".join((
    "<w:document ",
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ',
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    "{0}</w:document>",
))

# Files inside .docx ZIP that may contain equations
_PRE_PROCESS_FILES = [
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
]


def pre_process_docx_math(input_stream: BinaryIO) -> BinaryIO:
    """Pre-process a DOCX stream: convert OMML equations to LaTeX text runs.

    Works by unzipping the DOCX in memory, transforming XML files that may
    contain equations, and re-zipping without writing to disk.

    Returns a new BytesIO stream with the processed DOCX.
    """
    output = BytesIO()

    with zipfile.ZipFile(input_stream, mode="r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, content in files.items():
            if name in _PRE_PROCESS_FILES:
                try:
                    updated = _convert_omml_in_xml(content)
                    zout.writestr(name, updated)
                except Exception:
                    logger.debug("OMML pre-processing failed for %s, using original", name)
                    zout.writestr(name, content)
            else:
                zout.writestr(name, content)

    output.seek(0)
    return output


def _convert_omml_in_xml(content: bytes) -> bytes:
    """Convert all OMML elements in an XML document to LaTeX text runs."""
    # Parse with namespace awareness
    root = ET.fromstring(content)
    _ns = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
           "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    modified = False

    # Process block equations: <m:oMathPara>
    for omath_para in root.iter(f"{OMML_NS}oMathPara"):
        latex_parts = []
        for omath in omath_para.iter(f"{OMML_NS}oMath"):
            try:
                latex = OmmlToLatex(omath).latex.strip()
                if latex:
                    latex_parts.append(latex)
            except Exception:
                continue

        if latex_parts:
            # Create a <w:p><w:r><w:t>$$...$$</w:t></w:r></w:p>
            p_elem = ET.Element(f"{W_NS}p")
            r_elem = ET.SubElement(p_elem, f"{W_NS}r")
            t_elem = ET.SubElement(r_elem, f"{W_NS}t")
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t_elem.text = " ".join(f"$${part}$$" for part in latex_parts)

            parent = _find_parent_et(root, omath_para)
            if parent is not None:
                idx = list(parent).index(omath_para)
                parent.remove(omath_para)
                parent.insert(idx, p_elem)
                modified = True

    # Process inline equations: <m:oMath> (remaining ones)
    for omath in list(root.iter(f"{OMML_NS}oMath")):
        try:
            latex = OmmlToLatex(omath).latex.strip()
        except Exception:
            continue

        if not latex:
            continue

        r_elem = ET.Element(f"{W_NS}r")
        t_elem = ET.SubElement(r_elem, f"{W_NS}t")
        t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_elem.text = f"${latex}$"

        parent = _find_parent_et(root, omath)
        if parent is not None:
            idx = list(parent).index(omath)
            parent.remove(omath)
            parent.insert(idx, r_elem)
            modified = True

    if modified:
        return ET.tostring(root, encoding="unicode").encode("utf-8")
    return content


def _find_parent_et(root: ET.Element, target: ET.Element) -> Optional[ET.Element]:
    """Find parent of target in an ElementTree."""
    for parent in root.iter():
        for child in parent:
            if child is target:
                return parent
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# CONVENIENCE: count equations in a DOCX
# ═══════════════════════════════════════════════════════════════════════════════


def count_equations_in_docx(docx_path: str) -> Dict[str, int]:
    """Count block and inline equations in a DOCX file.

    Useful for diagnostics and round-trip auditing.
    """
    counts = {"block": 0, "inline": 0}

    with open(docx_path, "rb") as f:
        with zipfile.ZipFile(f, "r") as zf:
            for name in _PRE_PROCESS_FILES:
                if name not in zf.namelist():
                    continue
                content = zf.read(name)
                root = ET.fromstring(content)

                for _ in root.iter(f"{OMML_NS}oMathPara"):
                    counts["block"] += 1
                for _ in root.iter(f"{OMML_NS}oMath"):
                    counts["inline"] += 1
                # Inline count includes those inside oMathPara, adjust
                counts["inline"] -= counts["block"]

    return counts

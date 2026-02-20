#!/usr/bin/env python3
"""
IEEE Markdown to DOCX Converter v7

Converts a markdown file to IEEE conference two-column format.
Double-click or run from command line. Prompts for file path if none given.

Specs (from IEEE conference-template-letter.docx, verified against XML):
  Page: US Letter 8.5" x 11"
  Margins: top=54pt (0.75"), bottom=72pt (1"), left/right=44.65pt (~0.62")
  Title: 24pt centered (single-column), after=6pt
  Authors: multi-author support via N-column continuous sections (up to 4/row)
    Single author: 11pt centered name + 10pt italic affiliation lines
    Multiple authors: 9pt per-column blocks with soft line breaks
  Abstract: 9pt bold justified, firstLine=13.60pt, after=10pt
  Keywords: 9pt bold italic justified, firstLine=13.70pt, after=6pt
  Body: 10pt Times New Roman justified, firstLine=14.40pt, after=6pt, line=11.4pt auto
  Heading 1: 10pt small caps centered, before=8pt, after=4pt
  Heading 2: 10pt italic left-aligned, before=6pt, after=3pt
  References: 8pt justified, indent=17.70pt hanging, after=2.5pt, line=9pt exact
  Equations: before=12pt, after=12pt
  Columns: 2-column with 18pt (0.25") gap (body only)
  Author columns: 10.80pt gap (per template)

Dependencies: python-docx (pip install python-docx)
"""

import re
import sys
import os
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, Twips, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml


# ============================================================================
# Constants
# ============================================================================

FONT = "Times New Roman"

# Page dimensions in twips (1 inch = 1440 twips)
# python-docx Inches() returns EMUs; we need raw twips for XML injection.
PAGE_W_TWIPS = 12240       # 8.5"
PAGE_H_TWIPS = 15840       # 11"
MARGIN_TOP_TWIPS = 1080    # 0.75"
MARGIN_BOTTOM_TWIPS = 1440 # 1"
MARGIN_LR_TWIPS = 893      # 44.65pt (template exact)
# Keep EMU versions for python-docx API calls
PAGE_W = Inches(8.5)
PAGE_H = Inches(11)
MARGIN_TOP = Inches(0.75)
MARGIN_BOTTOM = Inches(1.0)
MARGIN_LR = Emu(MARGIN_LR_TWIPS * 914)  # twips -> EMU (1 twip = 914.4 EMU)

# Font sizes
TITLE_PT = 24
AUTHOR_PT = 11
AUTHOR_AFFIL_PT = 9       # Affiliation lines within author block (template: sz=18 = 9pt)
AFFIL_PT = 10              # Legacy single-author affiliation
ABSTRACT_PT = 9
BODY_PT = 10
H1_PT = 10
H2_PT = 10
REF_PT = 8
KEYWORDS_PT = 9

# Author column gap (template: 10.80pt between author columns)
AUTHOR_COL_SPACE_TWIPS = 216  # 10.80pt = 216 twips

# Spacing (in twips; 1pt = 20 twips)
BODY_FIRST_INDENT = 288         # 14.40pt (template: BodyText firstLine)
ABSTRACT_FIRST_INDENT = 272     # 13.60pt (template: Abstract firstLine)
KEYWORDS_FIRST_INDENT = 274     # 13.70pt (template: Keywords firstLine)
COL_SPACE_TWIPS = 360           # 18pt = 0.25" column gap
REF_INDENT_TWIPS = 354          # 17.70pt (template: references hanging indent)


# ============================================================================
# Markdown Parser
# ============================================================================

def parse_markdown(filepath):
    """Parse IEEE-structured markdown into a dict."""
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    result = {
        "title": "",
        "authors": [],       # list of {"name": str, "lines": [str, ...]}
        "abstract": [],
        "keywords": "",
        "sections": [],
        "references": [],
    }

    state = "front"
    in_references = False
    current_section = None
    h1_count = 0
    h2_count = 0
    i = 0

    while i < len(lines):
        line = lines[i].rstrip("\n")
        trimmed = line.strip()
        i += 1

        # References
        if re.match(r"^##\s+References", trimmed, re.IGNORECASE) or trimmed == "REFERENCES":
            in_references = True
            current_section = None
            continue

        if in_references:
            ref_match = re.match(r"^\[(\d+)\]\s*(.*)", trimmed)
            if ref_match:
                ref_text = ref_match.group(2)
                # Collect continuation lines
                while i < len(lines) and lines[i].strip() and not re.match(r"^\[\d+\]", lines[i].strip()):
                    ref_text += " " + lines[i].strip()
                    i += 1
                result["references"].append(ref_text)
            continue

        # Title (H1)
        if line.startswith("# ") and state == "front":
            result["title"] = line[2:].strip()
            state = "post_title"
            continue

        # Author/affiliation after title
        # Each **Name** starts a new author; *italic* lines are affiliation info
        # This block continues until ## Abstract (case-insensitive) is reached.
        # Horizontal rules (---) and blank lines are ignored here.
        if state == "post_title":
            if re.match(r"^---+$", trimmed):
                continue
            if trimmed == "":
                continue
            bold_match = re.match(r"^\*\*(.+?)\*\*\s*$", trimmed)
            if bold_match:
                result["authors"].append({
                    "name": bold_match.group(1).strip(),
                    "lines": [],
                })
                continue
            italic_match = re.match(r"^\*(.+?)\*\s*$", trimmed)
            if italic_match and result["authors"]:
                result["authors"][-1]["lines"].append(
                    italic_match.group(1).strip()
                )
                continue
            # Non-matching lines: fall through to abstract/heading checks below

        # Abstract (case-insensitive)
        if re.match(r"^##\s+Abstract", trimmed, re.IGNORECASE):
            state = "abstract"
            current_section = None
            continue

        if state == "abstract":
            if trimmed == "" or re.match(r"^---+$", trimmed):
                continue
            if re.match(r"^##\s", trimmed):
                state = "body"
                # fall through to heading parse
            else:
                result["abstract"].append(trimmed)
                continue

        # Keywords
        if re.match(r"^##\s+Keywords", trimmed, re.IGNORECASE):
            state = "keywords"
            current_section = None
            continue

        if state == "keywords":
            if trimmed == "" or re.match(r"^---+$", trimmed):
                continue
            if re.match(r"^##\s", trimmed):
                state = "body"
                # fall through
            else:
                result["keywords"] = strip_markdown(trimmed)
                continue

        # Skip horizontal rules in body (visual separators, no output)
        if re.match(r"^---+$", trimmed):
            continue

        # H2 section heading
        if re.match(r"^## ", trimmed):
            state = "body"
            heading = trimmed[3:].strip()
            h1_count += 1
            h2_count = 0
            current_section = {
                "level": 1,
                "heading": heading,
                "number": h1_count,
                "content": [],
            }
            result["sections"].append(current_section)
            continue

        # H3 subsection heading
        if re.match(r"^### ", trimmed):
            heading = trimmed[4:].strip()
            h2_count += 1
            current_section = {
                "level": 2,
                "heading": heading,
                "number": h2_count,
                "content": [],
            }
            result["sections"].append(current_section)
            continue

        # Body text
        if current_section is not None and state == "body":
            if trimmed:
                # Detect display equation: $$...$$
                eq_match = re.match(r"^\$\$(.+?)\$\$$", trimmed)
                if eq_match:
                    current_section["content"].append(("equation", eq_match.group(1)))
                else:
                    current_section["content"].append(trimmed)

    return result


def strip_markdown(text):
    """Remove **bold** and *italic* markers."""
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    return text


def to_roman(n):
    vals = [(1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),
            (100,"C"),(90,"XC"),(50,"L"),(40,"XL"),
            (10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")]
    result = ""
    for v, s in vals:
        while n >= v:
            result += s
            n -= v
    return result


def to_letter(n):
    return chr(64 + n)  # A=1, B=2


# ============================================================================
# Document Builder
# ============================================================================

def make_run(text, size=BODY_PT, bold=False, italic=False, small_caps=False,
             subscript=False, superscript=False, font=FONT):
    """Create a configured run element."""
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:cs"), font)
    rPr.append(rFonts)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(size * 2))  # half-points
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(size * 2))
    rPr.append(szCs)

    if bold:
        rPr.append(OxmlElement("w:b"))
        rPr.append(OxmlElement("w:bCs"))
    if italic:
        rPr.append(OxmlElement("w:i"))
        rPr.append(OxmlElement("w:iCs"))
    if small_caps:
        rPr.append(OxmlElement("w:smallCaps"))
    if subscript:
        vertAlign = OxmlElement("w:vertAlign")
        vertAlign.set(qn("w:val"), "subscript")
        rPr.append(vertAlign)
    if superscript:
        vertAlign = OxmlElement("w:vertAlign")
        vertAlign.set(qn("w:val"), "superscript")
        rPr.append(vertAlign)

    run.append(rPr)

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    return run


NBSP = "\u00A0"  # non-breaking space

# LaTeX command -> Unicode glyph map
LATEX_GLYPHS = {
    # Greek lowercase
    "\\alpha": "\u03B1", "\\beta": "\u03B2", "\\gamma": "\u03B3",
    "\\delta": "\u03B4", "\\epsilon": "\u03B5", "\\varepsilon": "\u03B5",
    "\\zeta": "\u03B6", "\\eta": "\u03B7", "\\theta": "\u03B8",
    "\\iota": "\u03B9", "\\kappa": "\u03BA", "\\lambda": "\u03BB",
    "\\mu": "\u03BC", "\\nu": "\u03BD", "\\xi": "\u03BE",
    "\\pi": "\u03C0", "\\rho": "\u03C1", "\\sigma": "\u03C3",
    "\\tau": "\u03C4", "\\upsilon": "\u03C5", "\\phi": "\u03C6",
    "\\varphi": "\u03C6", "\\chi": "\u03C7", "\\psi": "\u03C8",
    "\\omega": "\u03C9",
    # Greek uppercase
    "\\Gamma": "\u0393", "\\Delta": "\u0394", "\\Theta": "\u0398",
    "\\Lambda": "\u039B", "\\Xi": "\u039E", "\\Pi": "\u03A0",
    "\\Sigma": "\u03A3", "\\Phi": "\u03A6", "\\Psi": "\u03A8",
    "\\Omega": "\u03A9",
    # Operators and relations
    "\\cdot": "\u00B7", "\\times": "\u00D7", "\\div": "\u00F7",
    "\\pm": "\u00B1", "\\mp": "\u2213",
    "\\leq": "\u2264", "\\geq": "\u2265", "\\neq": "\u2260",
    "\\approx": "\u2248", "\\equiv": "\u2261", "\\sim": "\u223C",
    "\\propto": "\u221D",
    "\\in": "\u2208", "\\notin": "\u2209", "\\subset": "\u2282",
    "\\supset": "\u2283", "\\cup": "\u222A", "\\cap": "\u2229",
    "\\emptyset": "\u2205",
    "\\infty": "\u221E",
    "\\partial": "\u2202",
    "\\nabla": "\u2207",
    "\\forall": "\u2200", "\\exists": "\u2203",
    "\\rightarrow": "\u2192", "\\leftarrow": "\u2190",
    "\\Rightarrow": "\u21D2", "\\Leftarrow": "\u21D0",
    "\\leftrightarrow": "\u2194", "\\Leftrightarrow": "\u21D4",
    "\\mapsto": "\u21A6",
    # Integrals and sums
    "\\int": "\u222B", "\\iint": "\u222C", "\\iiint": "\u222D",
    "\\oint": "\u222E",
    "\\sum": "\u2211", "\\prod": "\u220F",
    # Misc
    "\\ldots": "\u2026", "\\dots": "\u2026", "\\cdots": "\u22EF",
    "\\prime": "\u2032",
    "\\neg": "\u00AC", "\\wedge": "\u2227", "\\vee": "\u2228",
    "\\oplus": "\u2295", "\\otimes": "\u2297",
    "\\dagger": "\u2020", "\\ddagger": "\u2021",
    "\\ell": "\u2113",
    "\\hbar": "\u210F",
    "\\quad": "\u2003",       # em-space
    "\\qquad": "\u2003\u2003",
}

# Math function names rendered as upright text (no glyph, just the name)
LATEX_FUNCTIONS = {
    "\\tanh", "\\cosh", "\\sinh", "\\sin", "\\cos", "\\tan",
    "\\log", "\\ln", "\\exp", "\\lim", "\\max", "\\min",
    "\\sup", "\\inf", "\\det", "\\dim", "\\ker", "\\arg",
    "\\deg", "\\gcd", "\\hom", "\\sec", "\\csc", "\\cot",
    "\\arcsin", "\\arccos", "\\arctan",
}


def resolve_latex(text):
    """
    Resolve LaTeX commands in text to Unicode.

    Processing order:
      0. Normalize double backslashes (markdown escaping) to single
      1. $$...$$ display equations (already handled by parser)
      2. $...$ inline math: resolve internals, convert spaces to NBSP
      3. \\text{...} -> plain text content
      4. \\mathbb{X} -> double-struck Unicode
      5. \\frac{a}{b} -> a/b
      6. \\command -> Unicode glyph or function name
    """
    # Step 0: Normalize double backslashes before LaTeX commands
    text = re.sub(r'\\\\([a-zA-Z])', r'\\\1', text)

    # Step 1: Process $...$ inline math spans
    # Replace spaces inside $...$ with NBSP, then strip the $ delimiters
    def resolve_inline_math(m):
        inner = m.group(1)
        # Resolve LaTeX inside the math span first
        inner = _resolve_latex_commands(inner)
        # Replace regular spaces with non-breaking spaces
        inner = inner.replace(" ", NBSP)
        return inner

    # Process $...$ (but not $$...$$, which are already handled)
    # Negative lookbehind/lookahead for $ to avoid matching $$
    text = re.sub(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', resolve_inline_math, text)

    # Also resolve any LaTeX commands outside of $...$ spans
    text = _resolve_latex_commands(text)

    return text


# Double-struck (blackboard bold) Unicode mappings
_MATHBB = {
    "A": "\U0001D538", "B": "\U0001D539", "C": "\u2102",
    "D": "\U0001D53B", "E": "\U0001D53C", "F": "\U0001D53D",
    "G": "\U0001D53E", "H": "\u210D", "I": "\U0001D540",
    "J": "\U0001D541", "K": "\U0001D542", "L": "\U0001D543",
    "M": "\U0001D544", "N": "\u2115", "O": "\U0001D546",
    "P": "\u2119", "Q": "\u211A", "R": "\u211D",
    "S": "\U0001D54A", "T": "\U0001D54B", "U": "\U0001D54C",
    "V": "\U0001D54D", "W": "\U0001D54E", "X": "\U0001D54F",
    "Y": "\U0001D550", "Z": "\u2124",
}


def _resolve_latex_commands(text):
    """Replace LaTeX commands with Unicode glyphs.
    Handles both \\cmd and \\\\cmd (markdown often escapes backslashes)."""
    # Normalize double backslashes to single for LaTeX command matching
    text = re.sub(r'\\\\([a-zA-Z])', r'\\\1', text)

    # \text{...} -> content as-is
    text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)

    # \mathbb{X} -> double-struck character
    def mathbb_replace(m):
        char = m.group(1)
        return _MATHBB.get(char, char)
    text = re.sub(r'\\mathbb\{([A-Z])\}', mathbb_replace, text)

    # \frac{a}{b} -> a/b (handles one level of nested braces)
    text = re.sub(r'\\frac\{((?:[^{}]|\{[^{}]*\})*)\}\{((?:[^{}]|\{[^{}]*\})*)\}',
                  r'(\1)/(\2)', text)

    # Function names: \tanh -> tanh, \cosh -> cosh, etc.
    # Use regex with word boundary to avoid \inf matching inside \infty
    for func in LATEX_FUNCTIONS:
        name = func[1:]  # strip backslash
        text = re.sub(re.escape(func) + r'(?![a-zA-Z])', name, text)

    # Glyph substitutions (longest match first to avoid partial matches)
    for cmd in sorted(LATEX_GLYPHS.keys(), key=len, reverse=True):
        text = text.replace(cmd, LATEX_GLYPHS[cmd])

    return text


def parse_math_text(text, size=BODY_PT, bold=False, italic=False, font=FONT):
    """
    Parse text containing LaTeX-style sub/superscripts into a list of runs.
    Resolves LaTeX commands to Unicode first, then handles sub/superscripts.

    Handles:
      X_{multi}  -> X with 'multi' subscripted
      X_c        -> X with 'c' subscripted (single char)
      X^{multi}  -> X with 'multi' superscripted
      X^c        -> X with 'c' superscripted (single char)
      $...$      -> inline math (LaTeX resolved, spaces become non-breaking)
      \alpha etc -> Unicode glyphs

    Returns a list of OxmlElement runs.
    """
    # First resolve LaTeX commands everywhere in the text
    text = resolve_latex(text)

    runs = []
    base_props = {"size": size, "bold": bold, "italic": italic, "font": font}

    # Pattern: match subscript/superscript notation
    # _{...} or _X (single char) or ^{...} or ^X (single char)
    # Single-char: any character that isn't whitespace, {, }, _, or ^
    pattern = re.compile(r'([_^])\{([^}]*)\}|([_^])([^\s{}_^])')

    pos = 0
    for m in pattern.finditer(text):
        # Add text before this match as normal run
        if m.start() > pos:
            preceding = text[pos:m.start()]
            if preceding:
                runs.append(make_run(preceding, **base_props))

        # Determine sub or super
        marker = m.group(1) or m.group(3)
        content = m.group(2) if m.group(1) else m.group(4)
        is_sub = (marker == "_")

        runs.append(make_run(
            content, **base_props,
            subscript=is_sub, superscript=not is_sub
        ))

        pos = m.end()

    # Remaining text after last match
    if pos < len(text):
        remaining = text[pos:]
        if remaining:
            runs.append(make_run(remaining, **base_props))

    # If no matches at all, return single run
    if not runs:
        runs.append(make_run(text, **base_props))

    return runs


def make_paragraph(runs, align="justify", space_before=0, space_after=80,
                   first_indent=None, left_indent=None, hanging=None,
                   keep_next=False, line_spacing=228, line_rule="auto",
                   tab_stops=None):
    """Create a paragraph element with formatting."""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")

    # OOXML schema requires specific element order in pPr:
    # keepNext, spacing, ind, jc (among others)

    if keep_next:
        pPr.append(OxmlElement("w:keepNext"))

    # Tab stops: list of (position_twips, alignment) tuples
    # alignment: "center", "end" (right), "start" (left)
    if tab_stops:
        tabs = OxmlElement("w:tabs")
        for pos, align_type in tab_stops:
            tab = OxmlElement("w:tab")
            tab.set(qn("w:val"), align_type)
            tab.set(qn("w:pos"), str(pos))
            tabs.append(tab)
        pPr.append(tabs)

    # Spacing
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(space_before))
    spacing.set(qn("w:after"), str(space_after))
    if line_spacing is not None:
        spacing.set(qn("w:line"), str(line_spacing))
        spacing.set(qn("w:lineRule"), line_rule)
    pPr.append(spacing)

    # Indentation
    if first_indent is not None or left_indent is not None or hanging is not None:
        ind = OxmlElement("w:ind")
        if first_indent is not None and hanging is None:
            ind.set(qn("w:firstLine"), str(first_indent))
        if left_indent is not None:
            ind.set(qn("w:left"), str(left_indent))
        if hanging is not None:
            ind.set(qn("w:hanging"), str(hanging))
        pPr.append(ind)

    # Alignment
    jc = OxmlElement("w:jc")
    align_map = {
        "justify": "both",
        "center": "center",
        "left": "start",
        "right": "end",
    }
    jc.set(qn("w:val"), align_map.get(align, "both"))
    pPr.append(jc)

    p.append(pPr)

    if isinstance(runs, list):
        for r in runs:
            p.append(r)
    else:
        p.append(runs)

    return p


def inject_continuous_two_col_section_break(body, insert_before_element):
    """
    Insert a continuous section break with 2-column layout.
    This is the key trick: python-docx can't create continuous section breaks,
    so we build the XML directly and inject it into the paragraph preceding
    the two-column content.
    """
    # We need to add a sectPr to the pPr of the paragraph BEFORE
    # the two-column content starts. This tells Word:
    # "end the current section here (single-column) and start a new
    #  continuous section with 2 columns."

    # Create a separator paragraph with the section properties
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")

    sectPr = OxmlElement("w:sectPr")

    # Page size
    pgSz = OxmlElement("w:pgSz")
    pgSz.set(qn("w:w"), str(PAGE_W_TWIPS))
    pgSz.set(qn("w:h"), str(PAGE_H_TWIPS))
    sectPr.append(pgSz)

    # Margins
    pgMar = OxmlElement("w:pgMar")
    pgMar.set(qn("w:top"), str(MARGIN_TOP_TWIPS))
    pgMar.set(qn("w:right"), str(MARGIN_LR_TWIPS))
    pgMar.set(qn("w:bottom"), str(MARGIN_BOTTOM_TWIPS))
    pgMar.set(qn("w:left"), str(MARGIN_LR_TWIPS))
    pgMar.set(qn("w:header"), "720")
    pgMar.set(qn("w:footer"), "720")
    pgMar.set(qn("w:gutter"), "0")
    sectPr.append(pgMar)

    # Single column for title section (no num attribute = 1 column)
    cols = OxmlElement("w:cols")
    cols.set(qn("w:space"), str(COL_SPACE_TWIPS))
    cols.set(qn("w:num"), "1")
    sectPr.append(cols)

    # Document grid
    docGrid = OxmlElement("w:docGrid")
    docGrid.set(qn("w:linePitch"), "360")
    sectPr.append(docGrid)

    # NOTE: No w:type element here. Default = "nextPage" for the first section,
    # but because the BODY section will have type=continuous, the body content
    # will continue on the same page after the title.

    pPr.append(sectPr)
    p.append(pPr)

    # Insert this paragraph before the target element
    insert_before_element.addprevious(p)


def inject_section_break(body, insert_before_element, num_cols=1,
                         col_space=COL_SPACE_TWIPS, continuous=True):
    """
    Insert a continuous (or next-page) section break before an element.
    Used to transition between different column layouts (e.g., title -> author
    columns -> body columns).
    """
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    sectPr = OxmlElement("w:sectPr")

    if continuous:
        sect_type = OxmlElement("w:type")
        sect_type.set(qn("w:val"), "continuous")
        sectPr.append(sect_type)

    pgSz = OxmlElement("w:pgSz")
    pgSz.set(qn("w:w"), str(PAGE_W_TWIPS))
    pgSz.set(qn("w:h"), str(PAGE_H_TWIPS))
    sectPr.append(pgSz)

    pgMar = OxmlElement("w:pgMar")
    pgMar.set(qn("w:top"), str(MARGIN_TOP_TWIPS))
    pgMar.set(qn("w:right"), str(MARGIN_LR_TWIPS))
    pgMar.set(qn("w:bottom"), str(MARGIN_BOTTOM_TWIPS))
    pgMar.set(qn("w:left"), str(MARGIN_LR_TWIPS))
    pgMar.set(qn("w:header"), "720")
    pgMar.set(qn("w:footer"), "720")
    pgMar.set(qn("w:gutter"), "0")
    sectPr.append(pgMar)

    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), str(num_cols))
    cols.set(qn("w:space"), str(col_space))
    if num_cols > 1:
        cols.set(qn("w:equalWidth"), "1")
    sectPr.append(cols)

    docGrid = OxmlElement("w:docGrid")
    docGrid.set(qn("w:linePitch"), "360")
    sectPr.append(docGrid)

    pPr.append(sectPr)
    p.append(pPr)
    insert_before_element.addprevious(p)


def make_author_paragraph(author, size=AUTHOR_AFFIL_PT):
    """
    Build a single author paragraph with soft line breaks between lines.

    Template structure (per author column):
      Name           (9pt, regular, centered)
      <br/>
      Department     (9pt, italic)
      <br/>
      Organization   (9pt, italic)
      <br/>
      City, Country  (9pt, regular)
      <br/>
      email/ORCID    (9pt, regular)

    All in one paragraph so it stays in a single column cell.
    """
    runs = []

    # Author name
    runs.append(make_run(author["name"], size=size))

    # Affiliation lines separated by soft breaks
    for line_text in author.get("lines", []):
        # Soft line break
        br_run = OxmlElement("w:r")
        br_rPr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(size * 2))
        br_rPr.append(sz)
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(size * 2))
        br_rPr.append(szCs)
        br_run.append(br_rPr)
        br_run.append(OxmlElement("w:br"))
        runs.append(br_run)

        # Affiliation text (italic)
        runs.append(make_run(line_text, size=size, italic=True))

    return make_paragraph(
        runs,
        align="center", space_before=0, space_after=40,
    )


def set_final_section_two_col(doc):
    """
    Set the final (document-level) section properties to 2 columns.
    The document's last sectPr (direct child of w:body) controls the final section.
    We rebuild it cleanly to ensure correct element order per OOXML schema.
    """
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))

    if sectPr is not None:
        body.remove(sectPr)

    # Build fresh sectPr with correct element order:
    # type, pgSz, pgMar, cols, docGrid
    sectPr = OxmlElement("w:sectPr")

    # type=continuous: body section continues on same page as title
    sect_type = OxmlElement("w:type")
    sect_type.set(qn("w:val"), "continuous")
    sectPr.append(sect_type)

    pgSz = OxmlElement("w:pgSz")
    pgSz.set(qn("w:w"), str(PAGE_W_TWIPS))
    pgSz.set(qn("w:h"), str(PAGE_H_TWIPS))
    sectPr.append(pgSz)

    pgMar = OxmlElement("w:pgMar")
    pgMar.set(qn("w:top"), str(MARGIN_TOP_TWIPS))
    pgMar.set(qn("w:right"), str(MARGIN_LR_TWIPS))
    pgMar.set(qn("w:bottom"), str(MARGIN_BOTTOM_TWIPS))
    pgMar.set(qn("w:left"), str(MARGIN_LR_TWIPS))
    pgMar.set(qn("w:header"), "720")
    pgMar.set(qn("w:footer"), "720")
    pgMar.set(qn("w:gutter"), "0")
    sectPr.append(pgMar)

    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), str(COL_SPACE_TWIPS))
    cols.set(qn("w:equalWidth"), "1")
    sectPr.append(cols)

    docGrid = OxmlElement("w:docGrid")
    docGrid.set(qn("w:linePitch"), "360")
    sectPr.append(docGrid)

    # sectPr must be the LAST child of w:body
    body.append(sectPr)


def build_document(parsed):
    """Build the IEEE-formatted DOCX from parsed markdown."""
    doc = Document()

    # Fix zoom percent in settings (python-docx omits the required attribute)
    settings = doc.settings.element
    zoom = settings.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")

    # Configure default style
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(BODY_PT)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Set page size/margins on the default section (will be overridden)
    section = doc.sections[0]
    section.page_width = PAGE_W
    section.page_height = PAGE_H
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM
    section.left_margin = MARGIN_LR
    section.right_margin = MARGIN_LR

    body = doc.element.body

    # Remove the default empty paragraph
    for p in body.findall(qn("w:p")):
        body.remove(p)

    # ---- SINGLE-COLUMN SECTION: Title ----

    # Title
    body.append(make_paragraph(
        make_run(parsed["title"], size=TITLE_PT),
        align="center", space_after=120,
    ))

    # ---- AUTHORS ----
    authors = parsed["authors"]
    num_authors = len(authors)

    if num_authors == 0:
        # No authors parsed; skip author block entirely
        pass

    elif num_authors == 1:
        # Single author: keep simple centered layout (no multi-col needed)
        author = authors[0]
        body.append(make_paragraph(
            make_run(author["name"], size=AUTHOR_PT),
            align="center", space_after=40,
        ))
        for affil_line in author.get("lines", []):
            body.append(make_paragraph(
                make_run(affil_line, size=AFFIL_PT, italic=True),
                align="center", space_after=40,
            ))

    else:
        # Multiple authors: use N-column continuous section
        # IEEE template uses up to 4 columns per row.
        # For N authors, use min(N, 4) columns per row.
        max_cols_per_row = min(num_authors, 4)

        # Process authors in rows of max_cols_per_row
        for row_start in range(0, num_authors, max_cols_per_row):
            row_authors = authors[row_start:row_start + max_cols_per_row]
            ncols = len(row_authors)

            # Build author paragraphs for this row
            author_paras = []
            for author in row_authors:
                author_paras.append(make_author_paragraph(author))

            # Append all author paragraphs to body
            for ap in author_paras:
                body.append(ap)

            # The first author paragraph of this row needs a section break
            # BEFORE it that starts the N-column section.
            # The last author paragraph needs the section properties embedded
            # in its pPr to END the N-column section.

            # Inject section break before the first author para of this row:
            # transitions from previous section to N-col continuous
            inject_section_break(
                body, author_paras[0],
                num_cols=ncols,
                col_space=AUTHOR_COL_SPACE_TWIPS,
                continuous=True,
            )

        # After all author rows, we need to end the last author section
        # and transition back. We'll handle this when we inject the
        # body section break before the abstract.

    # ---- Mark where two-column content begins ----

    # Abstract: inline format per IEEE template
    # "Abstract--" (italic) followed by body text (bold), all in one paragraph
    abstract_text = strip_markdown(" ".join(parsed["abstract"]))
    abstract_runs = [
        make_run("Abstract", size=ABSTRACT_PT, bold=True, italic=True),
        make_run("\u2014", size=ABSTRACT_PT, bold=True),
    ] + parse_math_text(abstract_text, size=ABSTRACT_PT, bold=True)
    abstract_para = make_paragraph(
        abstract_runs,
        align="justify", space_before=360, space_after=200,
        first_indent=ABSTRACT_FIRST_INDENT,
    )
    body.append(abstract_para)

    # Inject continuous section break BEFORE abstract paragraph
    # This ends the author section (single or multi-col) and starts two-column body
    inject_continuous_two_col_section_break(body, abstract_para)

    # Keywords
    if parsed["keywords"]:
        body.append(make_paragraph(
            [
                make_run("Keywords" + chr(0x2014), size=KEYWORDS_PT, bold=True, italic=True),
                make_run(parsed["keywords"], size=KEYWORDS_PT, bold=True, italic=True),
            ],
            align="justify", space_after=120,
            first_indent=KEYWORDS_FIRST_INDENT,
        ))

    # ---- Body sections ----
    equation_counter = 0
    for sec in parsed["sections"]:
        if sec["level"] == 1:
            heading = sec["heading"]
            # Strip any existing numbering prefix (Arabic, Roman, etc.)
            heading = re.sub(r"^[A-Za-z0-9]+\.\s*", "", heading)
            display = to_roman(sec["number"]) + ". " + heading

            # IEEE template: mixed case + small caps style (NOT .upper())
            body.append(make_paragraph(
                make_run(display, size=H1_PT, small_caps=True),
                align="center", space_before=160, space_after=80,
                keep_next=True,
            ))

        elif sec["level"] == 2:
            heading = sec["heading"]
            heading = re.sub(r"^[A-Za-z0-9]+\.\s*", "", heading)
            display = to_letter(sec["number"]) + ". " + heading

            body.append(make_paragraph(
                make_run(display, size=H2_PT, italic=True),
                align="left", space_before=120, space_after=60,
                keep_next=True,
            ))

        # Content paragraphs
        for para_text in sec["content"]:
            # Display equation: $$...$$ -> centered with right-justified number
            if isinstance(para_text, tuple) and para_text[0] == "equation":
                eq_text = strip_markdown(para_text[1])
                # Resolve LaTeX and replace spaces with non-breaking spaces
                eq_text = resolve_latex(eq_text)
                eq_text = eq_text.replace(" ", NBSP)
                equation_counter += 1

                # Tab to center, equation, tab to right, (number)
                tab1 = OxmlElement("w:r")
                tab1.append(OxmlElement("w:tab"))

                tab2 = OxmlElement("w:r")
                tab2.append(OxmlElement("w:tab"))

                # Column width = 5040 twips; center=2520, right=5040
                body.append(make_paragraph(
                    [tab1]
                    + parse_math_text(eq_text, size=BODY_PT)
                    + [tab2, make_run("(" + str(equation_counter) + ")", size=BODY_PT)],
                    align="left", space_before=240, space_after=240,
                    tab_stops=[(2520, "center"), (5040, "end")],
                ))
                continue

            # Blockquote
            if isinstance(para_text, str) and para_text.startswith("> "):
                quote = strip_markdown(para_text[2:])
                body.append(make_paragraph(
                    parse_math_text(quote, size=BODY_PT, italic=True),
                    align="justify", space_after=120,
                    left_indent=int(Inches(0.2)),
                ))
                continue

            # Numbered list with bold label: "1. **Label** rest"
            list_match = re.match(r"^(\d+)\.\s+\*\*(.+?)\*\*\s*(.*)", para_text)
            if list_match:
                num, label, rest = list_match.groups()
                rest = strip_markdown(rest)
                body.append(make_paragraph(
                    [
                        make_run(num + ". ", size=BODY_PT),
                        make_run(label + " ", size=BODY_PT, bold=True),
                        make_run(rest, size=BODY_PT),
                    ],
                    align="justify", space_after=120,
                    left_indent=int(Inches(0.25)),
                    hanging=int(Inches(0.25)),
                ))
                continue

            # Bold label paragraph: "**Label:** rest"
            bold_match = re.match(r"^\*\*(.+?)\*\*\s*(.*)", para_text)
            if bold_match and bold_match.group(2):
                label, rest = bold_match.groups()
                rest = strip_markdown(rest)
                body.append(make_paragraph(
                    [make_run(label + " ", size=BODY_PT, bold=True)]
                    + parse_math_text(rest, size=BODY_PT),
                    align="justify", space_after=120,
                    first_indent=BODY_FIRST_INDENT,
                ))
                continue

            # Bullet list item
            # Bullet hangs left; tab pushes text to indent position
            if re.match(r"^[-*]\s", para_text):
                item = strip_markdown(re.sub(r"^[-*]\s+", "", para_text))

                # Build bullet run with tab
                bullet_run = OxmlElement("w:r")
                bRPr = OxmlElement("w:rPr")
                brf = OxmlElement("w:rFonts")
                brf.set(qn("w:ascii"), FONT)
                brf.set(qn("w:hAnsi"), FONT)
                brf.set(qn("w:cs"), FONT)
                bRPr.append(brf)
                bsz = OxmlElement("w:sz")
                bsz.set(qn("w:val"), str(BODY_PT * 2))
                bRPr.append(bsz)
                bszCs = OxmlElement("w:szCs")
                bszCs.set(qn("w:val"), str(BODY_PT * 2))
                bRPr.append(bszCs)
                bullet_run.append(bRPr)
                bt = OxmlElement("w:t")
                bt.text = "\u2022"
                bullet_run.append(bt)

                # Tab run
                tab_run = OxmlElement("w:r")
                tab_run.append(OxmlElement("w:tab"))

                body.append(make_paragraph(
                    [bullet_run, tab_run]
                    + parse_math_text(item, size=BODY_PT),
                    align="justify", space_after=40,
                    left_indent=576,    # 28.8pt - text block position
                    hanging=288,        # 14.4pt - bullet hangs into this
                    line_spacing=228,
                ))
                continue

            # Regular paragraph
            cleaned = strip_markdown(para_text)
            if not cleaned.strip():
                continue
            body.append(make_paragraph(
                parse_math_text(cleaned, size=BODY_PT),
                align="justify", space_after=120,
                first_indent=BODY_FIRST_INDENT,
                line_spacing=228,  # ~11.4pt
            ))

    # ---- References ----
    body.append(make_paragraph(
        make_run("References", size=H1_PT, small_caps=True),
        align="center", space_before=200, space_after=80,
    ))

    for i, ref_text in enumerate(parsed["references"]):
        ref_text = strip_markdown(ref_text)
        body.append(make_paragraph(
            [
                make_run("[" + str(i + 1) + "]" + NBSP, size=REF_PT),
                make_run(ref_text, size=REF_PT),
            ],
            align="justify", space_after=50,
            left_indent=REF_INDENT_TWIPS,
            hanging=REF_INDENT_TWIPS,
            line_spacing=180, line_rule="exact",  # 9pt exact
        ))

    # Set the final document section to two-column
    set_final_section_two_col(doc)

    return doc


# ============================================================================
# Main
# ============================================================================

def main():
    # Get input file path
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    else:
        input_path = input("Enter path to markdown file: ").strip().strip('"').strip("'")

    input_path = Path(input_path).expanduser().resolve()

    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    if not input_path.suffix.lower() == ".md":
        print(f"Warning: Expected .md file, got {input_path.suffix}")

    # Output path: same directory, same stem, _IEEE.docx
    output_path = input_path.with_name(input_path.stem + "_IEEE.docx")

    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")
    print()

    # Parse
    print("Parsing markdown...")
    parsed = parse_markdown(str(input_path))
    print(f"  Title:      {parsed['title'][:60]}...")
    print(f"  Authors:    {len(parsed['authors'])}")
    for a in parsed["authors"]:
        print(f"              {a['name']} ({len(a['lines'])} affil lines)")
    print(f"  Abstract:   {len(parsed['abstract'])} paragraph(s)")
    print(f"  Sections:   {len(parsed['sections'])}")
    print(f"  References: {len(parsed['references'])}")
    print()

    # Build
    print("Building IEEE document...")
    doc = build_document(parsed)

    # Save
    doc.save(str(output_path))
    print(f"Saved: {output_path}")

    # Pause if double-clicked (no args)
    if len(sys.argv) <= 1:
        input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except ImportError as e:
        print(f"Missing dependency: {e}")
        print()
        print("Install with:  pip install python-docx")
        print()
        input("Press Enter to exit...")
    except Exception as e:
        print(f"Error: {e}")
        print()
        input("Press Enter to exit...")
        raise

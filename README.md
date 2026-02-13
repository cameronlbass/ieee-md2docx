# ieee-md2docx

A Python script that converts Markdown files to IEEE conference-format DOCX documents. Write your paper in Markdown, run the converter, get a two-column IEEE-compliant `.docx` ready for submission.

## What It Does

- Produces a two-column IEEE conference paper from a single `.md` file
- Formats title, multi-author blocks, abstract, keywords, body sections, equations, and references
- Auto-numbers sections (Roman numerals for H1, letters for H2)
- Auto-numbers display equations with flush-right numbering
- Resolves LaTeX commands (`\alpha`, `\sum`, `\frac{a}{b}`, etc.) to Unicode
- Handles subscripts, superscripts, inline math, and display equations
- All spacing, indents, and font sizes verified against the official IEEE conference template XML

## Quick Start

**Requirements:** Python 3.7+ and `python-docx`

```
pip install python-docx
```

**Usage:**

```
python ieee_md2docx.py paper.md
```

Output: `paper_IEEE.docx` in the same directory.

Or double-click the script and enter the file path when prompted.

## Markdown Format

See [FORMAT_GUIDE.md](FORMAT_GUIDE.md) for the complete reference. Here's the minimal structure:

```markdown
# Paper Title

**Author Name**
*Department, University*
*City, Country*
*email@example.edu*

## Abstract
Your abstract text here.

## Keywords
keyword one, keyword two, keyword three

## Introduction
Body text with inline math $O(n \log n)$ and citations [1].

### Subsection
More text. Display equations:

$$E = mc^2$$

## Conclusion
Concluding remarks.

## References
[1] A. Author, "Title," Journal, vol. 1, pp. 1-10, 2024.
```

### Multiple Authors

```markdown
**Alice Smith**
*Dept. of CS, University A*
*City, Country*

**Bob Jones**
*Dept. of EE, University B*
*City, Country*
```

Authors are automatically arranged in equal-width columns (up to 4 per row).

## What's Supported

| Feature | Syntax |
|---------|--------|
| Sections | `## Heading` (auto-numbered I, II, III...) |
| Subsections | `### Heading` (auto-numbered A, B, C...) |
| Bold labels | `**Label:** rest of text` |
| Bullet lists | `- item` or `* item` |
| Numbered lists | `1. **Label** description` |
| Blockquotes | `> quoted text` |
| Inline math | `$expression$` (non-breaking spaces) |
| Display equations | `$$expression$$` (centered, auto-numbered) |
| Subscripts | `X_i` or `X_{multi}` |
| Superscripts | `X^2` or `X^{n+1}` |
| Greek letters | `\alpha`, `\beta`, `\Gamma`, etc. |
| Math operators | `\sum`, `\int`, `\leq`, `\infty`, etc. |
| Fractions | `\frac{a}{b}` renders as (a)/(b) |
| Blackboard bold | `\mathbb{R}` renders as double-struck R |
| References | `[N] text` with multi-line continuation |

## What's Not Supported

Images, tables, footnotes, HTML, heading levels 4+, multi-line display equations, nested lists, and link syntax. Add these manually in Word after conversion.

## Specs

All values verified against the official `conference-template-letter.docx` XML:

| Element | Size | Spacing |
|---------|------|---------|
| Title | 24pt centered | after 6pt |
| Author (single) | 11pt name, 10pt affiliation | centered |
| Author (multi) | 9pt per column | 10.80pt column gap |
| Abstract | 9pt bold | 13.60pt first indent, after 10pt |
| Keywords | 9pt bold italic | 13.70pt first indent, after 6pt |
| Body | 10pt Times New Roman | 14.40pt first indent, after 6pt, 11.4pt line |
| Heading 1 | 10pt small caps | centered, before 8pt, after 4pt |
| Heading 2 | 10pt italic | left-aligned, before 6pt, after 3pt |
| References | 8pt | 17.70pt hanging indent, 9pt exact line, after 2.5pt |
| Equations | 10pt | centered, before/after 12pt |
| Page | US Letter | margins: top 0.75", bottom 1", sides 44.65pt |
| Body columns | 2 | 18pt (0.25") gap |

## License

MIT

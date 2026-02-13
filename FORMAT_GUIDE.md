# How to Format Your Markdown Document for Translation to IEEE-Compliant DOCX

This guide describes the markdown conventions recognized by `ieee_markdown_converter_v5.py`. Follow this structure and your `.md` file will produce a two-column IEEE conference paper in `.docx` format.

---

## Document Structure

The converter expects sections in this order:

```
# Title
**Author Name**
*Affiliation line*

## Abstract
(paragraph text)

## Keywords
(comma-separated terms)

## First Section Heading
(body content)

### Subsection Heading
(body content)

## Second Section Heading
...

## References
[1] Reference text...
[2] Reference text...
```

Each element is described below.

---

## Front Matter

### Title

Use a single `#` heading as the first non-empty line. This becomes the 24pt centered title spanning the full page width.

```markdown
# Format as Pedagogy: Understanding the Epistemic Values Encoded in IEEE Paper Structure
```

### Author and Affiliation

After the title, provide one or more author blocks. Each author starts with a bold name on its own line, followed by italic affiliation lines.

**Single author:**

```markdown
**Cameron L. Bass**
*Independent Scholar*
```

**Multiple authors:**

```markdown
**Alice M. Smith**
*Department of Computer Science*
*University of Example, Springfield, USA*
*alice@example.edu*

**Bob T. Jones**
*Department of Electrical Engineering*
*Institute of Technology, Metro City, UK*
*bob.jones@insttech.ac.uk*

**Carol R. Chen**
*Applied Mathematics Division*
*National Research Lab, Taipei, Taiwan*
*chen.carol@nrl.tw*
```

Each `**Name**` line starts a new author. Every `*italic line*` following it is an affiliation line for that author. Blank lines between author blocks are fine.

**Layout behavior:**

- **1 author:** Centered name at 11pt, affiliation lines at 10pt italic. No column layout.
- **2-4 authors:** Each author gets an equal-width column in a single row, matching the IEEE template's multi-column continuous section with 10.80pt column gaps.
- **5+ authors:** Authors wrap into rows of 4 columns. A paper with 5 authors produces a row of 4 and a row of 1; 7 authors produces rows of 4 and 3.

Within each column, the author name appears at 9pt regular and affiliation lines at 9pt italic, separated by soft line breaks — exactly as the IEEE template structures it.

**Affiliation line conventions (from IEEE template):**

```markdown
**Given Name Surname**
*dept. name of organization*
*name of organization*
*City, Country*
*email address or ORCID*
```

You can use as many or as few affiliation lines as needed. The converter does not enforce a specific number.

### Abstract

Use `## Abstract` as the heading. The heading itself is consumed (not printed); instead, the converter produces the IEEE inline format: *Abstract*--- followed by the abstract body in 9pt bold.

```markdown
## Abstract
The IEEE paper format encodes specific epistemic values. This paper analyzes
IEEE conventions as expressions of disciplinary knowledge-making practices
rather than arbitrary typographical rules.
```

Multiple paragraphs under `## Abstract` are joined into a single block. Blank lines between them are ignored.

### Keywords

Use `## Keywords` as the heading. Provide a single line of comma-separated terms.

```markdown
## Keywords
citation systems, disciplinary conventions, epistemic values, IEEE format
```

Renders as *Keywords*--- followed by the terms, all in 9pt bold italic.

---

## Body Sections

### Section Headings (H1 equivalent)

Use `##` for top-level body sections. The converter auto-numbers them with Roman numerals and renders in 10pt small caps, centered.

```markdown
## Introduction
```

Produces: **I. Introduction** (small caps, centered)

```markdown
## Writing in the Disciplines: Why Conventions Matter
```

Produces: **II. Writing in the Disciplines: Why Conventions Matter**

If your heading already contains a decimal number prefix (e.g., from a prior outline), the converter strips it and substitutes the Roman numeral:

```markdown
## 3. Results and Analysis
```

Produces: **III. Results and Analysis**

### Subsection Headings (H2 equivalent)

Use `###` for subsections. The converter auto-numbers them with capital letters and renders in 10pt italic, left-aligned. The letter counter resets with each new `##` section.

```markdown
### Selecting a Template
```

Produces: *A. Selecting a Template*

```markdown
### Maintaining the Integrity of the Specifications
```

Produces: *B. Maintaining the Integrity of the Specifications*

As with `##`, decimal prefixes like `2.1` are stripped and replaced.

### Heading Levels Beyond H2

The converter currently handles `##` (H1) and `###` (H2) only. Markdown `####` and deeper headings are not recognized and will be treated as body text. If you need sub-subsection structure, use bold labels (described below).

---

## Body Content

All body content must appear after a `##` or `###` heading. Text that appears before any heading (other than front matter) is silently dropped.

### Regular Paragraphs

Plain text lines become justified body paragraphs at 10pt with a 14.4pt first-line indent.

```markdown
The template is used to format your paper and style the text. All margins,
column widths, line spaces, and text fonts are prescribed.
```

Blank lines between paragraphs are fine; the converter ignores them. Consecutive non-blank lines under the same heading are treated as separate paragraphs (one per line). If you want a single paragraph to span multiple source lines, put it all on one line.

### Bold Labels

A line beginning with `**Label:** rest of text` produces a paragraph with the label in bold followed by regular-weight text, preserving the first-line indent.

```markdown
**Critical convention:** Methods sections use past tense and passive voice.
```

Produces: **Critical convention:** Methods sections use past tense and passive voice.

### Numbered Lists with Bold Labels

Lines matching `N. **Label** rest` produce a hanging-indent list item with the number and bold label.

```markdown
1. **Explain problems, not rules.** Rather than "Use passive voice," explain why.
2. **Provide comparative examples.** Show how content appears across formats.
3. **Discuss departures.** Examine why papers depart from IMRAD.
```

### Bullet Lists

Lines beginning with `- ` or `* ` (hyphen-space or asterisk-space) produce bullet list items with a hanging indent.

```markdown
- Biology papers: predominantly IMRAD
- Engineering papers: less than 50% follow strict IMRAD
- Computational/theoretical papers: most depart from IMRAD
```

**Note:** The bullet marker `*` must be followed by a space. A line like `*italic text*` is not a bullet; it is italic markup.

### Blockquotes

Lines beginning with `> ` produce an indented italic paragraph.

```markdown
> The scientific paper is a fraud in the sense that it misrepresents the
> processes of thought that accompanied the work.
```

Only single-line blockquotes are supported. Each `> ` line becomes its own blockquote paragraph.

---

## Inline Formatting

### Bold and Italic

Standard markdown:

```markdown
This has **bold text** and *italic text* and even ***bold italic***.
```

Bold and italic markers are stripped in most contexts (body text, references, keywords). They are recognized for structural purposes in the front matter (author = bold, affiliation = italic) and for bold-label detection in body paragraphs.

### Inline Code

Backtick spans are stripped:

```markdown
Use `passive voice` in methods sections.
```

Renders as: Use passive voice in methods sections. (No monospace; the backticks are simply removed.)

---

## Mathematics

### How LaTeX Resolution Works

The converter resolves LaTeX commands (e.g., `\alpha`, `\sum`, `\frac{a}{b}`) to Unicode in **all** body text, not only inside dollar delimiters. Subscript `_` and superscript `^` notation is also processed everywhere. So a bare line like:

```markdown
The variable \alpha_i converges to \infty.
```

will produce the correct Greek letter, subscript, and infinity symbol without any delimiters.

However, dollar delimiters serve two purposes that bare LaTeX does not:

- **`$...$`** (inline): Spaces inside become non-breaking spaces, preventing the expression from line-breaking mid-formula.
- **`$$...$$`** (display): The expression is centered in the column with an auto-incrementing equation number flush right.

**Recommendation:** Use `$...$` around any inline expression you want kept together on one line. Use `$$...$$` for any equation that should be displayed and numbered. Bare LaTeX works for isolated symbols but offers no layout protection.

### Inline Math

Wrap expressions in single `$` delimiters. Spaces inside become non-breaking spaces and LaTeX commands resolve to Unicode glyphs.

```markdown
The complexity is $O(n \log n)$ in the average case.
```

Renders with the math expression kept together (no line-break inside).

### Display Equations

Wrap a single-line expression in `$$`. The equation is centered in the column with an auto-incrementing number flush right.

```markdown
$$E = mc^2$$
```

Produces a centered equation with `(1)` at the right margin.

```markdown
$$\sum_{i=1}^{n} x_i = X$$
```

Produces the summation with subscript/superscript and `(2)` at right.

**Constraint:** The entire display equation must be on one line. Multi-line `$$` blocks are not supported.

### Subscripts and Superscripts

Use `_` for subscript and `^` for superscript, with braces for multi-character spans:

| Markdown | Result |
|----------|--------|
| `X_i` | X with subscript i |
| `X_{ij}` | X with subscript ij |
| `X^2` | X with superscript 2 |
| `X^{n+1}` | X with superscript n+1 |
| `X_i^2` | X with subscript i and superscript 2 |

These work both inside and outside `$...$` spans.

### Supported LaTeX Commands

The converter maps LaTeX commands to Unicode. These work in both inline `$...$` and display `$$...$$` contexts, and also in bare body text.

**Greek letters:**
`\alpha` `\beta` `\gamma` `\delta` `\epsilon` `\zeta` `\eta` `\theta`
`\iota` `\kappa` `\lambda` `\mu` `\nu` `\xi` `\pi` `\rho` `\sigma`
`\tau` `\upsilon` `\phi` `\chi` `\psi` `\omega`
and uppercase: `\Gamma` `\Delta` `\Theta` `\Lambda` `\Xi` `\Pi` `\Sigma` `\Phi` `\Psi` `\Omega`

**Operators:**
`\cdot` `\times` `\div` `\pm` `\mp` `\leq` `\geq` `\neq` `\approx`
`\equiv` `\sim` `\propto` `\in` `\notin` `\subset` `\supset` `\cup`
`\cap` `\emptyset` `\infty` `\partial` `\nabla` `\forall` `\exists`

**Arrows:**
`\rightarrow` `\leftarrow` `\Rightarrow` `\Leftarrow`
`\leftrightarrow` `\Leftrightarrow` `\mapsto`

**Large operators:**
`\int` `\iint` `\iiint` `\oint` `\sum` `\prod`

**Functions** (rendered upright, not italic):
`\sin` `\cos` `\tan` `\log` `\ln` `\exp` `\lim` `\max` `\min`
`\sup` `\inf` `\det` `\dim` `\ker` `\arg` `\tanh` `\cosh` `\sinh`
`\sec` `\csc` `\cot` `\arcsin` `\arccos` `\arctan`

**Fractions:**
`\frac{a}{b}` renders as `(a)/(b)`. One level of nesting is supported.

**Blackboard bold:**
`\mathbb{R}` renders as the double-struck Unicode character (works for A-Z).

**Spacing:**
`\quad` and `\qquad` produce em-spaces. `\ldots` / `\dots` produce ellipsis.

**Misc:**
`\ell` `\hbar` `\neg` `\wedge` `\vee` `\oplus` `\otimes` `\dagger` `\ddagger` `\prime`

---

## References

### Section Header

Use either form:

```markdown
## References
```

or simply:

```markdown
REFERENCES
```

### Reference Entries

Each reference begins with `[N]` where N is a sequential integer, followed by the reference text. The converter handles multi-line references: continuation lines (non-blank, not starting with `[N]`) are joined to the previous entry.

```markdown
## References
[1] G. Eason, B. Noble, and I. N. Sneddon, "On certain integrals of
    Lipschitz-Hankel type," Phil. Trans. Roy. Soc. London, vol. A247,
    pp. 529-551, Apr. 1955.
[2] J. Clerk Maxwell, A Treatise on Electricity and Magnetism, 3rd ed.,
    vol. 2. Oxford: Clarendon, 1892, pp. 68-73.
[3] M. Young, The Technical Writer's Handbook. Mill Valley, CA:
    University Science, 1989.
```

References render at 8pt with a 17.70pt hanging indent and 9pt exact line spacing. The bracket number hangs left; the reference text aligns at the indent.

A blank line between references is fine (it terminates the multi-line join for the previous entry, and the next `[N]` starts fresh).

---

## Horizontal Rules

Lines consisting only of three or more hyphens (`---`) are silently skipped. You can use them as visual separators in your source without affecting output.

---

## What the Converter Does Not Handle

The following are **not** supported and will either be ignored or produce unexpected output:

- **Images / figures:** No `![alt](url)` processing. No figure insertion.
- **Tables:** No markdown table syntax. Tables must be added manually in Word after conversion.
- **Footnotes:** No `[^1]` footnote syntax.
- **HTML tags:** Raw HTML in markdown is not processed.
- **Heading levels 4+:** `####` and deeper are treated as plain body text.
- **Multi-line display equations:** `$$` must open and close on the same line.
- **Nested lists:** Only single-level bullets and numbered lists.
- **Link syntax:** `[text](url)` is not resolved; the raw markdown passes through to the output as plain text.
- **Multi-author layout beyond 8:** The converter handles up to ~8 authors in 2 rows of 4 columns. Larger author lists will work but may produce many narrow rows. For 9+ authors, consider editing the DOCX output.

---

## Complete Minimal Example

```markdown
# My Paper Title

**Jane A. Smith**
*Department of Computer Science*
*University of Example, City, Country*
*jane.smith@example.edu*

**John B. Doe**
*Department of Mathematics*
*Institute of Technology, City, Country*
*jdoe@insttech.edu*

## Abstract
This paper presents a novel approach to solving the widget problem.
We demonstrate a 40% improvement over existing methods using our
proposed framework.

## Keywords
widget optimization, novel framework, performance analysis

## Introduction
The widget problem has been studied extensively [1]. Prior approaches
suffer from $O(n^2)$ complexity in the worst case.

### Background
Smith et al. [2] introduced the first polynomial-time solution.

### Our Contribution
We propose a method achieving $O(n \log n)$ time with the recurrence:

$$T(n) = 2T(n/2) + O(n)$$

## Methods
The algorithm was implemented in Python 3.11. All experiments were
conducted on a standard workstation with 32 GB RAM.

**Complexity analysis:** The dominant term is the merge step, bounded
by $\sum_{i=1}^{k} n_i \leq n$.

## Results
- Baseline: 142ms average runtime
- Our method: 87ms average runtime
- Improvement: 38.7% reduction in wall-clock time

## Conclusion
We presented an efficient algorithm for the widget problem, achieving
sub-quadratic performance. Future work includes extending to the
multi-widget variant.

## References
[1] A. Author, "Title of first paper," Journal Name, vol. 1, no. 2,
    pp. 10-20, Jan. 2020.
[2] B. Author and C. Author, "Title of second paper," in Proc. Conf.
    Name, City, Country, 2021, pp. 100-110.
```

---

## Running the Converter

```
python ieee_md2docx.py my_paper.md
```

Or double-click the script and enter the path when prompted. Output is saved alongside the input as `my_paper_IEEE.docx`.

Requires `python-docx`:
```
pip install python-docx
```

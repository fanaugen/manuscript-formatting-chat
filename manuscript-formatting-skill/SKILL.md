---
name: russian-memoir-manuscript-formatting
description: Normalize a Russian memoir manuscript assembled from ODT chapter files while preserving every word, spelling choice, and paragraph boundary. Use for cautious print-manuscript preparation, not copy-editing or rewriting.
---

# Russian Memoir Manuscript Formatting

Prepare a self-publishing manuscript from source `.odt` files without editing the author's text. Preserve words, spelling, punctuation, and paragraph division. Do not infer corrections from general publishing practice.

## Safety boundary

- Work on copies; leave every source `.odt` unchanged.
- Do not reword, spell-check, modernize typography inside quotations, join/split paragraphs, or normalize hyphens/dashes globally.
- Treat a manual line break as content unless it is demonstrably only a layout artifact. Flag uncertain cases instead of changing them.
- Maintain a machine-readable change log containing source file, paragraph index, rule, old value, and new value for every textual normalization.
- Before delivery, compare the normalized text with a source-text extraction. Differences must be limited to pre-approved mechanical rules and recorded in the log.

## Source inventory and order

1. Establish order from the folder names (parts) and the numeric file prefixes. A numbered chapter file follows `{part number}-{two-digit chapter number}`; sort by those numbers, not by the overview file.
2. Use `0-00 Схема книги.odt` only for the Roman-numeral titles of parts. Treat it as editorial metadata, not manuscript body text, unless asked to include it.
3. Treat `0-01 Обложка.odt` and `0-02 Эпиграфы.odt` as front matter. Preserve their separate visual roles.
4. Usually treat one numbered `.odt` as one chapter. Record exceptions rather than guessing.
5. Preserve rich-text emphasis, superscripts, footnotes, and images unless a specific approved style maps them to a manuscript style.

## Classification before normalization

Classify each paragraph before assigning a style. Never use leading spaces, blank paragraphs, font size, or all caps as the only signal.

- **Part title:** a major division of the book, beginning on a right-hand (odd-numbered) page.
- **Chapter title:** normally the first short line in a chapter file; often all caps. It begins on a new page.
- **Epigraph or pull quote:** a visually distinct quotation, possibly followed by a source line.
- **Body paragraph:** prose; fully justified, first-line indent, no extra vertical space between paragraphs.
- **List/enumeration:** a true sequence or set; keep it distinct from a paragraph beginning with a dash as dialogue or prose.
- **Topic separator:** a paragraph consisting only of `***`, or a `***` prefix that the author uses to introduce a new untitled topic. Render the standalone separator centered with vertical space above and below; do not delete it.
- **Uncertain:** retain source structure and place it in a review report.

For a chapter whose first short line is not all caps, classify it as a title only when its role is supported by its file name, position, and surrounding layout. A pre-title quotation must not be mistaken for a chapter title.

## Approved mechanical normalization rules

Apply a rule only after it is approved for this edition and log each affected instance.

1. Replace formatting-only runs of leading spaces in body paragraphs with the body style's first-line indent. Remove trailing layout spaces.
2. Remove empty paragraphs used only for vertical spacing, then express intentional space through paragraph styles. Do not remove blank paragraphs that clearly separate a distinct text unit until reviewed.
3. Convert formatting-only multi-space runs inside titles, quotes, sources, and body text to ordinary single spaces. Do not collapse spaces when they form deliberate visual poetry, tables, or other content.
4. Normalize full dates with accidental internal spaces to `dd.mm.yyyy` when the year is present. Do not add ` г.` unless separately approved.
5. Convert authorial three-dot ellipses (`...`) to the horizontal ellipsis (`…`, U+2026), including in source quotations unless they are code, literal technical notation, or another documented exception.
6. Do not blanket-convert `-` to `–` or `—`. Classify uses first: dialogue dash, parenthetical dash, hyphenated word, numerical range, minus sign, and authorial spelling can require different treatment.
7. Preserve paragraph division. A line break may be removed only when review confirms it was introduced solely to force a page or visual line position.

## Baseline manuscript styles

Use named paragraph styles rather than direct formatting. The following is a starting system; confirm page size, margins, and final typeface before production.

- **Body:** Georgia 11 pt; justified; first-line indent about 2 em (approximately 0.8–0.9 cm at 11 pt); zero space before/after; line spacing selected for the final page size so the page is readable without excessive word spacing.
- **Chapter title:** centered; no first-line indent; begins on a new page; space after controlled by the style.
- **Part title:** centered; no first-line indent; begins on the next odd page by using a section break, not blank pages made from returns.
- **Opening epigraph:** before a chapter title, in italics, right-aligned in an approximately half-width measure; the source is a separate right-aligned non-italic paragraph.
- **Inline/end quotation:** regular (not italic) text, indented on both left and right, with a dedicated quotation style and controlled space before/after.
- **Topic separator:** centered `***`; no indent; controlled space before/after.
- **Dash list:** paragraphs that begin with a dash as list markers use a hanging-indent list style; numbered entries use a true ordered-list style. Neither gets extra vertical separation from adjacent body paragraphs.
- **Bold-led glossary/topic entries:** preserve as ordinary paragraphs with their leading bold concept; do not convert them into list items.

Avoid forced justified lines with very few words. Use the word processor's justification controls (including sensible limits on character spacing) and inspect the rendered pages for conspicuous gaps. Never insert manual spaces to repair a line.

## Production workflow

1. Make a read-only inventory: files, source order, paragraph count, whitespace patterns, style usage, manual line breaks, separators, and suspicious dates.
2. Agree the editorial decisions listed in **Decisions required**.
3. Extract and concatenate chapters in numeric file order into one master `.docx`, using `0-00` only to label parts, preserving inline rich text and assigning semantic styles.
4. Apply the approved mechanical normalization rules and write the change log plus an exceptions report.
5. Add page geometry, headers/footers, page numbers, part/chapter title rules, and a generated table of contents after all chapter titles use consistent heading styles.
6. Render the `.docx` to page images and inspect every page. Verify title rectos, chapter starts, paragraphs, quotes, lists, separators, headers, footers, and page numbering.
7. Deliver the `.docx` and supporting reports. Apple Pages can import and edit `.docx`; do not promise a reliable native `.pages` generator.

## Decisions required

Confirm these before a final normalization pass:

1. Trim size, binding/bleed requirements, and final margins.
2. Header text for left/right pages, placement/format of page numbers, and whether the front matter counts in the pagination.

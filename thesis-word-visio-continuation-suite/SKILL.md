---
name: thesis-word-visio-continuation-suite
description: Use for Chinese thesis or graduation-paper Word DOCX delivery that involves tables, figures, pagination, captions, table-of-contents, or embedded Visio/OLE original diagrams. Trigger whenever a thesis DOCX is edited, especially when content movement may affect table page breaks, when the user mentions 续表, 表格跨页, 图表排号, 目录错误, Visio原图, 可编辑图, Word论文修订, or preserving editable embedded diagrams.
---

# Thesis Word Visio Continuation Suite

Use this skill for thesis DOCX work where layout matters. It coordinates Word delivery, editable Visio/OLE preservation, figure/table numbering, table-of-contents checks, PDF visual verification, and continuation-table handling.

## Core Rule

Whenever a change may affect pagination, table location, table height, captions, headings, figures, or surrounding text, run a continuation-table review before final delivery.

Changes that require the review:

- Adding, deleting, rewriting, or moving paragraphs.
- Replacing screenshots, diagrams, or embedded Visio objects.
- Renumbering figures or tables.
- Updating fields, cross-references, or the table of contents.
- Changing margins, fonts, line spacing, table width, row height, captions, or figure size.
- Exporting through Word after OOXML edits, because Word may repaginate.

## Continuation Table Policy

When the same logical table crosses a page boundary, add a continuation title at the top of the continued part.

Do not create a new table number for a continued part. The continuation keeps the original table number.

Ask the user which continuation-title style to use when it is not already known. The likely styles are:

- `表5 技术可行性分析表（续）`
- `续表5 技术可行性分析表`
- `续表5`

If the user has shown or confirmed a style in the current project, reuse it consistently without asking again. For this user's current thesis workflow, prefer `表X 标题（续）` unless the user says otherwise.

Only add a continuation title when the table actually crosses pages. If two independent tables are on the same page, each keeps its own caption and number. If a table was manually split but both parts remain on the same page, re-evaluate the split point and remove or avoid unnecessary continuation titles.

## Required Workflow

1. Work on a copy of the source DOCX unless the user explicitly asks to overwrite.
2. Preserve user hand edits. Do not revert unrelated content.
3. For Visio/OLE diagrams, use `word-visio-ole-editor`:
   - Inspect editable Visio objects before and after risky figure work.
   - Preserve embedded Visio/OLE objects.
   - Do not flatten Visio originals into screenshots unless the user explicitly allows it.
4. Make the requested text, figure, table, caption, or structure edits.
5. Update fields and TOC using Word COM when the document has a table of contents, captions, or references.
6. Export PDF.
7. Run continuation-table review on the exported PDF and DOCX.
8. Fix any table that crosses pages without a proper continuation title.
9. Re-export PDF and repeat until continuation tables, captions, and TOC are correct.
10. Verify final deliverables:
    - Figure captions are sequential.
    - Table captions are sequential, but continuation captions reuse the same table number.
    - TOC contains headings only, not figure/table captions or embedded diagrams.
    - Editable Visio/OLE count is preserved relative to the chosen baseline.
    - PDF spot-check pages around modified areas look correct.

## Continuation Review Method

Preferred automated pass:

```powershell
python "$env:USERPROFILE\.codex\skills\thesis-word-visio-continuation-suite\scripts\audit_continuation_tables.py" `
  --docx ".\paper.docx" `
  --pdf ".\paper.pdf" `
  --out ".\continuation_audit.json"
```

Treat the audit script as a detector, not an absolute judge. Always visually inspect the PDF pages it flags.

When automation is inconclusive, use Word COM to inspect each table's page span:

- `table.Range.Information(wdActiveEndPageNumber)` for the end page.
- `table.Range.Information(wdActiveEndAdjustedPageNumber)` when page numbering restarts matter.
- Compare the page of the caption immediately before the table with the first and last page of the table range.

## Continuation Fix Guidelines

For an actual cross-page table:

- Keep the original caption before the first part.
- Add the continuation caption at the top of the first continued page or continued table part.
- Repeat the header row if the table format expects it.
- Preserve the table's original style and borders.
- Do not increment the table number for the continuation.
- Update references only to the logical table number.

For a false continuation:

- If both parts are on the same page, remove the continuation caption or merge/adjust the split if that is safer.
- If two independent tables are adjacent, keep separate numbers and titles.

## Visio/OLE Integration

Before and after figure work, use:

```powershell
& "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\scripts\Invoke-WordVisioOleEditor.ps1" `
  -Action inspect `
  -InputPath ".\paper.docx"
```

If editing a Visio original:

- Activate and edit the embedded object with the `word-visio-ole-editor` workflow.
- Keep it double-click editable in Word.
- Export PDF only after saving the updated OLE preview back into Word.

## Final Response

Report the final DOCX/PDF paths and verification evidence:

- Continuation-table review completed.
- Continuation title style used.
- Figure/table numbering result.
- Visio/OLE count, if relevant.
- Any remaining manual check needed.

---
name: word-visio-ole-editor
description: Edit Microsoft Visio drawings embedded as OLE objects inside Word .docx/.doc files on Windows. Use when Codex must inspect, activate, modify, preserve, or verify Word-embedded Visio original figures, including requests in Chinese such as 修改Word里的Visio原图, 直接改Word中的Visio图, 保留可双击编辑的Visio对象, 不要截图替换, or edit embedded OLE diagrams. Requires local Microsoft Word; editing Visio internals requires local Microsoft Visio.
---

# Word Visio OLE Editor

Use this skill for Word documents that contain embedded Visio/OLE drawings and the user wants the original editable object modified in place. The goal is to preserve the double-click-editable Visio object inside Word, not flatten it into PNG/JPG/EMF unless the user explicitly permits that fallback.

For Chinese thesis delivery that also needs figure/table renumbering, TOC checks, and cross-page continuation-table handling, use the bundled `thesis-word-visio-continuation-suite` skill in this repository.

## Operating Rules

- Work on a copy by default; never overwrite the source unless the user explicitly asks.
- Use Microsoft Word COM to open `.docx/.doc`, enumerate `InlineShapes` and `Shapes`, and identify OLE objects whose `ProgID` contains `Visio`.
- Use Microsoft Visio COM only after confirming Visio is installed and the object can be activated.
- Preserve the embedded OLE object, page size, connectors, and shape identities where practical.
- If the Word figure is a picture, EMF/WMF, screenshot, or PDF image rather than a Visio OLE object, state that it is not an editable Visio original and offer to redraw or replace it.
- Close Word and release COM objects in `finally` blocks. Avoid leaving `WINWORD.EXE` or `VISIO.EXE` processes running.

## Quick Commands

Inspect embedded Visio/OLE objects:

```powershell
& "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\scripts\Invoke-WordVisioOleEditor.ps1" -Action inspect -InputPath ".\paper.docx"
```

Replace text inside one embedded Visio object:

```powershell
& "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\scripts\Invoke-WordVisioOleEditor.ps1" `
  -Action replace-text `
  -InputPath ".\paper.docx" `
  -OutputPath ".\paper-visio-edited.docx" `
  -ObjectIndex 1 `
  -FindText "旧文字" `
  -ReplaceText "新文字"
```

Run a task-specific edit script against the activated embedded Visio document:

```powershell
& "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\scripts\Invoke-WordVisioOleEditor.ps1" `
  -Action run-script `
  -InputPath ".\paper.docx" `
  -OutputPath ".\paper-visio-edited.docx" `
  -ObjectIndex 1 `
  -EditScriptPath ".\edit-visio.ps1"
```

The custom edit script receives `$VisioApp`, `$VisioDocument`, `$TargetObject`, and `$WordDocument` in scope.

## Workflow

1. Inspect the Word file with `-Action inspect`.
2. Report whether editable Visio OLE objects were found, including their index, location, and `ProgID`.
3. If the user asked for simple text changes, use `-Action replace-text`.
4. For layout, connector, shape, or style changes, write a small PowerShell edit script and run `-Action run-script`.
5. Save to a new `.docx`, reopen or export a PDF if layout matters, and verify the target is still an embedded Visio/OLE object.

## Custom Edit Script Pattern

Use this shape traversal pattern inside the `-EditScriptPath` file:

```powershell
function Visit-VisioShape {
    param($Shape)
    if ($Shape.Text -eq "Old") { $Shape.Text = "New" }
    for ($i = 1; $i -le $Shape.Shapes.Count; $i++) {
        Visit-VisioShape -Shape $Shape.Shapes.Item($i)
    }
}

for ($p = 1; $p -le $VisioDocument.Pages.Count; $p++) {
    $page = $VisioDocument.Pages.Item($p)
    for ($s = 1; $s -le $page.Shapes.Count; $s++) {
        Visit-VisioShape -Shape $page.Shapes.Item($s)
    }
}
```

## Verification

After editing:

- Run `inspect` on the output file and confirm the same target object still reports a Visio `ProgID`.
- Open or export the Word document to PDF when visual layout matters.
- If activation fails, check that Microsoft Visio is installed and that the object is not merely a static preview image.

## Reference

Load `references/com-patterns.md` when writing non-trivial Visio COM edits, diagnosing activation failures, or distinguishing OLE objects from static pictures.

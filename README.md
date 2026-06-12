# Word Visio OLE Editor

Codex skills for Windows Word thesis/document workflows that need editable Visio preservation.

This repository contains two related skills:

- `word-visio-ole-editor`: edit Microsoft Visio drawings embedded as OLE objects inside Word `.docx` and `.doc` files.
- `thesis-word-visio-continuation-suite`: coordinate Chinese thesis DOCX delivery with editable Visio preservation, figure/table renumbering, TOC checks, PDF verification, and mandatory cross-page continuation-table review.

## Typical Requests

- 修改 Word 里的 Visio 原图
- 直接改 Word 中可双击编辑的 Visio 图
- 保留内嵌 Visio/OLE 对象，不要截图替换
- 论文图表重新排号
- 表格跨页时制作续表标题
- 修改论文内容后重新检查目录、图号、表号和续表

## Requirements

- Windows
- Microsoft Word desktop app
- Microsoft Visio desktop app for editing Visio internals
- PowerShell
- Python for the continuation-table audit helper
- Codex skills support

If the Word figure is only a PNG/JPG/screenshot/EMF/PDF image, it is not an editable Visio original. In that case the workflow should explain the limitation and offer redraw or image replacement.

## Install

Clone or copy this repository into your Codex skills directory:

```powershell
git clone https://github.com/xiaoyezi7788/word-visio-ole-editor "$env:USERPROFILE\.codex\skills\word-visio-ole-editor"
```

To use the thesis continuation suite as a separate discoverable skill, copy or symlink its folder:

```powershell
Copy-Item -Recurse "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\thesis-word-visio-continuation-suite" `
  "$env:USERPROFILE\.codex\skills\thesis-word-visio-continuation-suite"
```

Restart Codex after installation.

## Example Commands

Inspect a Word document for embedded Visio/OLE objects:

```powershell
& "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\scripts\Invoke-WordVisioOleEditor.ps1" -Action inspect -InputPath ".\paper.docx"
```

Replace text inside the first embedded Visio object:

```powershell
& "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\scripts\Invoke-WordVisioOleEditor.ps1" `
  -Action replace-text `
  -InputPath ".\paper.docx" `
  -OutputPath ".\paper-visio-edited.docx" `
  -ObjectIndex 1 `
  -FindText "旧文字" `
  -ReplaceText "新文字"
```

Audit continuation-table candidates after exporting DOCX to PDF:

```powershell
python "$env:USERPROFILE\.codex\skills\thesis-word-visio-continuation-suite\scripts\audit_continuation_tables.py" `
  --docx ".\paper.docx" `
  --pdf ".\paper.pdf" `
  --out ".\continuation_audit.json"
```

## Continuation Table Rule

When the same logical table crosses a page boundary, add a continuation title at the top of the continued part and do not assign a new table number. The continuation style should be confirmed with the user when unknown, for example:

- `表5 技术可行性分析表（续）`
- `续表5 技术可行性分析表`
- `续表5`

Whenever thesis content changes may affect pagination, rerun the continuation-table review before final delivery.

## License

MIT

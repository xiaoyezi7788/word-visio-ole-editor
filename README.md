# Word Visio OLE Editor

A Codex skill for editing Microsoft Visio drawings embedded as OLE objects inside Word `.docx` and `.doc` files on Windows.

The skill is designed for requests such as:

- 修改 Word 里的 Visio 原图
- 直接改 Word 中可双击编辑的 Visio 图
- 保留内嵌 Visio/OLE 对象，不要截图替换

## What It Does

- Inspects Word documents for embedded Visio/OLE objects.
- Activates an embedded Visio object through Word COM.
- Replaces text inside the embedded Visio drawing.
- Runs a custom PowerShell edit script against the activated Visio document.
- Saves the edited object back into the Word document while preserving the editable OLE object.

## Requirements

- Windows
- Microsoft Word desktop app
- Microsoft Visio desktop app for editing Visio internals
- PowerShell
- Codex skills support

If the Word figure is only a PNG/JPG/screenshot/EMF/PDF image, it is not an editable Visio original. In that case the skill should explain the limitation and offer redraw or image replacement workflows.

## Install

Clone or copy this repository into your Codex skills directory:

```powershell
git clone https://github.com/xiaoyezi7788/word-visio-ole-editor "$env:USERPROFILE\.codex\skills\word-visio-ole-editor"
```

Restart Codex after installation.

## Example Commands

Inspect a Word document:

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

Use a custom edit script:

```powershell
& "$env:USERPROFILE\.codex\skills\word-visio-ole-editor\scripts\Invoke-WordVisioOleEditor.ps1" `
  -Action run-script `
  -InputPath ".\paper.docx" `
  -OutputPath ".\paper-visio-edited.docx" `
  -ObjectIndex 1 `
  -EditScriptPath ".\edit-visio.ps1"
```

## License

MIT

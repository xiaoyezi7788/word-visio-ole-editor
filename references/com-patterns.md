# COM Patterns For Word-Embedded Visio

## Identify Targets

Word can store embedded objects as either `InlineShapes` or floating `Shapes`.

- `InlineShape.Type = 1` usually means embedded OLE object.
- `Shape.Type = 7` usually means embedded OLE object.
- `OLEFormat.ProgID` is the most useful discriminator. Treat values containing `Visio` as candidate editable Visio objects.

Common Visio ProgIDs include `Visio.Drawing`, `Visio.Drawing.11`, `Visio.Drawing.15`, and `Visio.Document`.

## Static Images Are Not Editable Originals

If the object is a picture, screenshot, EMF/WMF, or PDF-rendered image, Word COM may expose it as a picture shape rather than an OLE object. It cannot be directly edited as Visio. The correct response is to explain the limitation and offer a redraw or replacement workflow.

## Activation

Activate the OLE object from Word before trying to access Visio:

```powershell
$ole = $target.Object.OLEFormat
$ole.Activate()
```

If activation opens Visio, use `GetActiveObject("Visio.Application")` first. Only create a new `Visio.Application` when the task needs standalone Visio and activation did not expose an active instance.

## Saving Back To Word

For embedded objects, save the active Visio document after editing, then save the Word document. Avoid exporting the Visio page and replacing the Word object unless the user asked for image replacement.

## Shape Traversal

Visio groups contain nested shapes. Recursively visit `Shape.Shapes` so text in grouped diagrams is not missed. Guard COM property access with `try/catch` because some shapes do not expose text or sub-shapes cleanly.

## Common Failure Cases

- Microsoft Visio is not installed.
- The Word object is a static preview image rather than an OLE object.
- Protected View or a modal Office dialog blocks automation.
- The OLE object is linked to a missing external file.
- Word and Visio are running in a different desktop/session than Codex.

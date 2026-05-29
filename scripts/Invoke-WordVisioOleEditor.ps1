param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("inspect", "replace-text", "run-script")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$OutputPath,
    [int]$ObjectIndex = 1,
    [string]$FindText,
    [string]$ReplaceText,
    [string]$EditScriptPath,
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    $executionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-ComProperty {
    param($Object, [string]$Name)
    try { return $Object.$Name } catch { return $null }
}

function Get-WordOleTargets {
    param($Document)

    $targets = New-Object System.Collections.Generic.List[object]
    $ordinal = 0

    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        $ole = Get-ComProperty -Object $shape -Name "OLEFormat"
        $progId = if ($ole) { Get-ComProperty -Object $ole -Name "ProgID" } else { $null }
        if ($progId -and $progId -match "Visio") {
            $ordinal++
            $targets.Add([pscustomobject]@{
                Index = $ordinal
                Kind = "InlineShape"
                CollectionIndex = $i
                ProgID = [string]$progId
                Type = [string](Get-ComProperty -Object $shape -Name "Type")
                Object = $shape
                OLEFormat = $ole
            })
        }
    }

    for ($i = 1; $i -le $Document.Shapes.Count; $i++) {
        $shape = $Document.Shapes.Item($i)
        $ole = Get-ComProperty -Object $shape -Name "OLEFormat"
        $progId = if ($ole) { Get-ComProperty -Object $ole -Name "ProgID" } else { $null }
        if ($progId -and $progId -match "Visio") {
            $ordinal++
            $targets.Add([pscustomobject]@{
                Index = $ordinal
                Kind = "Shape"
                CollectionIndex = $i
                ProgID = [string]$progId
                Type = [string](Get-ComProperty -Object $shape -Name "Type")
                Object = $shape
                OLEFormat = $ole
            })
        }
    }

    return $targets
}

function Get-VisioApplication {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application")
    }
    catch {
        try {
            return New-Object -ComObject Visio.Application
        }
        catch {
            throw "Microsoft Visio COM automation is unavailable. Install Microsoft Visio or verify it can launch in this Windows user session."
        }
    }
}

function Activate-VisioTarget {
    param($Target)

    try {
        $Target.OLEFormat.Activate()
    }
    catch {
        try {
            $Target.OLEFormat.DoVerb()
        }
        catch {
            throw "Failed to activate Word embedded Visio OLE object index $($Target.Index): $($_.Exception.Message)"
        }
    }

    Start-Sleep -Milliseconds 800
    $visio = Get-VisioApplication
    $visio.Visible = [bool]$Visible
    $doc = $visio.ActiveDocument
    if ($null -eq $doc) {
        throw "Visio was available, but no active Visio document appeared after activating the embedded object."
    }

    [pscustomobject]@{
        VisioApp = $visio
        VisioDocument = $doc
    }
}

function Visit-VisioShapes {
    param($Shapes, [scriptblock]$Visitor)

    for ($i = 1; $i -le $Shapes.Count; $i++) {
        $shape = $Shapes.Item($i)
        & $Visitor $shape

        try {
            if ($shape.Shapes.Count -gt 0) {
                Visit-VisioShapes -Shapes $shape.Shapes -Visitor $Visitor
            }
        }
        catch {
            # Not all Visio shapes expose nested Shapes reliably.
        }
    }
}

function Replace-VisioText {
    param($VisioDocument, [string]$Find, [string]$Replacement)

    if ([string]::IsNullOrEmpty($Find)) {
        throw "-FindText is required for replace-text."
    }
    if ($null -eq $Replacement) {
        $Replacement = ""
    }

    $state = [pscustomobject]@{ Count = 0 }
    for ($p = 1; $p -le $VisioDocument.Pages.Count; $p++) {
        $page = $VisioDocument.Pages.Item($p)
        Visit-VisioShapes -Shapes $page.Shapes -Visitor {
            param($shape)
            try {
                $text = [string]$shape.Text
                if ($text.Contains($Find)) {
                    $shape.Text = $text.Replace($Find, $Replacement)
                    $state.Count++
                }
            }
            catch {
                # Shape has no editable text or denies access.
            }
        }
    }

    return $state.Count
}

$inputFull = Resolve-FullPath -Path $InputPath
if (-not (Test-Path -LiteralPath $inputFull)) {
    throw "InputPath does not exist: $inputFull"
}

if ($Action -ne "inspect") {
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        throw "-OutputPath is required for $Action."
    }
    $outputFull = Resolve-FullPath -Path $OutputPath
    Copy-Item -LiteralPath $inputFull -Destination $outputFull -Force
    $wordPath = $outputFull
}
else {
    $wordPath = $inputFull
}

$word = $null
$doc = $null
$visio = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0

    $readOnly = ($Action -eq "inspect")
    $doc = $word.Documents.Open($wordPath, $false, $readOnly)
    $targets = @(Get-WordOleTargets -Document $doc)

    if ($Action -eq "inspect") {
        $result = [pscustomobject]@{
            InputPath = $wordPath
            VisioOleCount = $targets.Count
            Targets = @($targets | ForEach-Object {
                [pscustomobject]@{
                    Index = $_.Index
                    Kind = $_.Kind
                    CollectionIndex = $_.CollectionIndex
                    ProgID = $_.ProgID
                    Type = $_.Type
                }
            })
        }
        $result | ConvertTo-Json -Depth 5
        return
    }

    if ($ObjectIndex -lt 1 -or $ObjectIndex -gt $targets.Count) {
        throw "ObjectIndex $ObjectIndex is out of range. Found $($targets.Count) embedded Visio/OLE object(s). Run -Action inspect first."
    }

    $target = $targets | Where-Object { $_.Index -eq $ObjectIndex } | Select-Object -First 1
    $activation = Activate-VisioTarget -Target $target
    $visio = $activation.VisioApp
    $visioDoc = $activation.VisioDocument

    $changes = $null
    if ($Action -eq "replace-text") {
        $changes = Replace-VisioText -VisioDocument $visioDoc -Find $FindText -Replacement $ReplaceText
    }
    elseif ($Action -eq "run-script") {
        if ([string]::IsNullOrWhiteSpace($EditScriptPath)) {
            throw "-EditScriptPath is required for run-script."
        }
        $editFull = Resolve-FullPath -Path $EditScriptPath
        if (-not (Test-Path -LiteralPath $editFull)) {
            throw "EditScriptPath does not exist: $editFull"
        }

        $VisioApp = $visio
        $VisioDocument = $visioDoc
        $TargetObject = $target.Object
        $WordDocument = $doc
        . $editFull
        $changes = "custom-script-ran"
    }

    try { $visioDoc.Save() | Out-Null } catch { }
    $doc.Save()

    [pscustomobject]@{
        OutputPath = $wordPath
        ObjectIndex = $ObjectIndex
        ProgID = $target.ProgID
        Action = $Action
        Changes = $changes
        PreservedEmbeddedObject = $true
    } | ConvertTo-Json -Depth 4
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) | Out-Null } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit() | Out-Null } catch { }
    }

    Release-ComObject -Object $doc
    Release-ComObject -Object $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

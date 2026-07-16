param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [switch]$Visible,
    [switch]$KeepOpen,
    [int]$StepDelayMs = 0,
    [switch]$ExportPdf,
    [switch]$ExportPng
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function ConvertTo-RgbFormula {
    param([string]$Color, [string]$Default = "#FFFFFF")
    if ([string]::IsNullOrWhiteSpace($Color)) {
        $Color = $Default
    }
    if ($Color -notmatch '^#[0-9a-fA-F]{6}$') {
        throw "Invalid color '$Color'. Use #RRGGBB."
    }
    $r = [Convert]::ToInt32($Color.Substring(1, 2), 16)
    $g = [Convert]::ToInt32($Color.Substring(3, 2), 16)
    $b = [Convert]::ToInt32($Color.Substring(5, 2), 16)
    return "RGB($r,$g,$b)"
}

function Get-JsonValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )
    if ($null -eq $Object) {
        return $Default
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $Default
    }
    return $property.Value
}

function Set-CellFormulaSafe {
    param(
        [Parameter(Mandatory = $true)]$Shape,
        [Parameter(Mandatory = $true)][string]$Cell,
        [Parameter(Mandatory = $true)][string]$Formula
    )
    try {
        $Shape.CellsU($Cell).FormulaU = $Formula
    } catch {
        # Some primitive shapes do not expose every style cell in all Visio versions.
    }
}

function Add-NodeShape {
    param(
        [Parameter(Mandatory = $true)]$Page,
        [Parameter(Mandatory = $true)]$Node
    )

    $x = [double](Get-JsonValue -Object $Node -Name "x")
    $y = [double](Get-JsonValue -Object $Node -Name "y")
    $w = [double](Get-JsonValue -Object $Node -Name "w" -Default 1.8)
    $h = [double](Get-JsonValue -Object $Node -Name "h" -Default 0.75)
    $type = ([string](Get-JsonValue -Object $Node -Name "type" -Default "process")).ToLowerInvariant()

    $left = $x - ($w / 2.0)
    $right = $x + ($w / 2.0)
    $bottom = $y - ($h / 2.0)
    $top = $y + ($h / 2.0)

    if ($type -in @("start", "end", "terminator")) {
        $shape = $Page.DrawOval($left, $top, $right, $bottom)
    } elseif ($type -eq "decision") {
        $size = [Math]::Max($w, $h)
        $shape = $Page.DrawRectangle($x - ($size / 2.0), $y + ($size / 2.0), $x + ($size / 2.0), $y - ($size / 2.0))
        Set-CellFormulaSafe -Shape $shape -Cell "Angle" -Formula "45 deg"
        $shape.Text = ""
        $textShape = $Page.DrawRectangle($x - ($w / 2.0), $y + ($h / 2.0), $x + ($w / 2.0), $y - ($h / 2.0))
        $textShape.Text = [string](Get-JsonValue -Object $Node -Name "text" -Default "")
        Set-CellFormulaSafe -Shape $textShape -Cell "LinePattern" -Formula "0"
        Set-CellFormulaSafe -Shape $textShape -Cell "FillPattern" -Formula "0"
        Set-CellFormulaSafe -Shape $textShape -Cell "VerticalAlign" -Formula "1"
        Set-CellFormulaSafe -Shape $textShape -Cell "Para.HorzAlign" -Formula "1"
    } else {
        $shape = $Page.DrawRectangle($left, $top, $right, $bottom)
    }

    if ($type -ne "decision") {
        $shape.Text = [string](Get-JsonValue -Object $Node -Name "text" -Default "")
    }

    Set-CellFormulaSafe -Shape $shape -Cell "FillForegnd" -Formula (ConvertTo-RgbFormula -Color (Get-JsonValue -Object $Node -Name "fill") -Default "#FFFFFF")
    Set-CellFormulaSafe -Shape $shape -Cell "LineColor" -Formula (ConvertTo-RgbFormula -Color (Get-JsonValue -Object $Node -Name "line") -Default "#404040")
    Set-CellFormulaSafe -Shape $shape -Cell "LineWeight" -Formula "1.25 pt"
    Set-CellFormulaSafe -Shape $shape -Cell "VerticalAlign" -Formula "1"
    Set-CellFormulaSafe -Shape $shape -Cell "Para.HorzAlign" -Formula "1"
    Set-CellFormulaSafe -Shape $shape -Cell "Char.Size" -Formula "10 pt"

    return $shape
}

function Get-NodeCenter {
    param($Shape)
    return @{
        x = [double]$Shape.CellsU("PinX").ResultIU
        y = [double]$Shape.CellsU("PinY").ResultIU
    }
}

function Add-Connector {
    param(
        [Parameter(Mandatory = $true)]$Page,
        [Parameter(Mandatory = $true)]$FromShape,
        [Parameter(Mandatory = $true)]$ToShape,
        $Edge
    )

    $from = Get-NodeCenter -Shape $FromShape
    $to = Get-NodeCenter -Shape $ToShape
    $line = $Page.DrawLine($from["x"], $from["y"], $to["x"], $to["y"])
    $label = Get-JsonValue -Object $Edge -Name "label"
    if ($null -ne $label -and -not [string]::IsNullOrWhiteSpace([string]$label)) {
        $line.Text = [string]$label
    }
    Set-CellFormulaSafe -Shape $line -Cell "EndArrow" -Formula "13"
    Set-CellFormulaSafe -Shape $line -Cell "LineColor" -Formula (ConvertTo-RgbFormula -Color (Get-JsonValue -Object $Edge -Name "color") -Default "#404040")
    Set-CellFormulaSafe -Shape $line -Cell "LineWeight" -Formula "1 pt"
    Set-CellFormulaSafe -Shape $line -Cell "Char.Size" -Formula "8 pt"
    try {
        $line.SendToBack()
    } catch {
        # Connector z-order is cosmetic; keep generation going if this method is unavailable.
    }
    return $line
}

function Wait-ForVisibleStep {
    param([int]$Milliseconds)
    if ($Milliseconds -gt 0) {
        Start-Sleep -Milliseconds $Milliseconds
    }
}

function Set-AutoLayout {
    param(
        [Parameter(Mandatory = $true)]$Spec,
        [double]$PageWidth,
        [double]$PageHeight
    )

    $nodes = @((Get-JsonValue -Object $Spec -Name "nodes"))
    $edgesValue = Get-JsonValue -Object $Spec -Name "edges"
    $edges = if ($null -ne $edgesValue) { @($edgesValue) } else { @() }
    $nodeById = @{}
    $indexById = @{}
    $nodeIndex = 0
    foreach ($node in $nodes) {
        $nodeId = [string](Get-JsonValue -Object $node -Name "id")
        $nodeById[$nodeId] = $node
        $indexById[$nodeId] = $nodeIndex
        $nodeIndex++
    }

    $rank = @{}
    foreach ($node in $nodes) {
        $rank[[string](Get-JsonValue -Object $node -Name "id")] = 0
    }

    $changed = $true
    $guard = 0
    while ($changed -and $guard -lt 100) {
        $changed = $false
        $guard++
        foreach ($edge in $edges) {
            $fromId = [string](Get-JsonValue -Object $edge -Name "from")
            $toId = [string](Get-JsonValue -Object $edge -Name "to")
            if ($rank.ContainsKey($fromId) -and $rank.ContainsKey($toId)) {
                if ($indexById.ContainsKey($fromId) -and $indexById.ContainsKey($toId) -and [int]$indexById[$toId] -le [int]$indexById[$fromId]) {
                    continue
                }
                $candidate = [int]$rank[$fromId] + 1
                if ($candidate -gt [int]$rank[$toId]) {
                    $rank[$toId] = $candidate
                    $changed = $true
                }
            }
        }
    }

    $groups = @{}
    foreach ($node in $nodes) {
        $id = [string](Get-JsonValue -Object $node -Name "id")
        $r = [int]$rank[$id]
        if (-not $groups.ContainsKey($r)) {
            $groups[$r] = New-Object System.Collections.ArrayList
        }
        [void]$groups[$r].Add($node)
    }

    $rankValues = @($groups.Keys | Sort-Object)
    $rankCount = [Math]::Max(1, $rankValues.Count)
    for ($ri = 0; $ri -lt $rankValues.Count; $ri++) {
        $items = @($groups[$rankValues[$ri]])
        $y = $PageHeight - 1.0 - (($PageHeight - 2.0) * ($ri / [Math]::Max(1, $rankCount - 1)))
        for ($i = 0; $i -lt $items.Count; $i++) {
            $node = $items[$i]
            if ($null -eq (Get-JsonValue -Object $node -Name "x")) {
                $node | Add-Member -NotePropertyName x -NotePropertyValue (($PageWidth * ($i + 1)) / ($items.Count + 1)) -Force
            }
            if ($null -eq (Get-JsonValue -Object $node -Name "y")) {
                $node | Add-Member -NotePropertyName y -NotePropertyValue $y -Force
            }
            if ($null -eq (Get-JsonValue -Object $node -Name "w")) {
                $node | Add-Member -NotePropertyName w -NotePropertyValue 1.8 -Force
            }
            if ($null -eq (Get-JsonValue -Object $node -Name "h")) {
                $node | Add-Member -NotePropertyName h -NotePropertyValue 0.75 -Force
            }
        }
    }
}

$fullSpecPath = Resolve-FullPath -Path $SpecPath
$fullOutputPath = Resolve-FullPath -Path $OutputPath
if (-not (Test-Path -LiteralPath $fullSpecPath)) {
    throw "SpecPath not found: $fullSpecPath"
}

$outDir = Split-Path -Parent $fullOutputPath
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$spec = Get-Content -LiteralPath $fullSpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
$specNodes = Get-JsonValue -Object $spec -Name "nodes"
if ($null -eq $specNodes -or @($specNodes).Count -eq 0) {
    throw "Spec must include at least one node."
}

$visioType = [type]::GetTypeFromProgID("Visio.Application")
if ($null -eq $visioType) {
    throw "Microsoft Visio COM registration was not found. Install Visio or repair Office COM registration."
}

$pageSpec = Get-JsonValue -Object $spec -Name "page"
$pageWidth = [double](Get-JsonValue -Object $pageSpec -Name "width" -Default 11.0)
$pageHeight = [double](Get-JsonValue -Object $pageSpec -Name "height" -Default 8.5)
Set-AutoLayout -Spec $spec -PageWidth $pageWidth -PageHeight $pageHeight

$visio = $null
$doc = $null
try {
    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = [bool]$Visible
    $doc = $visio.Documents.Add("")
    $page = $visio.ActivePage
    $page.PageSheet.CellsU("PageWidth").FormulaU = "$pageWidth in"
    $page.PageSheet.CellsU("PageHeight").FormulaU = "$pageHeight in"
    $title = Get-JsonValue -Object $spec -Name "title"
    if ($null -ne $title -and -not [string]::IsNullOrWhiteSpace([string]$title)) {
        $page.Name = [string]$title
    }

    $shapeById = @{}
    foreach ($node in @($specNodes)) {
        $nodeId = Get-JsonValue -Object $node -Name "id"
        if ($null -eq $nodeId -or [string]::IsNullOrWhiteSpace([string]$nodeId)) {
            throw "Every node requires a non-empty id."
        }
        $shape = Add-NodeShape -Page $page -Node $node
        $shapeById[[string]$nodeId] = $shape
        Wait-ForVisibleStep -Milliseconds $StepDelayMs
    }

    $specEdges = Get-JsonValue -Object $spec -Name "edges"
    if ($null -ne $specEdges) {
        foreach ($edge in @($specEdges)) {
            $fromId = [string](Get-JsonValue -Object $edge -Name "from")
            $toId = [string](Get-JsonValue -Object $edge -Name "to")
            if (-not $shapeById.ContainsKey($fromId)) {
                throw "Edge references missing from node '$fromId'."
            }
            if (-not $shapeById.ContainsKey($toId)) {
                throw "Edge references missing to node '$toId'."
            }
            [void](Add-Connector -Page $page -FromShape $shapeById[$fromId] -ToShape $shapeById[$toId] -Edge $edge)
            Wait-ForVisibleStep -Milliseconds $StepDelayMs
        }
    }

    $doc.SaveAs($fullOutputPath)

    if ($ExportPdf) {
        $pdfPath = [System.IO.Path]::ChangeExtension($fullOutputPath, ".pdf")
        $doc.ExportAsFixedFormat(1, $pdfPath, 1, 0)
    }

    if ($ExportPng) {
        $pngPath = [System.IO.Path]::ChangeExtension($fullOutputPath, ".png")
        $page.Export($pngPath)
    }

    Write-Output "Created $fullOutputPath"
} finally {
    if ($null -ne $doc -and -not $KeepOpen) {
        try { $doc.Close() } catch {}
    }
    if ($null -ne $visio -and -not $Visible -and -not $KeepOpen) {
        try { $visio.Quit() } catch {}
    }
}

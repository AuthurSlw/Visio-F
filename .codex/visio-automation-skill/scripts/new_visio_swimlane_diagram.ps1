# Generic horizontal swimlane Visio generator for the global visio-automation skill.
param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [switch]$Visible,
    [switch]$KeepOpen,
    [switch]$UpdateExisting,
    [string[]]$UpdateIds = @(),
    [switch]$ExportPng,
    [int]$StepDelayMs = 0
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

function Get-JsonValue {
    param(
        $Object,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return $property.Value
}

function ConvertTo-RgbFormula {
    param([string]$Color, [string]$Default = "#FFFFFF")
    if ([string]::IsNullOrWhiteSpace($Color)) { $Color = $Default }
    if ($Color -notmatch '^#[0-9a-fA-F]{6}$') {
        throw "Invalid color '$Color'. Use #RRGGBB."
    }
    $r = [Convert]::ToInt32($Color.Substring(1, 2), 16)
    $g = [Convert]::ToInt32($Color.Substring(3, 2), 16)
    $b = [Convert]::ToInt32($Color.Substring(5, 2), 16)
    return "RGB($r,$g,$b)"
}

function Set-CellFormulaSafe {
    param($Shape, [string]$Cell, [string]$Formula)
    try { $Shape.CellsU($Cell).FormulaU = $Formula } catch {}
}

function Wait-Step {
    if ($StepDelayMs -gt 0) {
        Start-Sleep -Milliseconds $StepDelayMs
    }
}

function Get-SafeName {
    param([Parameter(Mandatory = $true)][string]$Value)
    $safe = $Value -replace '[^A-Za-z0-9_]', '_'
    if ([string]::IsNullOrWhiteSpace($safe)) { return "item" }
    return $safe
}

function Set-ShapeNameSafe {
    param($Shape, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return }
    try { $Shape.NameU = $Name } catch {}
}

function Remove-ShapesByPrefix {
    param($Page, [string]$Prefix)
    if ([string]::IsNullOrWhiteSpace($Prefix)) { return }
    for ($i = $Page.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Page.Shapes.Item($i)
        try {
            if ([string]$shape.NameU -like "$Prefix*") {
                $shape.Delete()
            }
        } catch {}
    }
}

function Bring-ShapesByPrefixToFront {
    param($Page, [string]$Prefix)
    if ([string]::IsNullOrWhiteSpace($Prefix)) { return }
    foreach ($shape in @($Page.Shapes)) {
        try {
            if ([string]$shape.NameU -like "$Prefix*") {
                [void]$shape.BringToFront()
            }
        } catch {}
    }
}

function Bring-ShapesByNamePatternToFront {
    param($Page, [string]$Pattern)
    if ([string]::IsNullOrWhiteSpace($Pattern)) { return }
    foreach ($shape in @($Page.Shapes)) {
        try {
            if ([string]$shape.NameU -like $Pattern) {
                [void]$shape.BringToFront()
            }
        } catch {}
    }
}

function Clear-GeneratedContent {
    param($Page, [bool]$ClearAll = $false)
    for ($i = $Page.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Page.Shapes.Item($i)
        try {
            if ($ClearAll -or [string]$shape.NameU -like "visio_auto_*") {
                $shape.Delete()
            }
        } catch {}
    }
}

function Should-Update {
    param([string]$Id)
    if ($script:UpdateIdSet.Count -eq 0) { return $true }
    return $script:UpdateIdSet.ContainsKey($Id)
}

function Set-JsonValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)]$Value
    )
    $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
}

function Scale-X {
    param([double]$Value)
    return $Value * $script:ScaleX
}

function Scale-Y {
    param([double]$Value)
    return $Value * $script:ScaleY
}

function Scale-W {
    param([double]$Value)
    return $Value * $script:ScaleX
}

function Scale-H {
    param([double]$Value)
    return $Value * $script:ScaleY
}

function Scale-Font {
    param([double]$Value)
    $scaled = $Value * $script:TextScale
    $maxFont = 14.0
    $maxFontVariable = Get-Variable -Scope Script -Name MaxFontSize -ErrorAction SilentlyContinue
    if ($null -ne $maxFontVariable) {
        $maxFont = [double]$script:MaxFontSize
    }
    return [Math]::Max(6.5, [Math]::Min($maxFont, $scaled))
}

function Get-TextVisualLength {
    param([string]$Text)
    $length = 0.0
    foreach ($char in ([string]$Text).ToCharArray()) {
        $code = [int][char]$char
        if ($code -le 127) {
            $length += 0.55
        } else {
            $length += 1.0
        }
    }
    return [Math]::Max(1.0, $length)
}

function Get-AutoFontSize {
    param(
        [string]$Text,
        [double]$LogicalWidth,
        [double]$LogicalHeight,
        [double]$PreferredFontSize = 9
    )
    $scaledWidth = [Math]::Max(0.2, (Scale-W -Value $LogicalWidth) - 0.16)
    $scaledHeight = [Math]::Max(0.18, (Scale-H -Value $LogicalHeight) - 0.10)
    $lines = @(([string]$Text) -split "`n")
    $lineCount = [Math]::Max(1, $lines.Count)
    $maxVisualLength = 1.0
    foreach ($line in $lines) {
        $maxVisualLength = [Math]::Max($maxVisualLength, (Get-TextVisualLength -Text ([string]$line)))
    }
    $widthFactor = if ($script:ReferenceStyle) { 0.86 } else { 0.92 }
    $lineHeightFactor = if ($script:ReferenceStyle) { 1.10 } else { 1.18 }
    $fontScale = 1.0
    $fontScaleVariable = Get-Variable -Scope Script -Name AdaptiveFontScale -ErrorAction SilentlyContinue
    if ($null -ne $fontScaleVariable) {
        $fontScale = [double]$script:AdaptiveFontScale
    }
    $byWidth = (($scaledWidth * 72.0) / ($maxVisualLength * $widthFactor)) * $fontScale
    $byHeight = (($scaledHeight * 72.0) / ($lineCount * $lineHeightFactor)) * $fontScale
    $preferred = Scale-Font -Value $PreferredFontSize
    $fit = [Math]::Min($script:MaxFontSize, [Math]::Min($preferred, [Math]::Min($byWidth, $byHeight)))
    if ($fit -lt $script:MinFontSize) {
        return [Math]::Max(6.5, $fit)
    }
    return [Math]::Max($script:MinFontSize, $fit)
}

function Set-ShapeStyle {
    param($Shape, [string]$Fill, [string]$Line = "#333333", [double]$FontSize = 9)
    Set-CellFormulaSafe -Shape $Shape -Cell "FillForegnd" -Formula (ConvertTo-RgbFormula -Color $Fill -Default "#FFFFFF")
    Set-CellFormulaSafe -Shape $Shape -Cell "LineColor" -Formula (ConvertTo-RgbFormula -Color $Line -Default "#333333")
    $weight = if ($script:ReferenceStyle) { 1.8 } else { 1.25 }
    Set-CellFormulaSafe -Shape $Shape -Cell "LineWeight" -Formula "$([Math]::Max(0.8, $weight * $script:TextScale)) pt"
    Set-CellFormulaSafe -Shape $Shape -Cell "VerticalAlign" -Formula "1"
    Set-CellFormulaSafe -Shape $Shape -Cell "Para.HorzAlign" -Formula "1"
    Set-CellFormulaSafe -Shape $Shape -Cell "Char.Size" -Formula "$(Scale-Font -Value $FontSize) pt"
}

function Add-RectText {
    param($Page, [double]$X, [double]$Y, [double]$W, [double]$H, [string]$Text, [string]$Fill = "#FFFFFF", [string]$Line = "#333333", [double]$FontSize = 9, [string]$Name = "")
    $sx = Scale-X -Value $X
    $sy = Scale-Y -Value $Y
    $sw = Scale-W -Value $W
    $sh = Scale-H -Value $H
    $shape = $Page.DrawRectangle($sx - ($sw / 2.0), $sy + ($sh / 2.0), $sx + ($sw / 2.0), $sy - ($sh / 2.0))
    $shape.Text = $Text
    Set-ShapeStyle -Shape $shape -Fill $Fill -Line $Line -FontSize (Get-AutoFontSize -Text $Text -LogicalWidth $W -LogicalHeight $H -PreferredFontSize $FontSize)
    Set-ShapeNameSafe -Shape $shape -Name $Name
    return $shape
}

function Get-LogicalNodeRect {
    param($Node)
    $x = [double](Get-JsonValue -Object $Node -Name "x")
    $y = [double](Get-JsonValue -Object $Node -Name "y")
    $w = [double](Get-JsonValue -Object $Node -Name "w" -Default 1.8)
    $h = [double](Get-JsonValue -Object $Node -Name "h" -Default 0.65)
    $type = ([string](Get-JsonValue -Object $Node -Name "type" -Default "process")).ToLowerInvariant()
    if ($type -eq "decision") {
        $size = [Math]::Max($w, $h)
        $w = $size
        $h = $size
    }
    return @{
        x = $x
        y = $y
        w = $w
        h = $h
        left = $x - ($w / 2.0)
        right = $x + ($w / 2.0)
        top = $y + ($h / 2.0)
        bottom = $y - ($h / 2.0)
    }
}

function Test-RectOverlap {
    param($A, $B, [double]$Padding = 0.12)
    return -not (
        ([double]$A["right"] + $Padding) -le [double]$B["left"] -or
        ([double]$A["left"] - $Padding) -ge [double]$B["right"] -or
        ([double]$A["top"] + $Padding) -le [double]$B["bottom"] -or
        ([double]$A["bottom"] - $Padding) -ge [double]$B["top"]
    )
}

function Test-AnyRectOverlap {
    param($Rect)
    foreach ($placed in @($script:PlacedRects)) {
        if (Test-RectOverlap -A $Rect -B $placed -Padding $script:ShapePadding) {
            return $true
        }
    }
    return $false
}

function Get-LaneForX {
    param([double]$X)
    foreach ($lane in @($script:Lanes)) {
        $x1 = [double](Get-JsonValue -Object $lane -Name "x1")
        $x2 = [double](Get-JsonValue -Object $lane -Name "x2")
        if ($X -ge $x1 -and $X -le $x2) {
            return $lane
        }
    }
    return $null
}

function Resolve-NodePlacement {
    param($Node)
    $originalX = [double](Get-JsonValue -Object $Node -Name "x")
    $originalY = [double](Get-JsonValue -Object $Node -Name "y")
    $lane = Get-LaneForX -X $originalX
    $laneLeft = 0.4
    $laneRight = $script:BaseWidth - 0.4
    if ($null -ne $lane) {
        $laneLeft = [double](Get-JsonValue -Object $lane -Name "x1") + 0.18
        $laneRight = [double](Get-JsonValue -Object $lane -Name "x2") - 0.18
    }
    $rect = Get-LogicalNodeRect -Node $Node
    $stepY = [double]$rect["h"] + $script:ShapePadding + 0.18
    $maxTop = $script:BaseHeight - 1.10
    $minBottom = 0.55

    for ($attempt = 0; $attempt -lt 64; $attempt++) {
        if ($attempt -eq 0) {
            $candidateX = $originalX
            $candidateY = $originalY
        } else {
            $band = [Math]::Ceiling($attempt / 2.0)
            $direction = if (($attempt % 2) -eq 1) { -1 } else { 1 }
            $candidateY = $originalY + ($direction * $band * $stepY)
            $xBand = [Math]::Floor(($attempt - 1) / 8)
            $xOffset = (($xBand % 3) - 1) * [Math]::Min(0.35, (($laneRight - $laneLeft) * 0.08))
            $candidateX = $originalX + $xOffset
        }

        $candidateX = [Math]::Max($laneLeft + ([double]$rect["w"] / 2.0), [Math]::Min($laneRight - ([double]$rect["w"] / 2.0), $candidateX))
        $candidateY = [Math]::Max($minBottom + ([double]$rect["h"] / 2.0), [Math]::Min($maxTop - ([double]$rect["h"] / 2.0), $candidateY))
        Set-JsonValue -Object $Node -Name "x" -Value $candidateX
        Set-JsonValue -Object $Node -Name "y" -Value $candidateY
        $candidateRect = Get-LogicalNodeRect -Node $Node
        if (-not (Test-AnyRectOverlap -Rect $candidateRect)) {
            $script:PlacedRects += $candidateRect
            return $Node
        }
    }

    $script:PlacedRects += (Get-LogicalNodeRect -Node $Node)
    return $Node
}

function Add-Node {
    param($Page, $Node, [string]$Name)
    $x = [double](Get-JsonValue -Object $Node -Name "x")
    $y = [double](Get-JsonValue -Object $Node -Name "y")
    $w = [double](Get-JsonValue -Object $Node -Name "w" -Default 1.8)
    $h = [double](Get-JsonValue -Object $Node -Name "h" -Default 0.65)
    $type = ([string](Get-JsonValue -Object $Node -Name "type" -Default "process")).ToLowerInvariant()
    $text = [string](Get-JsonValue -Object $Node -Name "text" -Default "")
    $defaultNodeFontSize = 9
    if ($script:ReferenceStyle) {
        $defaultNodeFontSize = 10
    }
    $fontSize = [double](Get-JsonValue -Object $Node -Name "fontSize" -Default $defaultNodeFontSize)
    $fill = [string](Get-JsonValue -Object $Node -Name "fill" -Default "#FFFFFF")
    if ($type -eq "document") { $fill = [string](Get-JsonValue -Object $Node -Name "fill" -Default "#CFFFD4") }
    if ($type -eq "decision") { $fill = [string](Get-JsonValue -Object $Node -Name "fill" -Default "#FFD19B") }

    if ($type -in @("start", "end", "terminator")) {
        $sx = Scale-X -Value $x
        $sy = Scale-Y -Value $y
        $sw = Scale-W -Value $w
        $sh = Scale-H -Value $h
        $shape = $Page.DrawOval($sx - ($sw / 2.0), $sy + ($sh / 2.0), $sx + ($sw / 2.0), $sy - ($sh / 2.0))
        $shape.Text = $text
        Set-ShapeStyle -Shape $shape -Fill $fill -FontSize (Get-AutoFontSize -Text $text -LogicalWidth $w -LogicalHeight $h -PreferredFontSize $fontSize)
        Set-ShapeNameSafe -Shape $shape -Name $Name
        return $shape
    }

    if ($type -eq "decision") {
        $size = [Math]::Max((Scale-W -Value $w), (Scale-H -Value $h))
        $sx = Scale-X -Value $x
        $sy = Scale-Y -Value $y
        $shape = $Page.DrawRectangle($sx - ($size / 2.0), $sy + ($size / 2.0), $sx + ($size / 2.0), $sy - ($size / 2.0))
        Set-CellFormulaSafe -Shape $shape -Cell "Angle" -Formula "45 deg"
        $shape.Text = ""
        Set-ShapeStyle -Shape $shape -Fill $fill
        Set-ShapeNameSafe -Shape $shape -Name $Name
        $textShape = Add-RectText -Page $Page -X $x -Y $y -W $w -H $h -Text $text -Fill "#FFFFFF" -Line "#FFFFFF" -FontSize $fontSize -Name "$Name`_text"
        Set-CellFormulaSafe -Shape $textShape -Cell "FillPattern" -Formula "0"
        Set-CellFormulaSafe -Shape $textShape -Cell "LinePattern" -Formula "0"
        return $shape
    }

    $shape = Add-RectText -Page $Page -X $x -Y $y -W $w -H $h -Text $text -Fill $fill -FontSize $fontSize -Name $Name
    return $shape
}

function Get-NodeInfo {
    param($Node)
    $type = ([string](Get-JsonValue -Object $Node -Name "type" -Default "process")).ToLowerInvariant()
    $w = Scale-W -Value ([double](Get-JsonValue -Object $Node -Name "w" -Default 1.8))
    $h = Scale-H -Value ([double](Get-JsonValue -Object $Node -Name "h" -Default 0.65))
    if ($type -eq "decision") {
        $size = [Math]::Max($w, $h)
        $w = $size
        $h = $size
    }
    return @{
        x = Scale-X -Value ([double](Get-JsonValue -Object $Node -Name "x"))
        y = Scale-Y -Value ([double](Get-JsonValue -Object $Node -Name "y"))
        w = $w
        h = $h
        type = $type
    }
}

function Get-Center {
    param($Shape)
    return @{
        x = [double]$Shape.CellsU("PinX").ResultIU
        y = [double]$Shape.CellsU("PinY").ResultIU
    }
}

function Get-BoundaryPoint {
    param($FromInfo, $ToInfo)
    $dx = [double]$ToInfo["x"] - [double]$FromInfo["x"]
    $dy = [double]$ToInfo["y"] - [double]$FromInfo["y"]
    if ([Math]::Abs($dx) -ge [Math]::Abs($dy)) {
        $sign = if ($dx -ge 0) { 1 } else { -1 }
        return @{
            x = [double]$FromInfo["x"] + ($sign * [double]$FromInfo["w"] / 2.0)
            y = [double]$FromInfo["y"]
        }
    }
    $signY = if ($dy -ge 0) { 1 } else { -1 }
    return @{
        x = [double]$FromInfo["x"]
        y = [double]$FromInfo["y"] + ($signY * [double]$FromInfo["h"] / 2.0)
    }
}

function Get-PortPoint {
    param($Info, [string]$Port)
    $normalized = (([string]$Port).ToLowerInvariant() -replace "_", "-" -replace " ", "-")
    if ($normalized -eq "left") {
        return @{ x = [double]$Info["x"] - ([double]$Info["w"] / 2.0); y = [double]$Info["y"] }
    }
    if ($normalized -eq "right") {
        return @{ x = [double]$Info["x"] + ([double]$Info["w"] / 2.0); y = [double]$Info["y"] }
    }
    if ($normalized -eq "top") {
        return @{ x = [double]$Info["x"]; y = [double]$Info["y"] + ([double]$Info["h"] / 2.0) }
    }
    if ($normalized -eq "bottom") {
        return @{ x = [double]$Info["x"]; y = [double]$Info["y"] - ([double]$Info["h"] / 2.0) }
    }
    if ($normalized -in @("top-left", "left-top")) {
        return @{ x = [double]$Info["x"] - ([double]$Info["w"] / 2.0); y = [double]$Info["y"] + ([double]$Info["h"] / 2.0) }
    }
    if ($normalized -in @("top-right", "right-top")) {
        return @{ x = [double]$Info["x"] + ([double]$Info["w"] / 2.0); y = [double]$Info["y"] + ([double]$Info["h"] / 2.0) }
    }
    if ($normalized -in @("bottom-left", "left-bottom")) {
        return @{ x = [double]$Info["x"] - ([double]$Info["w"] / 2.0); y = [double]$Info["y"] - ([double]$Info["h"] / 2.0) }
    }
    if ($normalized -in @("bottom-right", "right-bottom")) {
        return @{ x = [double]$Info["x"] + ([double]$Info["w"] / 2.0); y = [double]$Info["y"] - ([double]$Info["h"] / 2.0) }
    }
    return $null
}

function Test-AllowedConnectorPort {
    param([string]$Port)
    $normalized = (([string]$Port).ToLowerInvariant() -replace "_", "-" -replace " ", "-")
    return $normalized -in @(
        "left", "right", "top", "bottom",
        "top-left", "left-top", "top-right", "right-top",
        "bottom-left", "left-bottom", "bottom-right", "right-bottom"
    )
}

function Get-PortOutwardOffset {
    param([string]$Port, [double]$Distance)
    $normalized = (([string]$Port).ToLowerInvariant() -replace "_", "-" -replace " ", "-")
    if ($normalized -like "*left*") {
        return @{ dx = -$Distance; dy = 0.0 }
    }
    if ($normalized -like "*right*") {
        return @{ dx = $Distance; dy = 0.0 }
    }
    if ($normalized -like "*top*") {
        return @{ dx = 0.0; dy = $Distance }
    }
    if ($normalized -like "*bottom*") {
        return @{ dx = 0.0; dy = -$Distance }
    }
    return @{ dx = 0.0; dy = 0.0 }
}

function Test-SamePoint {
    param($P1, $P2)
    return ([Math]::Abs([double]$P1["x"] - [double]$P2["x"]) -lt 0.001 -and [Math]::Abs([double]$P1["y"] - [double]$P2["y"]) -lt 0.001)
}

function Add-PointIfDistinct {
    param($List, $Point)
    if ($List.Count -eq 0 -or -not (Test-SamePoint -P1 $List[$List.Count - 1] -P2 $Point)) {
        [void]$List.Add($Point)
    }
}

function Add-OrthogonalPoint {
    param($List, $Point)
    if ($List.Count -eq 0) {
        Add-PointIfDistinct -List $List -Point $Point
        return
    }
    $last = $List[$List.Count - 1]
    if (([Math]::Abs([double]$last["x"] - [double]$Point["x"]) -gt 0.001) -and ([Math]::Abs([double]$last["y"] - [double]$Point["y"]) -gt 0.001)) {
        Add-PointIfDistinct -List $List -Point @{ x = [double]$Point["x"]; y = [double]$last["y"] }
    }
    Add-PointIfDistinct -List $List -Point $Point
}

function Add-ConnectorPortStubs {
    param($Points, [string]$FromPortName, [string]$ToPortName)
    if ($script:ConnectorPortStub -le 0 -or $Points.Count -lt 2) {
        return $Points
    }
    $distance = [Math]::Max(0.04, $script:ConnectorPortStub * $script:TextScale)
    $terminalClearance = [Math]::Max(0.0, $script:TargetTerminalArrowVisualClearance * $script:TextScale)
    $start = $Points[0]
    $end = $Points[$Points.Count - 1]
    $fromOffset = Get-PortOutwardOffset -Port $FromPortName -Distance $distance
    $toOffset = Get-PortOutwardOffset -Port $ToPortName -Distance $distance
    $terminalOffset = Get-PortOutwardOffset -Port $ToPortName -Distance $terminalClearance
    $startStub = @{ x = [double]$start["x"] + [double]$fromOffset["dx"]; y = [double]$start["y"] + [double]$fromOffset["dy"] }
    $endStub = @{ x = [double]$end["x"] + [double]$toOffset["dx"]; y = [double]$end["y"] + [double]$toOffset["dy"] }
    $terminalEnd = @{ x = [double]$end["x"] + [double]$terminalOffset["dx"]; y = [double]$end["y"] + [double]$terminalOffset["dy"] }

    $expanded = New-Object System.Collections.ArrayList
    Add-PointIfDistinct -List $expanded -Point $start
    Add-OrthogonalPoint -List $expanded -Point $startStub
    for ($i = 1; $i -lt ($Points.Count - 1); $i++) {
        Add-OrthogonalPoint -List $expanded -Point $Points[$i]
    }
    Add-OrthogonalPoint -List $expanded -Point $endStub
    if ($script:TargetTerminalArrowVisualClearance -gt 0) {
        Add-OrthogonalPoint -List $expanded -Point $terminalEnd
    } elseif (-not $script:TargetTerminalArrowOutsideModule) {
        Add-OrthogonalPoint -List $expanded -Point $end
    }
    return $expanded
}

function Add-LineSegment {
    param($Page, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [bool]$Arrow, [string]$Label, [string]$Name = "")
    $line = $Page.DrawLine($X1, $Y1, $X2, $Y2)
    if ($Arrow) {
        Set-CellFormulaSafe -Shape $line -Cell "EndArrow" -Formula "13"
        Set-CellFormulaSafe -Shape $line -Cell "EndArrowSize" -Formula "$script:TargetTerminalArrowSize"
    }
    if ($Arrow -and -not [string]::IsNullOrWhiteSpace($Label)) {
        $line.Text = $Label
    }
    Set-CellFormulaSafe -Shape $line -Cell "LineColor" -Formula "RGB(35,40,45)"
    $lineWeight = if ($script:ReferenceStyle) { 1.8 } else { 1.1 }
    Set-CellFormulaSafe -Shape $line -Cell "LineWeight" -Formula "$([Math]::Max(0.8, $lineWeight * $script:TextScale)) pt"
    Set-CellFormulaSafe -Shape $line -Cell "Char.Size" -Formula "$(Scale-Font -Value 7) pt"
    Set-ShapeNameSafe -Shape $line -Name $Name
    return $line
}

function Get-RouteOffset {
    param([string]$Key)
    if (-not $script:RouteUsage.ContainsKey($Key)) {
        $script:RouteUsage[$Key] = 0
    }
    $index = [int]$script:RouteUsage[$Key]
    $script:RouteUsage[$Key] = $index + 1
    if ($index -eq 0) {
        return 0.0
    }
    $band = [Math]::Ceiling($index / 2.0)
    $direction = if (($index % 2) -eq 1) { 1 } else { -1 }
    return $direction * $band * [Math]::Max(0.08, 0.14 * $script:TextScale)
}

function Get-PointKey {
    param($Point)
    return "$([Math]::Round([double]$Point["x"], 3)),$([Math]::Round([double]$Point["y"], 3))"
}

function Get-SegmentKey {
    param($P1, $P2)
    $k1 = Get-PointKey -Point $P1
    $k2 = Get-PointKey -Point $P2
    if ([string]::CompareOrdinal($k1, $k2) -le 0) {
        return "$k1|$k2"
    }
    return "$k2|$k1"
}

function Register-ConnectorInterface {
    param([string]$Key, [string]$Owner)
    if (-not $script:DisallowSharedConnectorSegments) {
        return
    }
    if ($script:ConnectorInterfaceUsage.ContainsKey($Key)) {
        throw "Connector interface '$Key' is already used by '$($script:ConnectorInterfaceUsage[$Key])'. '$Owner' must use another module input/output port."
    }
    $script:ConnectorInterfaceUsage[$Key] = $Owner
}

function Register-ConnectorSegment {
    param($P1, $P2, [string]$Owner)
    if (-not $script:DisallowSharedConnectorSegments) {
        return
    }
    if ([Math]::Abs([double]$P1["x"] - [double]$P2["x"]) -lt 0.001 -and [Math]::Abs([double]$P1["y"] - [double]$P2["y"]) -lt 0.001) {
        return
    }
    $key = Get-SegmentKey -P1 $P1 -P2 $P2
    if ($script:ConnectorSegmentUsage.ContainsKey($key)) {
        throw "Connector segment '$key' is already used by '$($script:ConnectorSegmentUsage[$key])'. '$Owner' must use a separate route."
    }
    $script:ConnectorSegmentUsage[$key] = $Owner
}

function Test-SegmentIntersectsNodeInterior {
    param($P1, $P2, $Info, [double]$Eps = 0.015)
    $left = [double]$Info["x"] - ([double]$Info["w"] / 2.0)
    $right = [double]$Info["x"] + ([double]$Info["w"] / 2.0)
    $bottom = [double]$Info["y"] - ([double]$Info["h"] / 2.0)
    $top = [double]$Info["y"] + ([double]$Info["h"] / 2.0)
    $x1 = [double]$P1["x"]
    $y1 = [double]$P1["y"]
    $x2 = [double]$P2["x"]
    $y2 = [double]$P2["y"]

    if ([Math]::Abs($x1 - $x2) -lt 0.001) {
        $minY = [Math]::Min($y1, $y2)
        $maxY = [Math]::Max($y1, $y2)
        return ($x1 -gt ($left + $Eps) -and $x1 -lt ($right - $Eps) -and $maxY -gt ($bottom + $Eps) -and $minY -lt ($top - $Eps))
    }
    if ([Math]::Abs($y1 - $y2) -lt 0.001) {
        $minX = [Math]::Min($x1, $x2)
        $maxX = [Math]::Max($x1, $x2)
        return ($y1 -gt ($bottom + $Eps) -and $y1 -lt ($top - $Eps) -and $maxX -gt ($left + $Eps) -and $minX -lt ($right - $Eps))
    }
    return $true
}

function Assert-ConnectorSegmentIsValid {
    param($P1, $P2, [string]$Owner, [string]$FromId, [string]$ToId)
    if ($script:RequireOrthogonalConnectors) {
        $isVertical = [Math]::Abs([double]$P1["x"] - [double]$P2["x"]) -lt 0.001
        $isHorizontal = [Math]::Abs([double]$P1["y"] - [double]$P2["y"]) -lt 0.001
        if (-not ($isVertical -or $isHorizontal)) {
            throw "$Owner has a diagonal connector segment. Use orthogonal waypoints from module edge center/corner ports."
        }
    }
    if (-not $script:DisallowConnectorIntersections -or $null -eq $script:NodeInfoById) {
        return
    }
    foreach ($nodeId in @($script:NodeInfoById.Keys)) {
        $info = $script:NodeInfoById[$nodeId]
        if (Test-SegmentIntersectsNodeInterior -P1 $P1 -P2 $P2 -Info $info) {
            throw "$Owner connector segment enters module '$nodeId'. Route from edge center/corner ports around module interiors."
        }
    }
}

function Add-Arrow {
    param($Page, $FromInfo, $ToInfo, $Edge, [string]$NamePrefix, [string]$FromId = "", [string]$ToId = "", [int]$EdgeIndex = -1)
    $fromPort = Get-JsonValue -Object $Edge -Name "fromPort"
    $toPort = Get-JsonValue -Object $Edge -Name "toPort"
    if ($script:RequireExplicitConnectorPorts -and ($null -eq $fromPort -or $null -eq $toPort)) {
        throw "Connector '$FromId->$ToId' must set explicit fromPort and toPort using edge center or corner ports."
    }
    if ($null -ne $fromPort -and -not (Test-AllowedConnectorPort -Port ([string]$fromPort))) {
        throw "Connector '$FromId->$ToId' has invalid fromPort '$fromPort'. Use left/right/top/bottom or a corner port."
    }
    if ($null -ne $toPort -and -not (Test-AllowedConnectorPort -Port ([string]$toPort))) {
        throw "Connector '$FromId->$ToId' has invalid toPort '$toPort'. Use left/right/top/bottom or a corner port."
    }
    $fromPortName = if ($null -ne $fromPort) { ([string]$fromPort).ToLowerInvariant() } else { "auto" }
    $toPortName = if ($null -ne $toPort) { ([string]$toPort).ToLowerInvariant() } else { "auto" }
    $edgeOwner = if ($EdgeIndex -ge 0) { "edge:$EdgeIndex $FromId->$ToId" } else { "$FromId->$ToId" }
    if (-not [string]::IsNullOrWhiteSpace($FromId)) {
        Register-ConnectorInterface -Key "from:${FromId}:$fromPortName" -Owner $edgeOwner
    }
    if (-not [string]::IsNullOrWhiteSpace($ToId)) {
        Register-ConnectorInterface -Key "to:${ToId}:$toPortName" -Owner $edgeOwner
    }
    $from = if ($null -ne $fromPort) { Get-PortPoint -Info $FromInfo -Port ([string]$fromPort) } else { $null }
    $to = if ($null -ne $toPort) { Get-PortPoint -Info $ToInfo -Port ([string]$toPort) } else { $null }
    if ($null -eq $from) { $from = Get-BoundaryPoint -FromInfo $FromInfo -ToInfo $ToInfo }
    if ($null -eq $to) { $to = Get-BoundaryPoint -FromInfo $ToInfo -ToInfo $FromInfo }
    $label = Get-JsonValue -Object $Edge -Name "label"
    $labelText = if ($null -ne $label) { [string]$label } else { "" }
    $points = New-Object System.Collections.ArrayList
    [void]$points.Add($from)
    $waypoints = Get-JsonValue -Object $Edge -Name "waypoints"
    if ($null -ne $waypoints) {
        foreach ($waypoint in @($waypoints)) {
            $waypointX = Get-JsonValue -Object $waypoint -Name "x"
            $waypointY = Get-JsonValue -Object $waypoint -Name "y"
            if ($null -eq $waypointX -or $null -eq $waypointY) {
                continue
            }
            [void]$points.Add(@{
                x = Scale-X -Value ([double]$waypointX)
                y = Scale-Y -Value ([double]$waypointY)
            })
        }
    } else {
        $dx = [double]$to["x"] - [double]$from["x"]
        $dy = [double]$to["y"] - [double]$from["y"]
        if ([Math]::Abs($dx) -gt 0.25 -and [Math]::Abs($dy) -gt 0.25) {
            if ([Math]::Abs($dx) -ge [Math]::Abs($dy)) {
                $midX = ([double]$from["x"] + [double]$to["x"]) / 2.0
                $midX += Get-RouteOffset -Key ("v:" + [Math]::Round($midX, 1))
                [void]$points.Add(@{ x = $midX; y = [double]$from["y"] })
                [void]$points.Add(@{ x = $midX; y = [double]$to["y"] })
            } else {
                $midY = ([double]$from["y"] + [double]$to["y"]) / 2.0
                $midY += Get-RouteOffset -Key ("h:" + [Math]::Round($midY, 1))
                [void]$points.Add(@{ x = [double]$from["x"]; y = $midY })
                [void]$points.Add(@{ x = [double]$to["x"]; y = $midY })
            }
        } elseif ([Math]::Abs($dx) -ge [Math]::Abs($dy)) {
            $routeKey = "h:" + [Math]::Round([double]$from["y"], 1) + ":" + [Math]::Round([Math]::Min([double]$from["x"], [double]$to["x"]), 1) + "-" + [Math]::Round([Math]::Max([double]$from["x"], [double]$to["x"]), 1)
            $offset = Get-RouteOffset -Key $routeKey
            if ([Math]::Abs($offset) -gt 0.001) {
                [void]$points.Add(@{ x = [double]$from["x"]; y = ([double]$from["y"] + $offset) })
                [void]$points.Add(@{ x = [double]$to["x"]; y = ([double]$to["y"] + $offset) })
            }
        } else {
            $routeKey = "v:" + [Math]::Round([double]$from["x"], 1) + ":" + [Math]::Round([Math]::Min([double]$from["y"], [double]$to["y"]), 1) + "-" + [Math]::Round([Math]::Max([double]$from["y"], [double]$to["y"]), 1)
            $offset = Get-RouteOffset -Key $routeKey
            if ([Math]::Abs($offset) -gt 0.001) {
                [void]$points.Add(@{ x = ([double]$from["x"] + $offset); y = [double]$from["y"] })
                [void]$points.Add(@{ x = ([double]$to["x"] + $offset); y = [double]$to["y"] })
            }
        }
    }
    [void]$points.Add($to)
    $points = Add-ConnectorPortStubs -Points $points -FromPortName $fromPortName -ToPortName $toPortName

    $lastLine = $null
    for ($i = 0; $i -lt ($points.Count - 1); $i++) {
        $isLast = ($i -eq ($points.Count - 2))
        $p1 = $points[$i]
        $p2 = $points[$i + 1]
        Assert-ConnectorSegmentIsValid -P1 $p1 -P2 $p2 -Owner $edgeOwner -FromId $FromId -ToId $ToId
        Register-ConnectorSegment -P1 $p1 -P2 $p2 -Owner $edgeOwner
        $segmentName = if ($isLast) { "$NamePrefix`_terminal_arrow" } else { "$NamePrefix`_seg_$i" }
        $lastLine = Add-LineSegment -Page $Page -X1 ([double]$p1["x"]) -Y1 ([double]$p1["y"]) -X2 ([double]$p2["x"]) -Y2 ([double]$p2["y"]) -Arrow $isLast -Label $labelText -Name $segmentName
    }
    return $lastLine
}

function Get-ActiveOrNewVisio {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application")
    } catch {
        return New-Object -ComObject Visio.Application
    }
}

function Find-OpenDocument {
    param($Visio, [string]$FullPath)
    foreach ($candidate in @($Visio.Documents)) {
        try {
            if ([string]::Equals([System.IO.Path]::GetFullPath($candidate.FullName), $FullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $candidate
            }
        } catch {}
    }
    return $null
}

function Get-ActiveDocumentSafe {
    param($Visio)
    try {
        if ($null -ne $Visio.ActiveDocument -and $Visio.ActiveDocument.Pages.Count -gt 0) {
            return $Visio.ActiveDocument
        }
    } catch {}
    return $null
}

$fullSpecPath = Resolve-FullPath -Path $SpecPath
$fullOutputPath = Resolve-FullPath -Path $OutputPath
$outDir = Split-Path -Parent $fullOutputPath
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$spec = Get-Content -LiteralPath $fullSpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
$script:UpdateIdSet = @{}
foreach ($id in @($UpdateIds)) {
    if (-not [string]::IsNullOrWhiteSpace($id)) {
        $script:UpdateIdSet[[string]$id] = $true
    }
}
$visioType = [type]::GetTypeFromProgID("Visio.Application")
if ($null -eq $visioType) {
    throw "Microsoft Visio COM registration was not found."
}

$visio = $null
$doc = $null
try {
    if ($UpdateExisting) {
        $visio = Get-ActiveOrNewVisio
    } else {
        $visio = New-Object -ComObject Visio.Application
    }
    $visio.Visible = [bool]$Visible
    if ($UpdateExisting) {
        $doc = Find-OpenDocument -Visio $visio -FullPath $fullOutputPath
        if ($null -eq $doc) {
            $doc = Get-ActiveDocumentSafe -Visio $visio
        }
        if ($null -eq $doc -and (Test-Path -LiteralPath $fullOutputPath)) {
            $doc = $visio.Documents.Open($fullOutputPath)
        }
    }
    if ($null -eq $doc) {
        $doc = $visio.Documents.Add("")
    }
    $page = $visio.ActivePage
    if ($UpdateExisting -and $null -ne $doc.Pages -and $doc.Pages.Count -gt 0) {
        $page = $doc.Pages.Item(1)
        $visio.ActiveWindow.Page = $page
    }
    $pageWidth = [double](Get-JsonValue -Object (Get-JsonValue -Object $spec -Name "page") -Name "width" -Default 27)
    $pageHeight = [double](Get-JsonValue -Object (Get-JsonValue -Object $spec -Name "page") -Name "height" -Default 16)
    $baseWidth = [double](Get-JsonValue -Object (Get-JsonValue -Object $spec -Name "page") -Name "baseWidth" -Default 27)
    $baseHeight = [double](Get-JsonValue -Object (Get-JsonValue -Object $spec -Name "page") -Name "baseHeight" -Default 16)
    $script:ScaleX = $pageWidth / $baseWidth
    $script:ScaleY = $pageHeight / $baseHeight
    $script:TextScale = [Math]::Min($script:ScaleX, $script:ScaleY)
    $script:BaseWidth = $baseWidth
    $script:BaseHeight = $baseHeight
    $script:Lanes = @((Get-JsonValue -Object $spec -Name "lanes"))
    $layoutSpec = Get-JsonValue -Object $spec -Name "layout"
    $script:ShapePadding = [double](Get-JsonValue -Object $layoutSpec -Name "shapePadding" -Default 0.16)
    $script:AutoAvoidOverlap = [bool](Get-JsonValue -Object $layoutSpec -Name "autoAvoidOverlap" -Default $true)
    $script:ReferenceStyle = [bool](Get-JsonValue -Object $layoutSpec -Name "referenceStyle" -Default $false)
    $script:ClearPageOnFullRedraw = [bool](Get-JsonValue -Object $layoutSpec -Name "clearPageOnFullRedraw" -Default $false)
    $script:DisallowSharedConnectorSegments = [bool](Get-JsonValue -Object $layoutSpec -Name "disallowSharedConnectorSegments" -Default $true)
    $script:RequireExplicitConnectorPorts = [bool](Get-JsonValue -Object $layoutSpec -Name "requireExplicitConnectorPorts" -Default $true)
    $script:RequireOrthogonalConnectors = [bool](Get-JsonValue -Object $layoutSpec -Name "requireOrthogonalConnectors" -Default $true)
    $script:DisallowConnectorIntersections = [bool](Get-JsonValue -Object $layoutSpec -Name "disallowConnectorIntersections" -Default $true)
    $script:ConnectorPortStub = [double](Get-JsonValue -Object $layoutSpec -Name "connectorPortStub" -Default 0.28)
    $script:KeepConnectorsOutsideModules = [bool](Get-JsonValue -Object $layoutSpec -Name "keepConnectorsOutsideModules" -Default $false)
    $script:TargetTerminalArrowOutsideModule = [bool](Get-JsonValue -Object $layoutSpec -Name "targetTerminalArrowOutsideModule" -Default $false)
    $script:TargetTerminalArrowVisualClearance = [double](Get-JsonValue -Object $layoutSpec -Name "targetTerminalArrowVisualClearance" -Default 0.03)
    $script:TargetTerminalArrowSize = [double](Get-JsonValue -Object $layoutSpec -Name "targetTerminalArrowSize" -Default 1.0)
    $script:LayerConnectorsUnderModules = [bool](Get-JsonValue -Object $layoutSpec -Name "layerConnectorsUnderModules" -Default $true)
    $script:AdaptiveFontScale = [double](Get-JsonValue -Object $layoutSpec -Name "adaptiveFontScale" -Default 1.0)
    $defaultMinFontSize = 5.5
    $defaultMaxFontSize = 14.0
    if ($script:ReferenceStyle) {
        $defaultMinFontSize = 7.5
        $defaultMaxFontSize = 12.5
    }
    $script:MinFontSize = [double](Get-JsonValue -Object $layoutSpec -Name "minFontSize" -Default $defaultMinFontSize)
    $script:MaxFontSize = [double](Get-JsonValue -Object $layoutSpec -Name "maxFontSize" -Default $defaultMaxFontSize)
    $script:PlacedRects = @()
    $script:RouteUsage = @{}
    $script:ConnectorInterfaceUsage = @{}
    $script:ConnectorSegmentUsage = @{}
    $script:NodeInfoById = @{}
    $currentWidth = [double]$page.PageSheet.CellsU("PageWidth").ResultIU
    $currentHeight = [double]$page.PageSheet.CellsU("PageHeight").ResultIU
    $canvasChanged = ([Math]::Abs($currentWidth - $pageWidth) -gt 0.01) -or ([Math]::Abs($currentHeight - $pageHeight) -gt 0.01)
    if ($UpdateExisting -and $canvasChanged -and $script:UpdateIdSet.Count -gt 0) {
        throw "Canvas size changed. Partial update was refused; run without -UpdateIds for a full redraw."
    }
    $page.PageSheet.CellsU("PageWidth").FormulaU = "$pageWidth in"
    $page.PageSheet.CellsU("PageHeight").FormulaU = "$pageHeight in"
    $title = Get-JsonValue -Object $spec -Name "title"
    if ($null -ne $title) { $page.Name = [string]$title }

    if ($UpdateExisting -and $script:UpdateIdSet.Count -eq 0) {
        Clear-GeneratedContent -Page $page -ClearAll $script:ClearPageOnFullRedraw
    }

    if (Should-Update -Id "note") {
        Remove-ShapesByPrefix -Page $page -Prefix "visio_auto_note"
        [void](Add-RectText -Page $page -X ($baseWidth / 2.0) -Y ($baseHeight - 0.55) -W ($baseWidth - 0.8) -H 0.55 -Text ([string](Get-JsonValue -Object $spec -Name "note" -Default "")) -Fill "#FFFFFF" -Line "#333333" -FontSize 11 -Name "visio_auto_note")
        Wait-Step
    }

    $headerY = $baseHeight - 1.2
    foreach ($lane in @($script:Lanes)) {
        $laneId = [string](Get-JsonValue -Object $lane -Name "id")
        if (-not (Should-Update -Id "lane:$laneId")) { continue }
        Remove-ShapesByPrefix -Page $page -Prefix "visio_auto_lane_$(Get-SafeName -Value $laneId)"
        $x1 = [double](Get-JsonValue -Object $lane -Name "x1")
        $x2 = [double](Get-JsonValue -Object $lane -Name "x2")
        $cx = ($x1 + $x2) / 2.0
        $w = $x2 - $x1
        [void](Add-RectText -Page $page -X $cx -Y $headerY -W $w -H 0.55 -Text ([string](Get-JsonValue -Object $lane -Name "title")) -Fill ([string](Get-JsonValue -Object $lane -Name "fill")) -Line "#333333" -FontSize ([double](Get-JsonValue -Object $lane -Name "fontSize" -Default 10.5)) -Name "visio_auto_lane_$(Get-SafeName -Value $laneId)_header")
        $sep = $page.DrawLine((Scale-X -Value $x1), (Scale-Y -Value 0.3), (Scale-X -Value $x1), (Scale-Y -Value ($baseHeight - 0.8)))
        Set-ShapeNameSafe -Shape $sep -Name "visio_auto_lane_$(Get-SafeName -Value $laneId)_separator"
        Set-CellFormulaSafe -Shape $sep -Cell "LineWeight" -Formula "1.5 pt"
        Wait-Step
    }
    $lastLane = @($script:Lanes)[-1]
    $rightX = [double](Get-JsonValue -Object $lastLane -Name "x2")
    if (Should-Update -Id "lane:right") {
        Remove-ShapesByPrefix -Page $page -Prefix "visio_auto_lane_right"
        $sepRight = $page.DrawLine((Scale-X -Value $rightX), (Scale-Y -Value 0.3), (Scale-X -Value $rightX), (Scale-Y -Value ($baseHeight - 0.8)))
        Set-ShapeNameSafe -Shape $sepRight -Name "visio_auto_lane_right_separator"
        Set-CellFormulaSafe -Shape $sepRight -Cell "LineWeight" -Formula "1.5 pt"
    }

    $shapeById = @{}
    $nodeInfoById = @{}
    foreach ($node in @((Get-JsonValue -Object $spec -Name "nodes"))) {
        if ($script:AutoAvoidOverlap) {
            $node = Resolve-NodePlacement -Node $node
        }
        $nodeId = [string](Get-JsonValue -Object $node -Name "id")
        if (-not (Should-Update -Id "node:$nodeId")) {
            $nodeInfoById[$nodeId] = Get-NodeInfo -Node $node
            continue
        }
        Remove-ShapesByPrefix -Page $page -Prefix "visio_auto_node_$(Get-SafeName -Value $nodeId)"
        $shapeById[$nodeId] = Add-Node -Page $page -Node $node -Name "visio_auto_node_$(Get-SafeName -Value $nodeId)"
        $nodeInfoById[$nodeId] = Get-NodeInfo -Node $node
        Wait-Step
    }
    $script:NodeInfoById = $nodeInfoById

    $ruleIndex = 0
    foreach ($rule in @((Get-JsonValue -Object $spec -Name "rules"))) {
        if (-not (Should-Update -Id "rule:$ruleIndex")) {
            $ruleIndex++
            continue
        }
        Remove-ShapesByPrefix -Page $page -Prefix "visio_auto_rule_$ruleIndex"
        $text = ([string](Get-JsonValue -Object $rule -Name "title")) + "`n" + ([string](Get-JsonValue -Object $rule -Name "text"))
        $ruleShape = Add-RectText -Page $page -X ([double](Get-JsonValue -Object $rule -Name "x")) -Y ([double](Get-JsonValue -Object $rule -Name "y")) -W ([double](Get-JsonValue -Object $rule -Name "w")) -H ([double](Get-JsonValue -Object $rule -Name "h")) -Text $text -Fill "#FFFFFF" -Line "#333333" -FontSize ([double](Get-JsonValue -Object $rule -Name "fontSize" -Default 8.5)) -Name "visio_auto_rule_$ruleIndex"
        Set-CellFormulaSafe -Shape $ruleShape -Cell "Para.HorzAlign" -Formula "0"
        Set-CellFormulaSafe -Shape $ruleShape -Cell "VerticalAlign" -Formula "0"
        Set-CellFormulaSafe -Shape $ruleShape -Cell "TxtMarginLeft" -Formula "0.08 in"
        Set-CellFormulaSafe -Shape $ruleShape -Cell "TxtMarginRight" -Formula "0.08 in"
        Set-CellFormulaSafe -Shape $ruleShape -Cell "TxtMarginTop" -Formula "0.06 in"
        Wait-Step
        $ruleIndex++
    }

    $edgeIndex = 0
    foreach ($edge in @((Get-JsonValue -Object $spec -Name "edges"))) {
        $fromId = [string](Get-JsonValue -Object $edge -Name "from")
        $toId = [string](Get-JsonValue -Object $edge -Name "to")
        $edgeShouldUpdate = ($script:UpdateIdSet.Count -eq 0) -or $script:UpdateIdSet.ContainsKey("edge:$edgeIndex") -or $script:UpdateIdSet.ContainsKey("node:$fromId") -or $script:UpdateIdSet.ContainsKey("node:$toId")
        if (-not $edgeShouldUpdate) {
            $edgeIndex++
            continue
        }
        Remove-ShapesByPrefix -Page $page -Prefix "visio_auto_edge_$edgeIndex"
        if ($nodeInfoById.ContainsKey($fromId) -and $nodeInfoById.ContainsKey($toId)) {
            [void](Add-Arrow -Page $page -FromInfo $nodeInfoById[$fromId] -ToInfo $nodeInfoById[$toId] -Edge $edge -NamePrefix "visio_auto_edge_$edgeIndex" -FromId $fromId -ToId $toId -EdgeIndex $edgeIndex)
            Wait-Step
        }
        $edgeIndex++
    }

    if ($script:LayerConnectorsUnderModules) {
        Bring-ShapesByPrefixToFront -Page $page -Prefix "visio_auto_node_"
        Bring-ShapesByPrefixToFront -Page $page -Prefix "visio_auto_rule_"
        Bring-ShapesByPrefixToFront -Page $page -Prefix "visio_auto_note"
    }

    $doc.SaveAs($fullOutputPath)
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

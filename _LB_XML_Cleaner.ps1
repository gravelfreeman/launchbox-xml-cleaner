[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$XmlPath,

    [Parameter()]
    [string]$RomsPath = (Get-Location).Path,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string[]]$RomExtensions = @('.zip', '.7z', '.chd'),

    [switch]$InPlace,

    [switch]$PreviewOnly
)

# Command-line usage:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -PreviewOnly -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -OutputPath "C:\Path\Cleaned.xml" -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -InPlace -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
#
# Flags:
#   -XmlPath        LaunchBox XML file to clean. Can also be passed as the first positional argument.
#   -RomsPath       Folder containing the ROMs/CHDs to scan.
#   -PreviewOnly    Simulate the cleanup without writing a file.
#   -OutputPath     Write to an explicit output path instead of replacing the source XML.
#   -InPlace        Replace the source XML. If -OutputPath is not provided, the default behavior already does a backup + replace.
#   -RomExtensions  File extensions treated as ROMs. Default: .zip, .7z, .chd

$ErrorActionPreference = 'Stop'
$progressId = 1

function Show-StageProgress {
    param(
        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [int]$Percent,

        [string]$CurrentOperation = ''
    )

    Write-Progress -Id $script:progressId -Activity 'Cleaning LaunchBox Platform XML' -Status $Status -PercentComplete $Percent -CurrentOperation $CurrentOperation
}

function Show-LoopProgress {
    param(
        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [int]$Current,

        [Parameter(Mandatory)]
        [int]$Total,

        [Parameter(Mandatory)]
        [int]$BasePercent,

        [Parameter(Mandatory)]
        [int]$SpanPercent,

        [string]$CurrentOperation = ''
    )

    if ($Total -le 0) {
        return
    }

    $step = [Math]::Max([int]($Total / 100), 1)
    if ($Current -ne 1 -and $Current -ne $Total -and ($Current % $step) -ne 0) {
        return
    }

    $percent = $BasePercent + [int][Math]::Floor(($Current / [double]$Total) * $SpanPercent)
    Show-StageProgress -Status $Status -Percent $percent -CurrentOperation $CurrentOperation
}

function Complete-StageProgress {
    Write-Progress -Id $script:progressId -Activity 'Cleaning LaunchBox Platform XML' -Completed
}

function New-CaseInsensitiveMap {
    return @{}
}

function Add-SetValue {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Map,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        $Map[$Value] = $true
    }
}

function Test-SetValue {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Map,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    return $Map.ContainsKey($Value)
}

function Get-ChildText {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlNode]$Parent,

        [Parameter(Mandatory)]
        [string[]]$Names
    )

    if ($null -eq $Parent) {
        return $null
    }

    foreach ($child in $Parent.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) {
            continue
        }

        foreach ($name in $Names) {
            if ($child.Name -ieq $name) {
                return [string]$child.InnerText
            }
        }
    }

    return $null
}

function ConvertTo-RomKey {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $text = $Value.Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $firstToken = ($text -split '\s+', 2 | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($firstToken)) {
        return $null
    }

    $candidate = $firstToken.Trim().Trim('"').TrimEnd('\', '/')
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return $null
    }

    $leaf = $candidate
    if ($candidate.IndexOf('\') -ge 0 -or $candidate.IndexOf('/') -ge 0) {
        $leaf = Split-Path -Path $candidate -Leaf
    }

    if ([string]::IsNullOrWhiteSpace($leaf)) {
        return $null
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        return $leaf
    }

    return $baseName
}

function Get-GameRomKey {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlNode]$GameNode
    )

    foreach ($fieldName in @('ApplicationPath', 'CommandLine', 'ApplicationCommandLine')) {
        $value = Get-ChildText -Parent $GameNode -Names @($fieldName)
        $romKey = ConvertTo-RomKey -Value $value
        if (-not [string]::IsNullOrWhiteSpace($romKey)) {
            return $romKey
        }
    }

    return $null
}

function Add-RomEvidence {
    param(
        [Parameter(Mandatory)]
        [hashtable]$GameRomEvidence,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$GameId,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$RomKey
    )

    if ([string]::IsNullOrWhiteSpace($GameId) -or [string]::IsNullOrWhiteSpace($RomKey)) {
        return
    }

    if (-not $GameRomEvidence.ContainsKey($GameId)) {
        $GameRomEvidence[$GameId] = New-CaseInsensitiveMap
    }

    Add-SetValue -Map $GameRomEvidence[$GameId] -Value $RomKey
}

function Test-GameHasLocalRom {
    param(
        [Parameter(Mandatory)]
        [hashtable]$LocalRomKeys,

        [AllowNull()]
        [hashtable]$RomEvidence
    )

    if ($null -eq $RomEvidence) {
        return $false
    }

    foreach ($romKey in $RomEvidence.Keys) {
        if (Test-SetValue -Map $LocalRomKeys -Value $romKey) {
            return $true
        }
    }

    return $false
}

function Save-FormattedXml {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlDocument]$Document,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.IndentChars = '  '
    $settings.NewLineChars = "`r`n"
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Replace
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)

    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try {
        $Document.Save($writer)
    }
    finally {
        $writer.Dispose()
    }
}

Show-StageProgress -Status 'Resolving paths' -Percent 2 -CurrentOperation 'Checking XML and ROM folders'
if (-not [string]::IsNullOrWhiteSpace($XmlPath)) {
    $resolvedXmlPath = (Resolve-Path -LiteralPath $XmlPath).Path
}
else {
    $xmlCandidates = @(Get-ChildItem -LiteralPath (Get-Location).Path -Filter '*.xml' -File)
    if ($xmlCandidates.Count -eq 1) {
        $resolvedXmlPath = $xmlCandidates[0].FullName
    }
    elseif ($xmlCandidates.Count -eq 0) {
        throw 'No XML file was provided and no .xml file was found in the current folder.'
    }
    else {
        throw 'No XML file was provided and multiple .xml files were found in the current folder. Pass -XmlPath or drag and drop the XML onto the launcher.'
    }
}
$resolvedRomsPath = (Resolve-Path -LiteralPath $RomsPath).Path

if ($InPlace -and $PSBoundParameters.ContainsKey('OutputPath')) {
    throw 'Use either -InPlace or -OutputPath, not both.'
}

if ($PreviewOnly -and $InPlace) {
    throw 'Use either -PreviewOnly or -InPlace, not both.'
}

$xmlItem = Get-Item -LiteralPath $resolvedXmlPath
$replaceOriginal = $InPlace -or -not $PSBoundParameters.ContainsKey('OutputPath')

if ($replaceOriginal) {
    $destinationPath = $resolvedXmlPath
    $backupPath = Join-Path -Path $xmlItem.DirectoryName -ChildPath ($xmlItem.BaseName + '.bak')
    $stagingPath = Join-Path -Path $xmlItem.DirectoryName -ChildPath ('{0}.new{1}' -f $xmlItem.BaseName, $xmlItem.Extension)
}
else {
    $destinationPath = [System.IO.Path]::GetFullPath($OutputPath)
    $backupPath = $null
    $stagingPath = $destinationPath
}

Show-StageProgress -Status 'Preparing ROM filters' -Percent 8 -CurrentOperation 'Building extension list'
$allowedExtensions = New-CaseInsensitiveMap
foreach ($extension in $RomExtensions) {
    if ([string]::IsNullOrWhiteSpace($extension)) {
        continue
    }

    if ($extension[0] -eq '.') {
        Add-SetValue -Map $allowedExtensions -Value $extension
    }
    else {
        Add-SetValue -Map $allowedExtensions -Value ('.' + $extension)
    }
}

Show-StageProgress -Status 'Scanning ROM folder' -Percent 12 -CurrentOperation $resolvedRomsPath
$localRomKeys = New-CaseInsensitiveMap
$romItems = @(Get-ChildItem -LiteralPath $resolvedRomsPath -Force)
for ($i = 0; $i -lt $romItems.Count; $i++) {
    $item = $romItems[$i]
    Show-LoopProgress -Status 'Scanning ROM folder' -Current ($i + 1) -Total $romItems.Count -BasePercent 12 -SpanPercent 10 -CurrentOperation $item.Name

    if ($item.PSIsContainer) {
        Add-SetValue -Map $localRomKeys -Value $item.Name
        continue
    }

    if (Test-SetValue -Map $allowedExtensions -Value $item.Extension) {
        Add-SetValue -Map $localRomKeys -Value ([System.IO.Path]::GetFileNameWithoutExtension($item.Name))
    }
}

Show-StageProgress -Status 'Loading XML' -Percent 24 -CurrentOperation $resolvedXmlPath
$xml = [System.Xml.XmlDocument]::new()
$xml.PreserveWhitespace = $false
$xml.Load($resolvedXmlPath)

$root = $xml.DocumentElement
if ($null -eq $root -or $root.Name -ne 'LaunchBox') {
    throw "Expected a <LaunchBox> root node in '$resolvedXmlPath'."
}

Show-StageProgress -Status 'Indexing additional applications' -Percent 30 -CurrentOperation 'Collecting ROM evidence by GameID'
$gameRomEvidence = New-CaseInsensitiveMap
$additionalApplicationNodes = @($root.SelectNodes('./AdditionalApplication'))
for ($i = 0; $i -lt $additionalApplicationNodes.Count; $i++) {
    $additionalApplicationNode = $additionalApplicationNodes[$i]
    $gameId = Get-ChildText -Parent $additionalApplicationNode -Names @('GameID', 'GameId')
    $romKey = Get-GameRomKey -GameNode $additionalApplicationNode
    Add-RomEvidence -GameRomEvidence $gameRomEvidence -GameId $gameId -RomKey $romKey
    Show-LoopProgress -Status 'Indexing additional applications' -Current ($i + 1) -Total $additionalApplicationNodes.Count -BasePercent 30 -SpanPercent 15 -CurrentOperation $romKey
}

Show-StageProgress -Status 'Evaluating main games' -Percent 46 -CurrentOperation 'Checking local ROM evidence'
$gameNodes = @($root.SelectNodes('./Game'))
$keptGameIds = New-CaseInsensitiveMap
$gamesToRemove = New-Object System.Collections.ArrayList
$gamesMissingRomKey = 0
$gamesKeptByAdditionalApplication = 0

for ($i = 0; $i -lt $gameNodes.Count; $i++) {
    $gameNode = $gameNodes[$i]
    $gameId = Get-ChildText -Parent $gameNode -Names @('ID', 'Id')
    $mainRomKey = Get-GameRomKey -GameNode $gameNode

    if (-not [string]::IsNullOrWhiteSpace($mainRomKey)) {
        Add-RomEvidence -GameRomEvidence $gameRomEvidence -GameId $gameId -RomKey $mainRomKey
    }
    elseif ([string]::IsNullOrWhiteSpace($gameId) -or -not $gameRomEvidence.ContainsKey($gameId)) {
        $gamesMissingRomKey++
        [void]$gamesToRemove.Add($gameNode)
        Show-LoopProgress -Status 'Evaluating main games' -Current ($i + 1) -Total $gameNodes.Count -BasePercent 46 -SpanPercent 22 -CurrentOperation 'Missing ROM evidence'
        continue
    }

    $evidence = $null
    if (-not [string]::IsNullOrWhiteSpace($gameId) -and $gameRomEvidence.ContainsKey($gameId)) {
        $evidence = $gameRomEvidence[$gameId]
    }

    if (Test-GameHasLocalRom -LocalRomKeys $localRomKeys -RomEvidence $evidence) {
        Add-SetValue -Map $keptGameIds -Value $gameId
        if (-not [string]::IsNullOrWhiteSpace($gameId) -and -not [string]::IsNullOrWhiteSpace($mainRomKey) -and -not (Test-SetValue -Map $localRomKeys -Value $mainRomKey)) {
            $gamesKeptByAdditionalApplication++
        }
        Show-LoopProgress -Status 'Evaluating main games' -Current ($i + 1) -Total $gameNodes.Count -BasePercent 46 -SpanPercent 22 -CurrentOperation $mainRomKey
        continue
    }

    [void]$gamesToRemove.Add($gameNode)
    Show-LoopProgress -Status 'Evaluating main games' -Current ($i + 1) -Total $gameNodes.Count -BasePercent 46 -SpanPercent 22 -CurrentOperation $mainRomKey
}

Show-StageProgress -Status 'Removing discarded games' -Percent 69 -CurrentOperation "$($gamesToRemove.Count) entries"
for ($i = 0; $i -lt $gamesToRemove.Count; $i++) {
    [void]$root.RemoveChild($gamesToRemove[$i])
    Show-LoopProgress -Status 'Removing discarded games' -Current ($i + 1) -Total $gamesToRemove.Count -BasePercent 69 -SpanPercent 6 -CurrentOperation ($i + 1).ToString()
}

$secondaryRemovalCounts = [ordered]@{
    AdditionalApplication = 0
    GameControllerSupport = 0
    AlternateName = 0
}

$secondaryStageMap = @{
    AdditionalApplication = @{ Base = 76; Span = 6 }
    GameControllerSupport = @{ Base = 82; Span = 6 }
    AlternateName = @{ Base = 88; Span = 6 }
}

foreach ($nodeName in @('AdditionalApplication', 'GameControllerSupport', 'AlternateName')) {
    $secondaryNodes = @($root.SelectNodes("./$nodeName"))
    $stage = $secondaryStageMap[$nodeName]
    Show-StageProgress -Status "Pruning $nodeName" -Percent $stage.Base -CurrentOperation "$($secondaryNodes.Count) entries"

    for ($i = 0; $i -lt $secondaryNodes.Count; $i++) {
        $secondaryNode = $secondaryNodes[$i]
        $gameId = Get-ChildText -Parent $secondaryNode -Names @('GameID', 'GameId')
        if (-not (Test-SetValue -Map $keptGameIds -Value $gameId)) {
            [void]$root.RemoveChild($secondaryNode)
            $secondaryRemovalCounts[$nodeName] = [int]$secondaryRemovalCounts[$nodeName] + 1
        }

        Show-LoopProgress -Status "Pruning $nodeName" -Current ($i + 1) -Total $secondaryNodes.Count -BasePercent $stage.Base -SpanPercent $stage.Span -CurrentOperation ($i + 1).ToString()
    }
}

if (-not $PreviewOnly) {
    Show-StageProgress -Status 'Writing XML' -Percent 96 -CurrentOperation $destinationPath
    if ($replaceOriginal) {
        if (Test-Path -LiteralPath $stagingPath) {
            Remove-Item -LiteralPath $stagingPath -Force
        }

        Save-FormattedXml -Document $xml -Path $stagingPath

        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Force
        }

        if (Test-Path -LiteralPath $destinationPath) {
            Move-Item -LiteralPath $destinationPath -Destination $backupPath
        }

        try {
            Move-Item -LiteralPath $stagingPath -Destination $destinationPath
        }
        catch {
            if (-not (Test-Path -LiteralPath $destinationPath) -and (Test-Path -LiteralPath $backupPath)) {
                Move-Item -LiteralPath $backupPath -Destination $destinationPath
            }
            throw
        }

        Write-Host "Backup created: $backupPath"
    }
    else {
        Save-FormattedXml -Document $xml -Path $destinationPath
    }
}

Show-StageProgress -Status 'Done' -Percent 100 -CurrentOperation 'Preparing summary'
Complete-StageProgress

Write-Host "Scanned ROM folder: $resolvedRomsPath"
Write-Host "Detected local ROM keys: $($localRomKeys.Count)"
Write-Host "Main <Game> entries kept: $($gameNodes.Count - $gamesToRemove.Count)"
Write-Host "Main <Game> entries removed: $($gamesToRemove.Count)"
Write-Host "Main <Game> entries removed due to missing ROM evidence: $gamesMissingRomKey"
Write-Host "Main <Game> entries kept because of local AdditionalApplication: $gamesKeptByAdditionalApplication"
Write-Host "Removed <AdditionalApplication>: $($secondaryRemovalCounts['AdditionalApplication'])"
Write-Host "Removed <GameControllerSupport>: $($secondaryRemovalCounts['GameControllerSupport'])"
Write-Host "Removed <AlternateName>: $($secondaryRemovalCounts['AlternateName'])"
if ($PreviewOnly) {
    Write-Host 'Preview only: no file written.'
}
else {
    Write-Host "Output path: $destinationPath"
}




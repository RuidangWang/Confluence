[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Path,

    [string] $Sheet,

    [int] $MaxRows = 0,

    [switch] $ListSheets
)

$ErrorActionPreference = 'Stop'

$resolvedPath = (Resolve-Path -LiteralPath $Path).Path
if ([System.IO.Path]::GetExtension($resolvedPath) -ne '.xlsx') {
    throw "Expected an .xlsx file: $resolvedPath"
}
if ($MaxRows -lt 0) {
    throw '-MaxRows must be zero or a positive number.'
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($resolvedPath)

function Read-WorkbookXml {
    param([Parameter(Mandatory = $true)][string] $EntryName)

    $entry = $script:archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        throw "The workbook is missing '$EntryName'."
    }

    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try {
            [xml] $document = $reader.ReadToEnd()
            return $document
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

try {
    $workbook = Read-WorkbookXml 'xl/workbook.xml'
    $relationships = Read-WorkbookXml 'xl/_rels/workbook.xml.rels'
    $targets = @{}
    foreach ($relationship in $relationships.SelectNodes("//*[local-name()='Relationship']")) {
        if ($relationship.GetAttribute('TargetMode') -ne 'External') {
            $targets[$relationship.GetAttribute('Id')] = $relationship.GetAttribute('Target')
        }
    }

    $sheets = @()
    foreach ($sheetNode in $workbook.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
        $relationshipId = $sheetNode.GetAttribute(
            'id',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        )
        $target = $targets[$relationshipId]
        if ([string]::IsNullOrWhiteSpace($target)) {
            throw "Could not resolve worksheet '$($sheetNode.GetAttribute('name'))'."
        }

        $target = $target.Replace('\', '/')
        if ($target.StartsWith('/')) {
            $entryName = $target.TrimStart('/')
        }
        elseif ($target.StartsWith('xl/')) {
            $entryName = $target
        }
        else {
            $entryName = "xl/$target"
        }

        $sheets += [pscustomobject]@{
            Name      = $sheetNode.GetAttribute('name')
            EntryName = $entryName
        }
    }

    if ($ListSheets) {
        foreach ($item in $sheets) {
            $document = Read-WorkbookXml $item.EntryName
            $dimension = $document.SelectSingleNode("//*[local-name()='dimension']")
            $range = if ($dimension) { $dimension.GetAttribute('ref') } else { 'unknown range' }
            Write-Output "$($item.Name)`t$range"
        }
        return
    }

    if ($Sheet) {
        $selectedSheets = @($sheets | Where-Object { $_.Name -eq $Sheet })
        if ($selectedSheets.Count -eq 0) {
            $names = ($sheets | ForEach-Object { $_.Name }) -join ', '
            throw "Worksheet '$Sheet' was not found. Available sheets: $names"
        }
    }
    else {
        $selectedSheets = $sheets
    }

    $sharedStrings = @()
    $sharedStringsEntry = $archive.GetEntry('xl/sharedStrings.xml')
    if ($null -ne $sharedStringsEntry) {
        $sharedStringsDocument = Read-WorkbookXml 'xl/sharedStrings.xml'
        foreach ($stringNode in $sharedStringsDocument.SelectNodes("//*[local-name()='si']")) {
            $text = ($stringNode.SelectNodes(".//*[local-name()='t']") |
                ForEach-Object { $_.InnerText }) -join ''
            $sharedStrings += ,$text
        }
    }

    foreach ($item in $selectedSheets) {
        $document = Read-WorkbookXml $item.EntryName
        Write-Output "## $($item.Name)"
        $printedRows = 0

        foreach ($row in $document.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $cells = @()
            foreach ($cell in $row.SelectNodes("./*[local-name()='c']")) {
                $type = $cell.GetAttribute('t')
                $valueNode = $cell.SelectSingleNode("./*[local-name()='v']")
                $value = ''

                if ($type -eq 's' -and $null -ne $valueNode) {
                    $value = $sharedStrings[[int] $valueNode.InnerText]
                }
                elseif ($type -eq 'inlineStr') {
                    $value = ($cell.SelectNodes(".//*[local-name()='t']") |
                        ForEach-Object { $_.InnerText }) -join ''
                }
                elseif ($null -ne $valueNode) {
                    $value = $valueNode.InnerText
                }
                elseif ($cell.SelectSingleNode("./*[local-name()='f']")) {
                    $value = '[formula has no cached value]'
                }

                if ($value -is [string]) {
                    $value = $value.Replace("`r", '\r').Replace("`n", '\n')
                }
                $cells += ('{0}={1}' -f $cell.GetAttribute('r'), $value)
            }

            if ($cells.Count -gt 0) {
                Write-Output ($cells -join ' | ')
                $printedRows++
                if ($MaxRows -gt 0 -and $printedRows -ge $MaxRows) {
                    break
                }
            }
        }
    }
}
finally {
    $archive.Dispose()
}

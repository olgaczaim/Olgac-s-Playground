#best way to get or delete is from basic url
#https://contonso.com/subsite/Lists/Ak%20Gemii/AllItems.aspx (for tr site)
#https://contonso.com/subsite/Lists/Workflow%20History/AllItems.aspx (for en site)
#but if you insist that you need to do it with powershell, here u asked for it.


Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# --- User-editable settings ---
$webUrl = "http://contonso.com/subsite"
#$listName = "Workflow History"
$listName = "İş Akışı Geçmişi"

# Provide a GUID (with or without braces) to search by WorkflowInstance, or set to $null to skip GUID search.
# from workflow history item url get the instance id ex:
#_layouts/15/WrkStat.aspx?List={A9C6F415-2D6E-4D6C-AFB0-6B1D37792A14}&WorkflowInstanceID={693640d2-9fdf-4d64-9ff4-8b6b8c223f58}
$targetInstanceId = "{3FEF74A3-7144-46EA-A1AC-8079AFC097F2}"

# Provide an ows_Item value (e.g. "3") to search by item id, or set to $null to skip ows_Item search.
$targetOwsItem = $null #"571"

# Set $exportCsvPath to a full path to export CSV, e.g. "C:\wsp\wf_history_ows_export.csv"
# Leave as $null or "" to skip CSV export.
$exportCsvPath = $null #"C:\exports\wf_history_ows_export.csv"

# If $true, script prints detailed readable dumps for each matched item
$parseDetailed = $null
# ------------------------------

function Normalize-OwsValue {
    param($raw)
    if ($null -eq $raw) { return $null }
    $v = $raw -replace '^\d+;#', ''
    $v = $v -replace '^\{(.*)\}$','$1'
    return $v
}

function Parse-OwsAttributesFromXml {
    param($xmlString)
    $ret = @{}
    if (-not $xmlString) { return $ret }
    $pattern = "ows_([^=]+)='([^']*)'"
    [regex]::Matches($xmlString, $pattern) | ForEach-Object {
        $name = $_.Groups[1].Value
        $val  = $_.Groups[2].Value
        $ret[$name] = Normalize-OwsValue $val
    }
    return $ret
}

# Prepare GUID search only if provided
$guidObj = $null
if ($targetInstanceId -and $targetInstanceId.ToString().Trim() -ne "") {
    $cleanGuid = $targetInstanceId.ToString().Trim()
    if ($cleanGuid.StartsWith("{") -and $cleanGuid.EndsWith("}")) {
        $cleanGuid = $cleanGuid.Substring(1, $cleanGuid.Length - 2)
    }
    try {
        $guidObj = [guid]::Parse($cleanGuid)
    } catch {
        Write-Warning "Provided targetInstanceId is not a valid GUID; GUID search will be skipped."
        $guidObj = $null
    }
}

# Normalize ows_Item if provided (string compare)
if ($targetOwsItem -ne $null) {
    $targetOwsItem = $targetOwsItem.ToString().Trim()
    if ($targetOwsItem -eq "") { $targetOwsItem = $null }
}

if (($guidObj -eq $null) -and ($targetOwsItem -eq $null)) {
    Write-Error "No search criteria provided. Set either `$targetInstanceId` or `$targetOwsItem` (or both)."
    return
}

Write-Host "Searching list '$listName' in web '$webUrl'..." -ForegroundColor Cyan
if ($guidObj) { Write-Host " - WorkflowInstance GUID: $guidObj" -ForegroundColor Cyan }
if ($targetOwsItem) { Write-Host " - ows_Item value    : '$targetOwsItem'" -ForegroundColor Cyan }

$spweb = Get-SPWeb $webUrl
if ($null -eq $spweb) {
    Write-Error "Unable to open web: $webUrl"
    return
}

try {
    $list = $spweb.Lists[$listName]
    if ($null -eq $list) {
        Write-Error "List '$listName' not found in web $webUrl"
        return
    }

    $matches = @()
    $allOwsKeys = New-Object System.Collections.Generic.HashSet[string]

    Write-Host "Scanning items in list '$listName'..." -ForegroundColor DarkCyan

    foreach ($it in $list.Items) {
        $xml = $it.Xml
        if (-not $xml) { continue }

        # Quick check: only parse XML if it can possibly match at least one active criterion
        $xmlLower = $xml.ToLower()
        $shouldParse = $false
        if ($guidObj -ne $null -and $xmlLower.Contains($guidObj.ToString().ToLower())) { $shouldParse = $true }
        if ($targetOwsItem -ne $null -and $xmlLower.Contains("ows_item='" + $targetOwsItem.ToLower() + "'")) { $shouldParse = $true }
        if (-not $shouldParse) { continue }

        $attrs = Parse-OwsAttributesFromXml $xml

        # record keys for CSV
        foreach ($k in $attrs.Keys) { [void]$allOwsKeys.Add($k) }

        # Evaluate matches according to enabled criteria
        $matchedByGuid = $false
        $matchedByOwsItem = $false

        if ($guidObj -ne $null) {
            if ($attrs.ContainsKey("WorkflowInstance")) {
                try {
                    if ($attrs["WorkflowInstance"] -and ([guid]::Parse($attrs["WorkflowInstance"]) -eq $guidObj)) {
                        $matchedByGuid = $true
                    }
                } catch {
                    # fallback to xml text check
                    if ($xmlLower.Contains($guidObj.ToString().ToLower())) { $matchedByGuid = $true }
                }
            } else {
                if ($xmlLower.Contains($guidObj.ToString().ToLower())) { $matchedByGuid = $true }
            }
        }

        if ($targetOwsItem -ne $null) {
            if ($attrs.ContainsKey("Item")) {
                if ($attrs["Item"] -and $attrs["Item"].ToString().Trim() -eq $targetOwsItem) {
                    $matchedByOwsItem = $true
                }
            } else {
                if ($xmlLower.Contains("ows_item='" + $targetOwsItem.ToLower() + "'")) { $matchedByOwsItem = $true }
            }
        }

        # only include items that matched at least one enabled criterion
        if (-not ($matchedByGuid -or $matchedByOwsItem)) { continue }

        # collect display fields
        $wfInstance = $null
        if ($it.Fields.ContainsField("WorkflowInstance")) { $wfInstance = $it["WorkflowInstance"] }
        elseif ($attrs.ContainsKey("WorkflowInstance")) { $wfInstance = $attrs["WorkflowInstance"] }

        $occurred = $null
        if ($it.Fields.ContainsField("Occurred")) { $occurred = $it["Occurred"] }
        elseif ($attrs.ContainsKey("Occurred")) { $occurred = $attrs["Occurred"] }

        $user = $null
        if ($it.Fields.ContainsField("User")) { $user = $it["User"] }
        elseif ($attrs.ContainsKey("User")) { $user = $attrs["User"] }

        $description = $null
        if ($it.Fields.ContainsField("Description")) { $description = $it["Description"] }
        elseif ($attrs.ContainsKey("Description")) { $description = $attrs["Description"] }

        $uniqueIdVal = $null
        if ($it.UniqueId) { $uniqueIdVal = $it.UniqueId.ToString() }
        elseif ($attrs.ContainsKey("GUID")) { $uniqueIdVal = $attrs["GUID"] }
        elseif ($attrs.ContainsKey("UniqueId")) { $uniqueIdVal = $attrs["UniqueId"] }

        $rawSnippet = $xml
        if ($rawSnippet.Length -gt 800) { $rawSnippet = $rawSnippet.Substring(0,800) + " ... [truncated]" }

        $matches += [PSCustomObject]@{
            ItemObject      = $it
            ID              = $it.ID
            UniqueId        = $uniqueIdVal
            WorkflowInstance= $wfInstance
            Occurred        = $occurred
            User            = $user
            Description     = $description
            OwsAttributes   = $attrs
            RawXmlSnippet   = $rawSnippet
            MatchedByGuid   = $matchedByGuid
            MatchedByItem   = $matchedByOwsItem
        }
    }

    if ($matches.Count -eq 0) {
        Write-Warning "No matches found for given criteria in list '$listName'."
        return
    }

    Write-Host "`nFound $($matches.Count) matching item(s)." -ForegroundColor Green

    # Overview table
    $overview = $matches | Select-Object ID, UniqueId, WorkflowInstance, Occurred, User, @{Name='MatchedBy';Expression={ if ($_.MatchedByGuid -and $_.MatchedByItem) { 'GUID+Item' } elseif ($_.MatchedByGuid) { 'GUID' } else { 'Item' } } }, `
                       @{Name='Description';Expression={ if ($_.Description) { $d = $_.Description.ToString(); if ($d.Length -gt 80) { $d.Substring(0,77) + '...' } else { $d } } else { $null } } }

    $overview | Format-Table -AutoSize

    # Detailed readable dump (optional)
    if ($parseDetailed) {
        foreach ($m in $matches) {
            Write-Host "`n`n=== Item ID: $($m.ID)   UniqueId: $($m.UniqueId) ===" -ForegroundColor Yellow
            Write-Host ("Matched by       : {0}" -f (if ($m.MatchedByGuid -and $m.MatchedByItem) { 'GUID + ows_Item' } elseif ($m.MatchedByGuid) { 'GUID' } else { 'ows_Item' } ))
            Write-Host ("Occurred         : {0}" -f (if ($m.Occurred -ne $null) { $m.Occurred } else { "<not present>" }))
            Write-Host ("WorkflowInstance : {0}" -f (if ($m.WorkflowInstance -ne $null) { $m.WorkflowInstance } elseif ($m.OwsAttributes.ContainsKey("WorkflowInstance")) { $m.OwsAttributes["WorkflowInstance"] } else { "<not present>" }))
            Write-Host ("User             : {0}" -f (if ($m.User -ne $null) { $m.User } else { "<not present>" }))
            Write-Host ("Description      : {0}" -f (if ($m.Description -ne $null) { $m.Description } else { "<not present>" }))
            Write-Host ("Raw XML snippet  :")
            Write-Host $m.RawXmlSnippet -ForegroundColor DarkGray

            Write-Host "`nParsed ows_ attributes (readable):" -ForegroundColor Cyan
            $enumerator = $m.OwsAttributes.GetEnumerator() | Sort-Object Name
            foreach ($entry in $enumerator) {
                $left = $entry.Name
                $val = $entry.Value
                if ($val -ne $null) { $displayVal = $val } else { $displayVal = "<empty>" }
                if ($displayVal -ne $null -and $displayVal.ToString().Length -gt 200) { $displayVal = $displayVal.ToString().Substring(0,197) + "..." }
                Write-Host ("  {0,-30}: {1}" -f $left, $displayVal)
            }
            Write-Host "----------------------------------------------------------------" -ForegroundColor Gray
        }
    }

    # CSV export (if requested)
    if ($exportCsvPath -and $exportCsvPath.Trim() -ne "") {
        Write-Host "`nPreparing CSV export to: $exportCsvPath" -ForegroundColor Cyan

        $owsCols = @()
        if ($allOwsKeys.Count -gt 0) { $owsCols = @($allOwsKeys) | Sort-Object }

        $standardCols = @("ItemID","UniqueId","WorkflowInstance","Occurred","User","Description","RawXmlSnippet","MatchedBy")
        $allCols = $standardCols + $owsCols

        $exportObjects = @()
        foreach ($m in $matches) {
            $row = [ordered]@{
                ItemID = $m.ID
                UniqueId = $m.UniqueId
                WorkflowInstance = $m.WorkflowInstance
                Occurred = $m.Occurred
                User = $m.User
                Description = $m.Description
                RawXmlSnippet = $m.RawXmlSnippet
                MatchedBy = (if ($m.MatchedByGuid -and $m.MatchedByItem) { 'GUID+Item' } elseif ($m.MatchedByGuid) { 'GUID' } else { 'Item' })
            }

            foreach ($k in $owsCols) {
                if ($m.OwsAttributes.ContainsKey($k)) { $row[$k] = $m.OwsAttributes[$k] } else { $row[$k] = "" }
            }

            $exportObjects += New-Object PSObject -Property $row
        }

        try {
            $folder = Split-Path -Parent $exportCsvPath
            if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
            $exportObjects | Export-Csv -Path $exportCsvPath -NoTypeInformation -Encoding UTF8
            Write-Host "CSV exported to: $exportCsvPath" -ForegroundColor Green
            Write-Host "CSV columns: $($allCols -join ', ')" -ForegroundColor DarkCyan
        } catch {
            Write-Error "Failed to export CSV: $_"
        }
    }

} finally {
    if ($spweb -ne $null) { $spweb.Dispose() }
}
